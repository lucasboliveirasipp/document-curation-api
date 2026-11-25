"""
gateway_apim_service.py

Este arquivo é o "telefone" do agente com a Gateway: ele faz login
(Azure AD) e conversa com o workflow "Gerador de Texto" para pedir uma
resposta ao modelo de IA. É o único lugar do projeto que sabe como
autenticar e chamar a Gateway — o resto do código só usa
GatewayClient.gerar_texto() sem se preocupar com token, payload ou
polling de status.
"""

import threading
import time
from datetime import datetime, timedelta
import os

import jwt
import requests

from src.util import sanitize_header_value

WORKFLOW_CONFIGURATION_CODE = os.getenv("WORKFLOW_CONFIGURATION_CODE", "WF_GERADOR_DE_TEXTO_AGNT_ASSISTIDO")
OPENAI_API_ENDPOINT = os.getenv("OPENAI_API_ENDPOINT")
OPENAI_API_VERSION = os.getenv("OPENAI_API_VERSION")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "teste")

_STATUS_SUCESSO = "WF_COMPLETED_SUCCESS"
_STATUS_FALHA = "WF_COMPLETED_ERROR"
_POLL_INTERVALO_SEGUNDOS = int(os.getenv("GATEWAY_POLL_INTERVALO_SEGUNDOS", "2"))
# Documentos grandes (Excel com milhares de linhas, PDFs longos etc.) geram
# vários lotes avaliados em sequência — cada lote é 1 chamada ao workflow, e
# quanto mais complexo/denso o conteúdo do lote, mais tempo o modelo demora
# para responder. 90s era curto demais e derrubava o arquivo inteiro assim
# que 1 único lote (de dezenas) demorasse um pouco mais. Configurável via
# env var em vez de fixo no código.
_POLL_TIMEOUT_SEGUNDOS = int(os.getenv("GATEWAY_POLL_TIMEOUT_SEGUNDOS", "180"))


class GatewayClient:
    """Injetado via Depends(GatewayClient) no controller — 1 instância
    nova por requisição HTTP, então token/token_expiry não são
    compartilhados entre requisições concorrentes (cada uma autentica a
    sua). Dentro de uma mesma requisição, porém, a curadoria de
    documentos grandes chama gerar_texto() de várias threads em
    paralelo (ver avaliar_documento em curadoria_service.py) usando essa
    MESMA instância — daí o lock em get_token(): sem ele, se o token
    expirar no meio do processamento, várias threads poderiam disparar
    authenticate() ao mesmo tempo (chamadas redundantes à Gateway)."""

    def __init__(self):
        self.workflow_url = os.getenv("APIM_WORKFLOW_URL")
        self.status_url = os.getenv("APIM_STATUS_URL")
        self.token = None
        self.token_expiry = None
        self._lock_token = threading.Lock()

    def get_token(self):
        """Devolve um token válido, autenticando de novo se ainda não
        houver token ou se ele estiver a menos de 60s de expirar
        (margem de segurança para não expirar no meio de uma chamada
        em andamento à Gateway). Protegido por lock porque múltiplas
        threads (lotes em paralelo) podem chamar isso ao mesmo tempo."""
        with self._lock_token:
            if self.token is None or self.token_expiry is None or datetime.now() >= self.token_expiry - timedelta(seconds=60):
                self.authenticate()
            return self.token

    def authenticate(self):
        """Realiza autenticação segura e armazena o token sanitizado."""
        auth = self._authenticate_apim()

        if not auth:
            raise RuntimeError("Falha ao autenticar na Gateway")

        try:
            data = auth.json()
        except ValueError:
            raise RuntimeError("Resposta inválida da Gateway (não é JSON)")

        token = data.get("access_token")
        if not token:
            raise RuntimeError("Token ausente na resposta da Gateway")

        self.token = sanitize_header_value(token)

        try:
            decoded = jwt.decode(self.token, options={"verify_signature": False})
            self.token_expiry = datetime.fromtimestamp(decoded["exp"])
        except Exception as e:
            raise RuntimeError(f"Erro ao decodificar token JWT: {e}")

    def _authenticate_apim(self):
        """Autentica via client_credentials no Azure AD."""
        tenant_id = os.getenv("TENANTID")
        client_id = os.getenv("GATEWAYAPPID")
        scope_uri = os.getenv("GATEWAYAPPIDURI")

        token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        token_data = {
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": os.getenv("GATEWAYCLIENTSECRET"),
            "scope": f"{scope_uri}/.default",
        }

        return requests.post(token_url, data=token_data, verify=False)

    def _extrair_texto_gerado(self, resposta: dict) -> str:
        """Extrai o texto gerado da resposta do workflow. Caminho
        confirmado no teste isolado: output_collection.output_datas[0]
        .workflow_step_output_collection.output.json_data.choices[0]
        .message.content. Se vier diferente do esperado (ex: a Gateway
        alterar o formato do wrapper do workflow), cai numa busca
        recursiva por qualquer dict com {"role": ..., "content": <str>}
        em vez de falhar direto — resiliente a mudança de formato desde
        que a estrutura {role, content} continue presente em algum
        nível da resposta."""
        try:
            saidas = resposta["output_collection"]["output_datas"]
            json_data = saidas[0]["workflow_step_output_collection"]["output"]["json_data"]
            return json_data["choices"][0]["message"]["content"]
        except (KeyError, IndexError, TypeError):
            pass

        def buscar(valor):
            if isinstance(valor, dict):
                if "content" in valor and "role" in valor and isinstance(valor.get("content"), str):
                    return valor["content"]
                for sub_valor in valor.values():
                    achado = buscar(sub_valor)
                    if achado:
                        return achado
            elif isinstance(valor, list):
                for item in valor:
                    achado = buscar(item)
                    if achado:
                        return achado
            return None

        texto = buscar(resposta)
        if texto is None:
            raise RuntimeError(
                "Não foi possível localizar o texto gerado na resposta do "
                f"workflow. Resposta completa recebida: {resposta}"
            )
        return texto

    def gerar_texto(
        self,
        system_prompt: str,
        user_prompt: str,
        model: str,
        *,
        temperature: float = 0,
        max_tokens: int = 4096,
        response_format: dict | None = None,
        top_p: float = 0,
        seed: int = 0,
        frequency_penalty: float = 0,
        presence_penalty: float = 0,
    ) -> str:
        """Gera texto pelo workflow "Gerador de Texto" da Gateway: inicia
        o workflow (POST), faz polling do status a cada
        _POLL_INTERVALO_SEGUNDOS até concluir ou até
        _POLL_TIMEOUT_SEGUNDOS esgotar, e devolve o texto gerado. Payload
        no formato confirmado pelo teste isolado (inclui
        openai_api_endpoint/version/key além de gen_params, mesmo não
        documentados oficialmente). Levanta RuntimeError se o workflow
        não iniciar, falhar ou não responder dentro do timeout."""
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.get_token()}",
        }
        payload = {
            "workflow_configuration_code": WORKFLOW_CONFIGURATION_CODE,
            "input_collection": {
                "input_properties": [
                    {
                        "workflow_step_order": 0,
                        "workflow_step_properties": [
                            {"full_path": "openai_api_endpoint", "actual_value": OPENAI_API_ENDPOINT},
                            {"full_path": "openai_api_version", "actual_value": OPENAI_API_VERSION},
                            {"full_path": "openai_api_key", "actual_value": OPENAI_API_KEY},
                            {"full_path": "gen_params.model", "actual_value": model},
                            {
                                "full_path": "gen_params.messages",
                                "actual_value": [
                                    {"role": "system", "content": system_prompt},
                                    {"role": "user", "content": user_prompt},
                                ],
                            },
                            {"full_path": "gen_params.temperature", "actual_value": temperature},
                            {"full_path": "gen_params.max_tokens", "actual_value": max_tokens},
                            {"full_path": "gen_params.top_p", "actual_value": top_p},
                            {"full_path": "gen_params.seed", "actual_value": seed},
                            {"full_path": "gen_params.response_format", "actual_value": response_format or {"type": "text"}},
                            {"full_path": "gen_params.frequency_penalty", "actual_value": frequency_penalty},
                            {"full_path": "gen_params.presence_penalty", "actual_value": presence_penalty},
                        ],
                    }
                ]
            },
        }

        response = requests.post(
            f"{self.workflow_url}/request/start-workflow",
            json=payload, headers=headers, verify=False,
        )
        response.raise_for_status()
        corpo = response.json()

        execucao_id = corpo.get("workflow_execution_id")
        if not execucao_id:
            raise RuntimeError(f"Workflow não retornou workflow_execution_id. Resposta: {corpo}")

        inicio = time.monotonic()
        while True:
            time.sleep(_POLL_INTERVALO_SEGUNDOS)

            status_response = requests.get(
                f"{self.status_url}/response/workflow-execution-status/{execucao_id}",
                headers=headers, verify=False,
            )
            status_response.raise_for_status()
            status_corpo = status_response.json()
            status = str(status_corpo.get("status", "")).upper()

            if status == _STATUS_SUCESSO:
                return self._extrair_texto_gerado(status_corpo)

            if status == _STATUS_FALHA:
                raise RuntimeError(f"Workflow falhou (status={status}). Resposta: {status_corpo}")

            if time.monotonic() - inicio > _POLL_TIMEOUT_SEGUNDOS:
                raise RuntimeError(f"Tempo esgotado esperando o workflow terminar (status atual={status}).")
