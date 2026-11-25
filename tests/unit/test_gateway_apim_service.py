# tests/unit/test_gateway_apim_service.py
"""
Este arquivo testa o GatewayClient de verdade, e não um substituto —
simula as respostas da rede (requests) e do token (jwt) para conferir
se o login na Gateway, a geração de texto e o tratamento de erros
funcionam como esperado, sem precisar de internet nem de credenciais
reais.
"""
from datetime import datetime, timedelta

import pytest
import requests

import src.services.gateway_apim_service as svc_mod
from src.services.gateway_apim_service import GatewayClient


class DummyResponse:
    def __init__(self, status=200, payload=None, invalid_json=False):
        self._status = status
        self._payload = payload if payload is not None else {}
        self._invalid_json = invalid_json

    def json(self):
        if self._invalid_json:
            raise ValueError("invalid json")
        return self._payload

    def raise_for_status(self):
        if self._status >= 400:
            raise requests.exceptions.RequestException(f"HTTP {self._status}")


def _mock_sanitize_ok(monkeypatch):
    monkeypatch.setattr(svc_mod, "sanitize_header_value", lambda v: v)


def _mock_jwt_decode(monkeypatch, exp_delta_seconds=3600, raises=None):
    def _decode(token, options=None):
        if raises:
            raise raises
        exp = int((datetime.now() + timedelta(seconds=exp_delta_seconds)).timestamp())
        return {"exp": exp}
    monkeypatch.setattr(svc_mod.jwt, "decode", _decode)


@pytest.fixture
def client():
    return GatewayClient()


# ─── authenticate / get_token ──────────────────────────────────────────────

def test_authenticate_success_armazena_token_sanitizado(client, monkeypatch):
    _mock_sanitize_ok(monkeypatch)
    _mock_jwt_decode(monkeypatch)
    monkeypatch.setattr(
        svc_mod.requests, "post",
        lambda url, **kwargs: DummyResponse(payload={"access_token": "tok-abc"}),
    )

    client.authenticate()

    assert client.token == "tok-abc"
    assert isinstance(client.token_expiry, datetime)


def test_authenticate_resposta_nao_json_levanta_erro(client, monkeypatch):
    monkeypatch.setattr(
        svc_mod.requests, "post",
        lambda url, **kwargs: DummyResponse(invalid_json=True),
    )

    with pytest.raises(RuntimeError, match="Resposta inválida da Gateway"):
        client.authenticate()


def test_authenticate_sem_access_token_levanta_erro(client, monkeypatch):
    monkeypatch.setattr(
        svc_mod.requests, "post",
        lambda url, **kwargs: DummyResponse(payload={"outro_campo": True}),
    )

    with pytest.raises(RuntimeError, match="Token ausente"):
        client.authenticate()


def test_authenticate_erro_ao_decodificar_jwt(client, monkeypatch):
    _mock_sanitize_ok(monkeypatch)
    _mock_jwt_decode(monkeypatch, raises=Exception("assinatura invalida"))
    monkeypatch.setattr(
        svc_mod.requests, "post",
        lambda url, **kwargs: DummyResponse(payload={"access_token": "tok"}),
    )

    with pytest.raises(RuntimeError, match="Erro ao decodificar token JWT"):
        client.authenticate()


def test_authenticate_resposta_falsy_levanta_erro(client, monkeypatch):
    monkeypatch.setattr(svc_mod.requests, "post", lambda url, **kwargs: None)

    with pytest.raises(RuntimeError, match="Falha ao autenticar na Gateway"):
        client.authenticate()


def test_get_token_autentica_quando_nao_ha_token(client, monkeypatch):
    chamadas = {"n": 0}

    def _fake_authenticate():
        chamadas["n"] += 1
        client.token = "novo-token"
        client.token_expiry = datetime.now() + timedelta(hours=1)

    monkeypatch.setattr(client, "authenticate", _fake_authenticate)

    resultado = client.get_token()

    assert resultado == "novo-token"
    assert chamadas["n"] == 1


def test_get_token_reutiliza_token_valido(client, monkeypatch):
    client.token = "token-valido"
    client.token_expiry = datetime.now() + timedelta(hours=1)

    def _fake_authenticate():
        raise AssertionError("não deveria reautenticar com token ainda válido")

    monkeypatch.setattr(client, "authenticate", _fake_authenticate)

    assert client.get_token() == "token-valido"


def test_get_token_thread_safe_autentica_apenas_1_vez_sob_concorrencia(client, monkeypatch):
    """A curadoria de documentos grandes chama gerar_texto() de várias
    threads simultâneas usando a mesma instância de GatewayClient (ver
    avaliar_documento em curadoria_service.py). Se 10 threads pedirem
    get_token() ao mesmo tempo sem token ainda carregado, o lock deve
    garantir que authenticate() seja disparado só 1 vez, não 10."""
    import threading
    import time as time_mod

    chamadas = {"n": 0}
    lock_contador = threading.Lock()

    def _fake_authenticate():
        with lock_contador:
            chamadas["n"] += 1
        time_mod.sleep(0.05)  # simula latência de rede, dá tempo de outras threads competirem
        client.token = "token-unico"
        client.token_expiry = datetime.now() + timedelta(hours=1)

    monkeypatch.setattr(client, "authenticate", _fake_authenticate)

    threads = [threading.Thread(target=client.get_token) for _ in range(10)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    assert chamadas["n"] == 1
    assert client.token == "token-unico"


def test_get_token_reautentica_quando_perto_de_expirar(client, monkeypatch):
    client.token = "token-quase-expirando"
    client.token_expiry = datetime.now() + timedelta(seconds=30)  # < 60s de margem
    chamadas = {"n": 0}

    def _fake_authenticate():
        chamadas["n"] += 1
        client.token = "token-renovado"
        client.token_expiry = datetime.now() + timedelta(hours=1)

    monkeypatch.setattr(client, "authenticate", _fake_authenticate)

    assert client.get_token() == "token-renovado"
    assert chamadas["n"] == 1


# ─── _extrair_texto_gerado ──────────────────────────────────────────────────

def test_extrair_texto_gerado_caminho_esperado(client):
    resposta = {
        "output_collection": {
            "output_datas": [
                {
                    "workflow_step_output_collection": {
                        "output": {
                            "json_data": {
                                "choices": [{"message": {"content": "texto gerado aqui"}}]
                            }
                        }
                    }
                }
            ]
        }
    }
    assert client._extrair_texto_gerado(resposta) == "texto gerado aqui"


def test_extrair_texto_gerado_busca_generica_quando_formato_diferente(client):
    resposta = {
        "algum_wrapper_inesperado": {
            "nivel2": [
                {"role": "assistant", "content": "achei por busca genérica"}
            ]
        }
    }
    assert client._extrair_texto_gerado(resposta) == "achei por busca genérica"


def test_extrair_texto_gerado_nao_encontrado_levanta_erro(client):
    resposta = {"nada_relevante": "sem conteudo aqui"}
    with pytest.raises(RuntimeError, match="Não foi possível localizar o texto gerado"):
        client._extrair_texto_gerado(resposta)


# ─── gerar_texto (fluxo completo com polling) ───────────────────────────────

def _authenticated(client):
    client.token = "tok-fixo"
    client.token_expiry = datetime.now() + timedelta(hours=1)
    client.workflow_url = "https://wf.test.example.com"
    client.status_url = "https://status.test.example.com"


def test_gerar_texto_sucesso_no_primeiro_poll(client, monkeypatch):
    _authenticated(client)
    monkeypatch.setattr(svc_mod.time, "sleep", lambda *_: None)

    def _fake_post(url, json=None, **kwargs):
        assert "/request/start-workflow" in url
        return DummyResponse(payload={"workflow_execution_id": "exec-1"})

    def _fake_get(url, **kwargs):
        assert "workflow-execution-status/exec-1" in url
        return DummyResponse(payload={
            "status": "WF_COMPLETED_SUCCESS",
            "output_collection": {
                "output_datas": [{
                    "workflow_step_output_collection": {
                        "output": {"json_data": {"choices": [{"message": {"content": "ok gerado"}}]}}
                    }
                }]
            },
        })

    monkeypatch.setattr(svc_mod.requests, "post", _fake_post)
    monkeypatch.setattr(svc_mod.requests, "get", _fake_get)

    resultado = client.gerar_texto("system", "user", "modelo-x")
    assert resultado == "ok gerado"


def test_gerar_texto_sem_execution_id_levanta_erro(client, monkeypatch):
    _authenticated(client)
    monkeypatch.setattr(svc_mod.requests, "post", lambda url, **kwargs: DummyResponse(payload={}))

    with pytest.raises(RuntimeError, match="não retornou workflow_execution_id"):
        client.gerar_texto("system", "user", "modelo-x")


def test_gerar_texto_workflow_falhou(client, monkeypatch):
    _authenticated(client)
    monkeypatch.setattr(svc_mod.time, "sleep", lambda *_: None)
    monkeypatch.setattr(
        svc_mod.requests, "post",
        lambda url, **kwargs: DummyResponse(payload={"workflow_execution_id": "exec-2"}),
    )
    monkeypatch.setattr(
        svc_mod.requests, "get",
        lambda url, **kwargs: DummyResponse(payload={"status": "WF_COMPLETED_ERROR"}),
    )

    with pytest.raises(RuntimeError, match="Workflow falhou"):
        client.gerar_texto("system", "user", "modelo-x")


def test_gerar_texto_timeout_espera(client, monkeypatch):
    _authenticated(client)
    monkeypatch.setattr(svc_mod.time, "sleep", lambda *_: None)
    monkeypatch.setattr(
        svc_mod.requests, "post",
        lambda url, **kwargs: DummyResponse(payload={"workflow_execution_id": "exec-3"}),
    )
    monkeypatch.setattr(
        svc_mod.requests, "get",
        lambda url, **kwargs: DummyResponse(payload={"status": "WF_RUNNING"}),
    )

    # Primeira chamada (inicio) retorna 0; chamadas seguintes já excedem o timeout.
    monotonic_valores = iter([0] + [1000] * 10)
    monkeypatch.setattr(svc_mod.time, "monotonic", lambda: next(monotonic_valores))

    with pytest.raises(RuntimeError, match="Tempo esgotado"):
        client.gerar_texto("system", "user", "modelo-x")


def test_gerar_texto_envia_payload_com_mensagens_corretas(client, monkeypatch):
    _authenticated(client)
    monkeypatch.setattr(svc_mod.time, "sleep", lambda *_: None)
    payload_capturado = {}

    def _fake_post(url, json=None, **kwargs):
        if "/request/start-workflow" in url:
            payload_capturado.update(json)
            return DummyResponse(payload={"workflow_execution_id": "exec-4"})
        raise AssertionError(f"POST inesperado: {url}")

    monkeypatch.setattr(svc_mod.requests, "post", _fake_post)
    monkeypatch.setattr(
        svc_mod.requests, "get",
        lambda url, **kwargs: DummyResponse(payload={
            "status": "WF_COMPLETED_SUCCESS",
            "choices": [{"message": {"role": "assistant", "content": "resposta"}}],
        }),
    )

    client.gerar_texto("prompt de sistema", "prompt do usuario", "modelo-y", temperature=0.2)

    props = payload_capturado["input_collection"]["input_properties"][0]["workflow_step_properties"]
    valores = {p["full_path"]: p["actual_value"] for p in props}
    assert valores["gen_params.model"] == "modelo-y"
    assert valores["gen_params.messages"] == [
        {"role": "system", "content": "prompt de sistema"},
        {"role": "user", "content": "prompt do usuario"},
    ]
    assert valores["gen_params.temperature"] == 0.2
