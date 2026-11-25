# tests/unit/test_curadoria_controller.py
"""
Testa as rotas /avaliar_documentos e /relatorio_html como um usuário
real usaria — enviando arquivos e recebendo a resposta HTTP. A Gateway
real é substituída por um dublê que nunca sai para a internet. Também
testa /health/liveness e /health/readiness.
"""
import io
import json

import docx
from fastapi.testclient import TestClient
from src.main import app
from src.services.gateway_apim_service import GatewayClient
from src.services import curadoria_service


def _docx_bytes_validos(texto: str = "Parágrafo de teste com conteúdo suficiente para extração.") -> bytes:
    """Gera um .docx mínimo, porém válido, em memória — usado nos testes
    que precisam passar da extração e chegar até a chamada ao modelo."""
    documento = docx.Document()
    documento.add_paragraph(texto)
    buffer = io.BytesIO()
    documento.save(buffer)
    return buffer.getvalue()


class FakeGatewayClientOK:
    """Fake do GatewayClient: nunca toca a rede. Distingue a chamada de
    avaliação (response_format json_object) da chamada de resumo
    narrativo (texto livre) pelo parâmetro response_format, igual às
    duas chamadas reais que curadoria_service.py faz."""

    def gerar_texto(self, system_prompt, user_prompt, model, *, temperature=0,
                     max_tokens=4096, response_format=None, **kwargs):
        if response_format and response_format.get("type") == "json_object":
            return json.dumps({
                "criterios": [
                    {
                        "id": "escrita_autocontida",
                        "status": "atende",
                        "justificativa": "Texto claro e autocontido.",
                        "risco_ia": "",
                        "trechos_problematicos": [],
                        "localizacao": ["Página 1"],
                        "sugestao_correcao": "",
                    },
                ]
            })
        return "Resumo de teste do diagnóstico deste documento."


class FakeGatewayClientAuthError:
    """Fake que simula falha de autenticação/rede na Gateway — usado
    para testar que o controller trata esse erro sem quebrar o lote."""

    def gerar_texto(self, *args, **kwargs):
        raise RuntimeError("Falha ao autenticar na Gateway")


def _with_override(fake_cls):
    """Substitui a dependência real GatewayClient (Depends(GatewayClient)
    no controller) pelo fake informado, para o teste não depender de
    rede. Sempre usar com _clear_override() no finally, senão o
    override "escapa" e afeta outros testes que rodem depois."""
    app.dependency_overrides[GatewayClient] = lambda: fake_cls()
    return TestClient(app)


def _clear_override():
    app.dependency_overrides.pop(GatewayClient, None)


# ─── /avaliar_documentos — validações do controller ────────────────────────

def test_avaliar_documentos_extensao_nao_permitida():
    """Extensão fora de ALLOWED_EXTENSIONS deve ser rejeitada com 400
    antes mesmo de tentar extrair ou avaliar o conteúdo do arquivo."""
    client = _with_override(FakeGatewayClientOK)
    try:
        files = [("files", ("arquivo.exe", b"conteudo", "application/octet-stream"))]
        resp = client.post("/avaliar_documentos", files=files)
        assert resp.status_code == 400
        assert "Extensão não permitida" in resp.json()["detail"]
    finally:
        _clear_override()


def test_avaliar_documentos_excede_quantidade_maxima():
    """Acima de MAX_FILES_PER_BATCH (padrão 5), o lote inteiro é
    rejeitado com 400 — a validação de quantidade acontece antes de ler
    qualquer arquivo, então nem importa o conteúdo dos 6 enviados."""
    client = _with_override(FakeGatewayClientOK)
    try:
        many = [("files", (f"{i}.pdf", b"x", "application/pdf")) for i in range(6)]
        resp = client.post("/avaliar_documentos", files=many)
        assert resp.status_code == 400
        assert "máximo" in resp.json()["detail"]
    finally:
        _clear_override()


def test_avaliar_documentos_excede_tamanho_maximo():
    """Arquivo acima de MAX_FILE_SIZE_MB (padrão 20MB) é rejeitado com
    400 individualmente — a validação de tamanho ocorre por arquivo,
    depois de ler os bytes (não dá para saber o tamanho real sem ler)."""
    client = _with_override(FakeGatewayClientOK)
    try:
        big_bytes = b"a" * (21 * 1024 * 1024)
        files = [("files", ("grande.pdf", big_bytes, "application/pdf"))]
        resp = client.post("/avaliar_documentos", files=files)
        assert resp.status_code == 400
        assert "excede o limite" in resp.json()["detail"]
    finally:
        _clear_override()


def test_avaliar_documentos_sucesso_fim_a_fim_com_fake_client():
    """Fluxo feliz completo: upload -> extração real do .docx -> fake
    LLM devolve avaliação 100% "atende" -> score 100 -> resumo
    narrativo do fake. Cobre a integração entre controller e service
    sem precisar da Gateway real."""
    client = _with_override(FakeGatewayClientOK)
    try:
        files = [("files", ("valido.docx", _docx_bytes_validos(),
                             "application/vnd.openxmlformats-officedocument.wordprocessingml.document"))]
        resp = client.post("/avaliar_documentos", files=files)
        assert resp.status_code == 200
        resultado = resp.json()["resultados"][0]
        assert resultado["arquivo"] == "valido.docx"
        assert "erro" not in resultado
        assert resultado["score_info"]["score"] == 100.0
        assert resultado["resumo_narrativo"] == "Resumo de teste do diagnóstico deste documento."
    finally:
        _clear_override()


def test_avaliar_documentos_extensao_valida_bytes_invalidos_nao_quebra_lote():
    """Extensão permitida, mas conteúdo não é um PDF de verdade: a
    extração falha, mas o endpoint continua respondendo 200 com o erro
    reportado dentro do item do lote — igual ao comportamento tolerante
    do Streamlit hoje (processar_arquivo nunca propaga essa exceção)."""
    client = _with_override(FakeGatewayClientOK)
    try:
        files = [("files", ("nao_e_um_pdf.pdf", b"isto nao e um pdf valido", "application/pdf"))]
        resp = client.post("/avaliar_documentos", files=files)
        assert resp.status_code == 200
        resultados = resp.json()["resultados"]
        assert len(resultados) == 1
        assert "erro" in resultados[0]
    finally:
        _clear_override()


# ─── curadoria_service — nível de serviço, sem passar pelo HTTP ────────────
# Os testes abaixo chamam curadoria_service diretamente (sem TestClient),
# úteis para verificar a integração avaliar_documento -> calcular_score
# sem o overhead de montar uma requisição HTTP completa. Cobertura mais
# extensa e granular de curadoria_service está em test_curadoria_service.py.

def test_avaliar_documento_e_calcular_score_com_fake_client():
    """Encadeamento avaliar_documento -> calcular_score com 1 único
    critério "atende": confirma que o score bate 100 e a classificação
    correta é aplicada quando não há nenhum problema."""
    blocos = [curadoria_service.Bloco(localizacao="Página 1", texto="Texto de exemplo bem escrito e claro.")]
    fake_client = FakeGatewayClientOK()

    avaliacao = curadoria_service.avaliar_documento(blocos, "PDF", fake_client)
    assert avaliacao["criterios"][0]["id"] == "escrita_autocontida"
    assert avaliacao["criterios"][0]["status"] == "atende"

    score_info = curadoria_service.calcular_score(avaliacao)
    assert score_info["score"] == 100.0
    assert score_info["classificacao"] == "Pronto para uso em IA"
    assert score_info["qtd_problemas_criticos"] == 0


def test_gerar_resumo_narrativo_com_fake_client():
    avaliacao = {"criterios": [{"id": "escrita_autocontida", "status": "atende", "justificativa": "ok"}]}
    score_info = {"classificacao": "Pronto para uso em IA", "score": 100.0}
    resumo = curadoria_service.gerar_resumo_narrativo(avaliacao, score_info, FakeGatewayClientOK())
    assert resumo == "Resumo de teste do diagnóstico deste documento."


def test_avaliar_documentos_erro_de_autenticacao_vira_item_com_erro_no_lote():
    """Documento válido (extração funciona), mas o client falha ao
    chamar a Gateway — processar_arquivo captura a exceção dentro de
    avaliar_documento e devolve {"erro": ...} no item do lote, então o
    HTTP da rota continua 200 (mesmo comportamento tolerante do
    Streamlit hoje: 1 arquivo com problema não derruba o lote inteiro)."""
    client = _with_override(FakeGatewayClientAuthError)
    try:
        files = [("files", ("valido.docx", _docx_bytes_validos(),
                             "application/vnd.openxmlformats-officedocument.wordprocessingml.document"))]
        resp = client.post("/avaliar_documentos", files=files)
        assert resp.status_code == 200
        assert "erro" in resp.json()["resultados"][0]
    finally:
        _clear_override()


# ─── /health ────────────────────────────────────────────────────────────────

def test_health_liveness_sem_autenticacao():
    with TestClient(app) as c:
        resp = c.get("/health/liveness")
        assert resp.status_code == 200
        assert resp.json()["status"] == "OK"


def test_health_readiness_sem_autenticacao():
    with TestClient(app) as c:
        resp = c.get("/health/readiness")
        assert resp.status_code == 200
        assert resp.json()["status"] == "OK"
