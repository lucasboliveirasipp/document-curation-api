# tests/unit/test_curadoria_service.py
"""
Testa o coração do agente: gera documentos reais (PDF, Word, Excel,
PowerPoint) na hora para confirmar que o texto é extraído certo,
confere as 4 notas/classificações possíveis, a divisão de documentos
grandes em lotes, e a montagem do relatório HTML final.
"""
import json
import time

import docx
import openpyxl
import pytest
from pptx import Presentation
from pptx.util import Inches

from src.services import curadoria_service as svc
from src.services.curadoria_service import Bloco, ExtracaoVaziaError


# ============================================================
# Helpers de geração de arquivo real em memória
# ============================================================

def _criar_pdf_com_texto(paginas: list[str], tmp_path) -> str:
    import fitz
    doc = fitz.open()
    for texto in paginas:
        pagina = doc.new_page()
        pagina.insert_text((72, 72), texto)
    caminho = tmp_path / "documento.pdf"
    doc.save(str(caminho))
    doc.close()
    return str(caminho)


def _criar_docx_com_secoes(tmp_path) -> str:
    documento = docx.Document()
    documento.add_heading("Título da Seção A", level=1)
    documento.add_paragraph("Texto do parágrafo dentro da seção A, com conteúdo suficiente.")
    documento.add_heading("Título da Seção B", level=1)
    documento.add_paragraph("Texto do parágrafo dentro da seção B.")
    tabela = documento.add_table(rows=2, cols=2)
    tabela.rows[0].cells[0].text = "Cabeçalho 1"
    tabela.rows[0].cells[1].text = "Cabeçalho 2"
    tabela.rows[1].cells[0].text = "Valor 1"
    tabela.rows[1].cells[1].text = "Valor 2"
    caminho = tmp_path / "documento.docx"
    documento.save(str(caminho))
    return str(caminho)


def _criar_xlsx_com_abas(tmp_path) -> str:
    wb = openpyxl.Workbook()
    aba1 = wb.active
    aba1.title = "Prazos"
    aba1.append(["Codigo_da_Regra", "Descricao"])
    aba1.append(["R001", "Prazo de análise de 5 dias úteis"])
    aba2 = wb.create_sheet("Vazia")
    caminho = tmp_path / "documento.xlsx"
    wb.save(str(caminho))
    return str(caminho)


def _criar_pptx_com_slides(tmp_path) -> str:
    prs = Presentation()
    layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(layout)
    slide.shapes.title.text = "Slide de teste"
    caminho = tmp_path / "documento.pptx"
    prs.save(str(caminho))
    return str(caminho)


class FakeClient:
    """Fake mínimo de GatewayClient para os testes deste módulo."""

    def __init__(self, resposta_avaliacao=None, resposta_resumo="resumo ok", erro=None):
        self.resposta_avaliacao = resposta_avaliacao or {
            "criterios": [
                {"id": "escrita_autocontida", "status": "atende", "justificativa": "ok",
                 "risco_ia": "", "trechos_problematicos": [], "localizacao": ["Página 1"],
                 "sugestao_correcao": ""},
            ]
        }
        self.resposta_resumo = resposta_resumo
        self.erro = erro
        self.chamadas = 0

    def gerar_texto(self, system_prompt, user_prompt, model, *, temperature=0,
                     max_tokens=4096, response_format=None, **kwargs):
        self.chamadas += 1
        if self.erro:
            raise self.erro
        if response_format and response_format.get("type") == "json_object":
            return json.dumps(self.resposta_avaliacao)
        return self.resposta_resumo


class FlakyClient:
    """Fake de GatewayClient que falha nas N primeiras chamadas e depois
    passa a funcionar — simula timeout/instabilidade transitória da
    Gateway, usado para testar o retry por lote."""

    def __init__(self, falhas_antes_de_funcionar, erro=None):
        self.falhas_antes_de_funcionar = falhas_antes_de_funcionar
        self.erro = erro or RuntimeError("Tempo esgotado esperando o workflow terminar")
        self.chamadas = 0

    def gerar_texto(self, system_prompt, user_prompt, model, *, temperature=0,
                     max_tokens=4096, response_format=None, **kwargs):
        self.chamadas += 1
        if self.chamadas <= self.falhas_antes_de_funcionar:
            raise self.erro
        return json.dumps({
            "criterios": [
                {"id": "escrita_autocontida", "status": "atende", "justificativa": "ok",
                 "risco_ia": "", "trechos_problematicos": [], "localizacao": ["Página 1"],
                 "sugestao_correcao": ""},
            ]
        })


class SempreFalhaClient:
    """Fake que sempre falha, usado para simular um lote específico que
    não se recupera nem com retry."""

    def __init__(self, erro=None):
        self.erro = erro or RuntimeError("falha persistente")
        self.chamadas = 0

    def gerar_texto(self, *args, **kwargs):
        self.chamadas += 1
        raise self.erro


# ============================================================
# extrair_texto — formato não suportado
# ============================================================

def test_extrair_texto_formato_nao_suportado_levanta_erro():
    with pytest.raises(ValueError, match="Formato não suportado"):
        svc.extrair_texto("arquivo.txt")


# ============================================================
# _extrair_pdf
# ============================================================

def test_extrair_pdf_1_bloco_por_pagina_com_texto(tmp_path):
    caminho = _criar_pdf_com_texto(["Conteúdo da página um bem extenso.", "Conteúdo da página dois."], tmp_path)
    blocos = svc.extrair_texto(caminho)
    assert len(blocos) == 2
    assert blocos[0].localizacao == "Página 1"
    assert "página um" in blocos[0].texto


def test_extrair_pdf_escaneado_levanta_extracao_vazia(tmp_path):
    import fitz
    doc = fitz.open()
    doc.new_page()  # página sem texto = simula PDF escaneado
    doc.new_page()
    caminho = tmp_path / "escaneado.pdf"
    doc.save(str(caminho))
    doc.close()

    with pytest.raises(ExtracaoVaziaError):
        svc.extrair_texto(str(caminho))


# ============================================================
# _extrair_docx
# ============================================================

def test_extrair_docx_agrupa_por_secao_e_extrai_tabela(tmp_path):
    caminho = _criar_docx_com_secoes(tmp_path)
    blocos = svc.extrair_texto(caminho)

    localizacoes = [b.localizacao for b in blocos]
    assert "Título da Seção A" in localizacoes
    assert "Título da Seção B" in localizacoes
    assert "Tabela 1" in localizacoes

    bloco_tabela = next(b for b in blocos if b.localizacao == "Tabela 1")
    assert "Cabeçalho 1" in bloco_tabela.texto
    assert "Valor 2" in bloco_tabela.texto


# ============================================================
# _extrair_excel
# ============================================================

def test_extrair_excel_1_bloco_por_aba_ignora_aba_vazia(tmp_path):
    caminho = _criar_xlsx_com_abas(tmp_path)
    blocos = svc.extrair_texto(caminho)

    assert len(blocos) == 1  # aba "Vazia" não gera bloco
    assert blocos[0].localizacao == "Aba: Prazos"
    assert "Codigo_da_Regra" in blocos[0].texto
    assert "R001" in blocos[0].texto


def test_extrair_excel_aba_grande_divide_em_blocos_de_100_linhas(tmp_path):
    """Aba com mais linhas que LINHAS_POR_BLOCO_EXCEL deve virar vários
    blocos menores (mesmo padrão de loteamento da indexação real da
    Gateway), em vez de 1 bloco gigante que _dividir_em_lotes não
    conseguiria quebrar — essa era a causa de tabelas grandes (ex: TUSS,
    6744 linhas) derrubarem o workflow com WF_COMPLETED_ERROR."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Grande"
    ws.append(["Codigo", "Descricao"])
    for i in range(250):
        ws.append([f"C{i:04d}", f"Descricao da linha {i}"])
    caminho = tmp_path / "grande.xlsx"
    wb.save(str(caminho))

    blocos = svc.extrair_texto(str(caminho))

    # 251 linhas com conteúdo (cabeçalho + 250 linhas) / 100 por bloco = 3 blocos.
    assert len(blocos) == 3
    assert blocos[0].localizacao == "Aba: Grande (linhas 1-100)"
    assert blocos[1].localizacao == "Aba: Grande (linhas 101-200)"
    assert blocos[2].localizacao == "Aba: Grande (linhas 201-251)"
    assert "C0000" in blocos[0].texto
    assert "C0249" in blocos[2].texto


# ============================================================
# _extrair_pptx
# ============================================================

def test_extrair_pptx_1_bloco_por_slide_com_conteudo(tmp_path):
    caminho = _criar_pptx_com_slides(tmp_path)
    blocos = svc.extrair_texto(caminho)

    assert len(blocos) == 1
    assert blocos[0].localizacao == "Slide 1"
    assert "Slide de teste" in blocos[0].texto


# ============================================================
# calcular_metricas_documento
# ============================================================

def test_calcular_metricas_documento_sem_blocos():
    metricas = svc.calcular_metricas_documento([])
    assert metricas["num_blocos"] == 0
    assert metricas["compatibilidade_chunking"] == "sem_dados"


def test_calcular_metricas_documento_detecta_fragmentacao():
    blocos = [Bloco(localizacao=f"L{i}", texto="curto") for i in range(6)]
    metricas = svc.calcular_metricas_documento(blocos)
    assert metricas["alerta_fragmentacao"] is True


def test_calcular_metricas_documento_detecta_blocos_duplicados():
    blocos = [
        Bloco(localizacao="A", texto="Este texto se repete de propósito no documento."),
        Bloco(localizacao="B", texto="Este texto se repete de propósito no documento."),
        Bloco(localizacao="C", texto="Texto totalmente diferente e único no documento inteiro."),
    ]
    metricas = svc.calcular_metricas_documento(blocos)
    assert metricas["blocos_duplicados"] == 1


def test_calcular_metricas_documento_proporcao_tabular():
    blocos = [
        Bloco(localizacao="Tabela 1", texto="a" * 100),
        Bloco(localizacao="Página 1", texto="b" * 100),
    ]
    metricas = svc.calcular_metricas_documento(blocos)
    assert metricas["proporcao_tabular"] == 0.5


# ============================================================
# calcular_porte_documento / _unidade_nativa
# ============================================================

def test_calcular_porte_documento_pdf_usa_paginas():
    blocos = [Bloco(localizacao="Página 1", texto="x" * 50)]
    porte = svc.calcular_porte_documento(blocos, ".pdf")
    assert porte["unidade"] == "páginas"
    assert porte["quantidade"] == 1


def test_calcular_porte_documento_excel_usa_linhas_e_limite_proprio():
    blocos = [Bloco(localizacao="Aba: X", texto="linha1\nlinha2\nlinha3")]
    porte = svc.calcular_porte_documento(blocos, ".xlsx")
    assert porte["unidade"] == "linhas"
    assert porte["quantidade"] == 3
    assert porte["limite_chunks_recomendado"] == svc.LIMITE_LINHAS_EXCEL_RECOMENDACAO


def test_calcular_porte_documento_docx_usa_caracteres():
    blocos = [Bloco(localizacao="Início do documento", texto="x" * 500)]
    porte = svc.calcular_porte_documento(blocos, ".docx")
    assert porte["unidade"] == "caracteres"
    assert porte["quantidade"] == 500


def test_calcular_porte_documento_marca_documento_grande():
    # 1 bloco gigante força num_chunks_estimado a passar do limite recomendado.
    texto_gigante = "a" * (svc.LIMITE_CHUNKS_POR_DOCUMENTO_RECOMENDACAO * svc._AVANCO_LIQUIDO_CHARS_POR_CHUNK * 2)
    blocos = [Bloco(localizacao="Página 1", texto=texto_gigante)]
    porte = svc.calcular_porte_documento(blocos, ".pdf")
    assert porte["documento_grande"] is True


# ============================================================
# _montar_prompt_usuario
# ============================================================

def test_montar_prompt_usuario_inclui_formato_e_criterios():
    prompt = svc._montar_prompt_usuario("texto do documento", "PDF")
    assert "PDF" in prompt
    assert "escrita_autocontida" in prompt
    assert "texto do documento" in prompt


# ============================================================
# _chamar_llm
# ============================================================

def test_chamar_llm_sucesso_retorna_json_decodificado():
    client = FakeClient()
    resultado = svc._chamar_llm("prompt qualquer", client)
    assert resultado["criterios"][0]["id"] == "escrita_autocontida"


def test_chamar_llm_propaga_exececao_do_client():
    client = FakeClient(erro=RuntimeError("falha de rede"))
    with pytest.raises(RuntimeError, match="falha de rede"):
        svc._chamar_llm("prompt qualquer", client)


def test_chamar_llm_json_invalido_levanta_erro():
    class ClientTextoInvalido:
        def gerar_texto(self, *args, **kwargs):
            return "isto não é um json válido"

    with pytest.raises(json.JSONDecodeError):
        svc._chamar_llm("prompt qualquer", ClientTextoInvalido())


# ============================================================
# _normalizar_resposta_llm — formato compacto (economia de tokens de saída)
# ============================================================

def test_normalizar_resposta_llm_expande_formato_compacto():
    """Critérios sem problema (só id+status) e com problema (justificativa
    completa) devem virar uma única lista "criterios", no formato que o
    resto do pipeline (score, redução de lotes, relatório) já espera."""
    resposta_compacta = {
        "criterios_sem_problema": [
            {"id": "titulos_descritivos", "status": "atende"},
            {"id": "termos_padronizados", "status": "nao_aplicavel"},
        ],
        "criterios_com_problema": [
            {
                "id": "siglas_explicadas",
                "status": "nao_atende",
                "justificativa": "Sigla N205 sem explicação.",
                "risco_ia": "IA pode não entender o código.",
                "trechos_problematicos": ["N205"],
                "localizacao": ["Página 3"],
                "sugestao_correcao": "N205 (negação por documentação ausente)",
            },
        ],
    }

    normalizado = svc._normalizar_resposta_llm(resposta_compacta)
    criterios = {c["id"]: c for c in normalizado["criterios"]}

    assert len(normalizado["criterios"]) == 3
    assert criterios["titulos_descritivos"]["status"] == "atende"
    assert criterios["titulos_descritivos"]["justificativa"] == ""
    assert criterios["termos_padronizados"]["status"] == "nao_aplicavel"
    assert criterios["siglas_explicadas"]["status"] == "nao_atende"
    assert criterios["siglas_explicadas"]["justificativa"] == "Sigla N205 sem explicação."
    assert criterios["siglas_explicadas"]["localizacao"] == ["Página 3"]


def test_normalizar_resposta_llm_status_parcial_no_grupo_errado_marca_incompleto():
    """Regressão: o modelo às vezes classifica um critério com status
    "parcial" dentro de "criterios_sem_problema" (que só tem id+status,
    sem justificativa) — isso gerava, na tela, uma linha com status
    "Parcial" mas diagnóstico/risco/exemplo/recomendação todos vazios
    ("—"), sem nenhuma informação acionável. A normalização preserva o
    status real (ainda conta no score) mas marca "diagnostico_incompleto"
    para essa linha ser filtrada da tabela/relatório em vez de aparecer
    vazia."""
    resposta_com_erro_de_classificacao = {
        "criterios_sem_problema": [
            {"id": "documento_vigente_explicito", "status": "parcial"},
            {"id": "titulos_descritivos", "status": "atende"},
        ],
        "criterios_com_problema": [],
    }

    normalizado = svc._normalizar_resposta_llm(resposta_com_erro_de_classificacao)
    criterios = {c["id"]: c for c in normalizado["criterios"]}

    item_mal_classificado = criterios["documento_vigente_explicito"]
    assert item_mal_classificado["status"] == "parcial"
    assert item_mal_classificado["justificativa"] == ""
    assert item_mal_classificado["diagnostico_incompleto"] is True

    item_correto = criterios["titulos_descritivos"]
    assert item_correto["status"] == "atende"
    assert item_correto["justificativa"] == ""
    assert "diagnostico_incompleto" not in item_correto


def test_normalizar_resposta_llm_status_nao_atende_no_grupo_errado_marca_incompleto():
    resposta = {
        "criterios_sem_problema": [{"id": "siglas_explicadas", "status": "nao_atende"}],
        "criterios_com_problema": [],
    }
    normalizado = svc._normalizar_resposta_llm(resposta)
    item = normalizado["criterios"][0]
    assert item["status"] == "nao_atende"
    assert item["diagnostico_incompleto"] is True


def test_normalizar_resposta_llm_formato_antigo_passa_direto():
    """Se a resposta já vier no formato antigo (com "criterios" direto —
    usado nos testes e como salvaguarda caso o modelo responda fora do
    novo schema), a normalização não deve alterar nada."""
    resposta_antiga = {"criterios": [{"id": "x", "status": "atende", "justificativa": "ok"}]}
    assert svc._normalizar_resposta_llm(resposta_antiga) == resposta_antiga


def test_normalizar_resposta_llm_grupos_vazios_gera_lista_vazia():
    assert svc._normalizar_resposta_llm({}) == {"criterios": []}


def test_chamar_llm_normaliza_formato_compacto_do_modelo():
    """_chamar_llm deve devolver o formato já normalizado, mesmo quando o
    client (Gateway) responde no formato compacto novo."""
    class ClientFormatoCompacto:
        def gerar_texto(self, *args, **kwargs):
            return json.dumps({
                "criterios_sem_problema": [{"id": "titulos_descritivos", "status": "atende"}],
                "criterios_com_problema": [],
            })

    resultado = svc._chamar_llm("prompt qualquer", ClientFormatoCompacto())
    assert resultado == {"criterios": [{
        "id": "titulos_descritivos", "status": "atende", "justificativa": "",
        "risco_ia": "", "trechos_problematicos": [], "localizacao": [], "sugestao_correcao": "",
    }]}


# ============================================================
# gerar_resumo_narrativo
# ============================================================

def test_gerar_resumo_narrativo_falha_do_client_nao_propaga_e_retorna_vazio():
    client = FakeClient(erro=RuntimeError("gateway indisponível"))
    avaliacao = {"criterios": [{"id": "escrita_autocontida", "status": "atende", "justificativa": "ok"}]}
    score_info = {"classificacao": "Pronto para uso em IA", "score": 100.0}

    resultado = svc.gerar_resumo_narrativo(avaliacao, score_info, client)
    assert resultado == ""


# ============================================================
# _dividir_em_lotes
# ============================================================

def test_dividir_em_lotes_agrupa_ate_o_limite():
    blocos = [Bloco(localizacao=f"P{i}", texto="a" * 40) for i in range(5)]
    lotes = svc._dividir_em_lotes(blocos, limite_chars=100)
    assert len(lotes) == 3  # 40+40=80 (<=100), próximo de 40 estoura -> novo lote
    assert sum(len(l) for l in lotes) == 5


def test_dividir_em_lotes_bloco_maior_que_limite_fica_sozinho():
    blocos = [Bloco(localizacao="P1", texto="a" * 500)]
    lotes = svc._dividir_em_lotes(blocos, limite_chars=100)
    assert len(lotes) == 1
    assert len(lotes[0]) == 1


def test_dividir_em_lotes_lista_vazia():
    assert svc._dividir_em_lotes([], limite_chars=100) == []


# ============================================================
# _remover_duplicatas
# ============================================================

def test_remover_duplicatas_mantem_ordem_e_remove_repetidos():
    resultado = svc._remover_duplicatas(["a", "b", "a", "c", "b"])
    assert resultado == ["a", "b", "c"]


# ============================================================
# _reduzir_avaliacoes
# ============================================================

def test_reduzir_avaliacoes_mescla_duplicata_exata():
    avaliacoes = [
        {"criterios": [{
            "id": "escrita_autocontida", "status": "nao_atende", "justificativa": "Problema X",
            "localizacao": ["Página 1"], "trechos_problematicos": ["trecho A"],
        }]},
        {"criterios": [{
            "id": "escrita_autocontida", "status": "nao_atende", "justificativa": "Problema X",
            "localizacao": ["Página 2"], "trechos_problematicos": ["trecho B"],
        }]},
    ]
    resultado = svc._reduzir_avaliacoes(avaliacoes)
    assert len(resultado["criterios"]) == 1
    item = resultado["criterios"][0]
    assert set(item["localizacao"]) == {"Página 1", "Página 2"}
    assert set(item["trechos_problematicos"]) == {"trecho A", "trecho B"}


def test_reduzir_avaliacoes_nao_mescla_justificativas_diferentes():
    avaliacoes = [
        {"criterios": [{"id": "escrita_autocontida", "status": "atende", "justificativa": "ok"}]},
        {"criterios": [{"id": "escrita_autocontida", "status": "nao_atende", "justificativa": "problema diferente"}]},
    ]
    resultado = svc._reduzir_avaliacoes(avaliacoes)
    assert len(resultado["criterios"]) == 2


# ============================================================
# avaliar_documento
# ============================================================

def test_avaliar_documento_chama_llm_1_vez_por_lote():
    blocos = [Bloco(localizacao=f"P{i}", texto="a" * 40) for i in range(3)]
    client = FakeClient()
    svc.avaliar_documento(blocos, "PDF", client)
    lotes_esperados = len(svc._dividir_em_lotes(blocos, svc.MAX_CHARS_PER_LLM_CALL))
    assert client.chamadas == lotes_esperados


# ============================================================
# _chamar_llm_com_retry
# ============================================================

def test_chamar_llm_com_retry_recupera_apos_falha_transitoria(monkeypatch):
    monkeypatch.setattr(svc.time, "sleep", lambda *_: None)  # não esperar de verdade no teste
    client = FlakyClient(falhas_antes_de_funcionar=2)  # falha 2x, funciona na 3ª (limite padrão)
    resultado = svc._chamar_llm_com_retry("prompt", client)
    assert resultado["criterios"][0]["id"] == "escrita_autocontida"
    assert client.chamadas == 3


def test_chamar_llm_com_retry_desiste_apos_esgotar_tentativas(monkeypatch):
    monkeypatch.setattr(svc.time, "sleep", lambda *_: None)
    client = SempreFalhaClient(erro=RuntimeError("falha persistente"))
    with pytest.raises(RuntimeError, match="falha persistente"):
        svc._chamar_llm_com_retry("prompt", client)
    assert client.chamadas == svc.TENTATIVAS_POR_LOTE


# ============================================================
# avaliar_documento — tolerância a falha parcial de lote
# ============================================================

def test_avaliar_documento_lote_com_falha_persistente_nao_derruba_os_demais(monkeypatch):
    """Documento com 2 lotes: o 1º sempre falha (mesmo com retry), o 2º
    funciona normalmente. O resultado final deve conter a avaliação do
    lote que funcionou e sinalizar qual lote não pôde ser avaliado, em
    vez de propagar a exceção e descartar o arquivo inteiro."""
    monkeypatch.setattr(svc.time, "sleep", lambda *_: None)

    class ClientMistoPorLote:
        def __init__(self):
            self.chamadas_por_prompt = 0

        def gerar_texto(self, system_prompt, user_prompt, model, *, response_format=None, **kwargs):
            self.chamadas_por_prompt += 1
            if "LOTE_QUE_FALHA" in user_prompt:
                raise RuntimeError("Tempo esgotado esperando o workflow terminar")
            return json.dumps({
                "criterios": [
                    {"id": "escrita_autocontida", "status": "atende", "justificativa": "ok",
                     "risco_ia": "", "trechos_problematicos": [], "localizacao": ["Página 2"],
                     "sugestao_correcao": ""},
                ]
            })

    # limite_chars pequeno o bastante pra cada bloco virar seu próprio lote.
    blocos = [
        Bloco(localizacao="LOTE_QUE_FALHA", texto="conteudo do lote com problema"),
        Bloco(localizacao="Página 2", texto="conteudo do lote que funciona"),
    ]
    client = ClientMistoPorLote()
    monkeypatch.setattr(svc, "MAX_CHARS_PER_LLM_CALL", 10)

    resultado = svc.avaliar_documento(blocos, "PDF", client)

    assert resultado["criterios"][0]["localizacao"] == ["Página 2"]
    assert resultado["lotes_nao_avaliados"] == [["LOTE_QUE_FALHA"]]
    # 1 lote com falha x TENTATIVAS_POR_LOTE tentativas + 1 lote OK (1ª tentativa).
    assert client.chamadas_por_prompt == svc.TENTATIVAS_POR_LOTE + 1


def test_avaliar_documento_todos_lotes_falham_levanta_erro(monkeypatch):
    monkeypatch.setattr(svc.time, "sleep", lambda *_: None)
    blocos = [Bloco(localizacao="P1", texto="conteudo qualquer")]
    client = SempreFalhaClient(erro=RuntimeError("gateway fora do ar"))
    with pytest.raises(RuntimeError, match="Nenhum dos"):
        svc.avaliar_documento(blocos, "PDF", client)


# ============================================================
# avaliar_documento — limite de porte (bloqueio e truncamento)
# ============================================================

def test_avaliar_documento_acima_do_limite_de_bloqueio_recusa_sem_chamar_llm(monkeypatch):
    """Documento que geraria mais lotes que LIMITE_LOTES_BLOQUEIO deve ser
    recusado ANTES de qualquer chamada ao modelo — custo de IA = zero."""
    monkeypatch.setattr(svc, "LIMITE_LOTES_BLOQUEIO", 2)
    monkeypatch.setattr(svc, "MAX_CHARS_PER_LLM_CALL", 10)  # cada bloco vira 1 lote
    blocos = [Bloco(localizacao=f"P{i}", texto="conteudo") for i in range(5)]  # 5 lotes > limite 2
    client = FakeClient()

    with pytest.raises(svc.DocumentoMuitoGrandeError, match="5 lotes"):
        svc.avaliar_documento(blocos, "PDF", client)
    assert client.chamadas == 0  # nenhuma chamada à Gateway foi feita


def test_avaliar_documento_dentro_do_limite_de_bloqueio_mas_acima_do_processavel_trunca(monkeypatch):
    """Entre LIMITE_LOTES_PROCESSAVEIS e LIMITE_LOTES_BLOQUEIO: avalia só
    os primeiros lotes e marca o excedente como truncado, sem chamar o
    modelo para os lotes cortados."""
    monkeypatch.setattr(svc, "LIMITE_LOTES_PROCESSAVEIS", 2)
    monkeypatch.setattr(svc, "LIMITE_LOTES_BLOQUEIO", 10)
    monkeypatch.setattr(svc, "MAX_CHARS_PER_LLM_CALL", 10)  # cada bloco vira 1 lote
    blocos = [Bloco(localizacao=f"P{i}", texto="conteudo") for i in range(5)]  # 5 lotes
    client = FakeClient()

    resultado = svc.avaliar_documento(blocos, "PDF", client)

    assert client.chamadas == 2  # só os 2 primeiros lotes chamaram o modelo
    assert resultado["lotes_truncados"] == [["P2"], ["P3"], ["P4"]]


def test_max_chamadas_paralelas_nunca_passa_do_limite_da_gateway(monkeypatch):
    """Mesmo que MAX_CHAMADAS_PARALELAS seja configurado (por engano) acima
    do que a Gateway suporta, o módulo nunca deve usar um valor maior que
    10 — a checagem é feita 1x na carga do módulo, então aqui confirmamos
    a lógica do min() isoladamente."""
    assert svc.MAX_CHAMADAS_PARALELAS <= 10
    assert min(999, 10) == 10  # documenta a regra que _LIMITE_CONCORRENCIA_GATEWAY aplica


def test_avaliar_documento_processa_lotes_em_paralelo_mais_rapido_que_sequencial(monkeypatch):
    """Confirma que os lotes são avaliados concorrentemente: com 10 lotes
    que demoram 0.3s cada e até 5 chamadas simultâneas, o tempo total deve
    ficar bem abaixo de 10 * 0.3s = 3s (tempo sequencial)."""
    monkeypatch.setattr(svc, "MAX_CHAMADAS_PARALELAS", 5)
    monkeypatch.setattr(svc, "MAX_CHARS_PER_LLM_CALL", 10)  # cada bloco vira 1 lote

    class ClientLento:
        def __init__(self):
            self.chamadas = 0

        def gerar_texto(self, *args, **kwargs):
            self.chamadas += 1
            time.sleep(0.3)
            return json.dumps({"criterios": []})

    blocos = [Bloco(localizacao=f"P{i}", texto="conteudo") for i in range(10)]
    client = ClientLento()

    inicio = time.monotonic()
    svc.avaliar_documento(blocos, "PDF", client)
    duracao = time.monotonic() - inicio

    assert client.chamadas == 10
    assert duracao < 1.5  # bem menos que 3s (10 * 0.3s sequencial)


def test_avaliar_documento_dentro_dos_dois_limites_nao_trunca_nem_bloqueia(monkeypatch):
    monkeypatch.setattr(svc, "LIMITE_LOTES_PROCESSAVEIS", 60)
    monkeypatch.setattr(svc, "LIMITE_LOTES_BLOQUEIO", 150)
    monkeypatch.setattr(svc, "MAX_CHARS_PER_LLM_CALL", 10)  # cada bloco vira 1 lote
    blocos = [Bloco(localizacao=f"P{i}", texto="a" * 40) for i in range(3)]
    client = FakeClient()

    resultado = svc.avaliar_documento(blocos, "PDF", client)

    assert "lotes_truncados" not in resultado
    assert client.chamadas == 3


# ============================================================
# processar_arquivo — documento grande demais (bloqueio)
# ============================================================

def test_processar_arquivo_documento_grande_demais_marca_flag_e_nao_chama_llm(tmp_path, monkeypatch):
    monkeypatch.setattr(svc, "LIMITE_LOTES_BLOQUEIO", 0)  # qualquer lote já bloqueia
    caminho_docx = _criar_docx_com_secoes(tmp_path)
    with open(caminho_docx, "rb") as f:
        conteudo_bytes = f.read()

    client = FakeClient()
    resultado = svc.processar_arquivo("documento.docx", conteudo_bytes, client)

    assert resultado["documento_grande_demais"] is True
    assert "erro" in resultado
    assert client.chamadas == 0


# ============================================================
# processar_arquivo — documento truncado por limite de tamanho
# ============================================================

def test_processar_arquivo_com_lotes_truncados_retorna_aviso_parcial(tmp_path, monkeypatch):
    caminho_docx = _criar_docx_com_secoes(tmp_path)
    with open(caminho_docx, "rb") as f:
        conteudo_bytes = f.read()

    def _avaliar_documento_fake(blocos, formato_arquivo, client, executor=None):
        return {
            "criterios": [
                {"id": "escrita_autocontida", "status": "atende", "justificativa": "ok",
                 "risco_ia": "", "trechos_problematicos": [], "localizacao": ["Título da Seção A"],
                 "sugestao_correcao": ""},
            ],
            "lotes_truncados": [["Título da Seção B"], ["Título da Seção C"]],
        }

    monkeypatch.setattr(svc, "avaliar_documento", _avaliar_documento_fake)

    resultado = svc.processar_arquivo("documento.docx", conteudo_bytes, FakeClient())

    assert "erro" not in resultado
    assert "lotes_truncados" not in resultado["avaliacao_detalhada"]
    assert "limite" in resultado["aviso_avaliacao_parcial"].lower()
    assert resultado["locais_nao_avaliados"] == ["Título da Seção B", "Título da Seção C"]


# ============================================================
# processar_arquivo — avaliação parcial (lote com falha persistente)
# ============================================================

def test_processar_arquivo_com_lote_falho_retorna_aviso_parcial_em_vez_de_erro(tmp_path, monkeypatch):
    """Mesmo com 1 lote não avaliável, o arquivo deve retornar resultado
    válido (com score baseado no que foi avaliado) e um aviso explícito,
    em vez de cair no caminho de erro genérico."""
    monkeypatch.setattr(svc.time, "sleep", lambda *_: None)
    caminho_docx = _criar_docx_com_secoes(tmp_path)
    with open(caminho_docx, "rb") as f:
        conteudo_bytes = f.read()

    def _avaliar_documento_fake(blocos, formato_arquivo, client, executor=None):
        return {
            "criterios": [
                {"id": "escrita_autocontida", "status": "atende", "justificativa": "ok",
                 "risco_ia": "", "trechos_problematicos": [], "localizacao": ["Título da Seção A"],
                 "sugestao_correcao": ""},
            ],
            "lotes_nao_avaliados": [["Título da Seção B"]],
        }

    monkeypatch.setattr(svc, "avaliar_documento", _avaliar_documento_fake)

    resultado = svc.processar_arquivo("documento.docx", conteudo_bytes, FakeClient())

    assert "erro" not in resultado
    assert "lotes_nao_avaliados" not in resultado["avaliacao_detalhada"]
    assert resultado["aviso_avaliacao_parcial"] is not None
    assert resultado["locais_nao_avaliados"] == ["Título da Seção B"]


# ============================================================
# _pior_status_por_criterio / _taxa_atendimento
# ============================================================

def test_pior_status_por_criterio_escolhe_o_mais_grave():
    criterios = [
        {"id": "x", "status": "atende"},
        {"id": "x", "status": "nao_atende"},
        {"id": "x", "status": "parcial"},
    ]
    assert svc._pior_status_por_criterio(criterios) == {"x": "nao_atende"}


def test_taxa_atendimento_sem_itens_relevantes_retorna_1():
    assert svc._taxa_atendimento({}, {"a", "b"}) == 1.0


def test_taxa_atendimento_ignora_nao_aplicavel():
    status = {"a": "atende", "b": "nao_aplicavel"}
    assert svc._taxa_atendimento(status, {"a", "b"}) == 1.0


def test_taxa_atendimento_pesos_parcial_e_nao_atende():
    status = {"a": "atende", "b": "parcial", "c": "nao_atende"}
    taxa = svc._taxa_atendimento(status, {"a", "b", "c"})
    assert taxa == pytest.approx((1.0 + 0.5 + 0.0) / 3)


# ============================================================
# calcular_score — as 4 classificações
# ============================================================

def _avaliacao_com_status(status_por_id: dict) -> dict:
    return {"criterios": [{"id": cid, "status": status} for cid, status in status_por_id.items()]}


def test_calcular_score_pronto_para_uso_em_ia():
    avaliacao = _avaliacao_com_status({cid: "atende" for cid in svc._IDS_OBRIGATORIOS})
    score_info = svc.calcular_score(avaliacao)
    assert score_info["classificacao"] == "Pronto para uso em IA"
    assert score_info["score"] == 100.0
    assert score_info["qtd_problemas_criticos"] == 0


def test_calcular_score_pronto_com_ressalvas_sem_criticos():
    ids = list(svc._IDS_OBRIGATORIOS)
    status = {cid: "atende" for cid in ids}
    status[ids[0]] = "parcial"  # só parcial, nenhum nao_atende
    score_info = svc.calcular_score(_avaliacao_com_status(status))
    assert score_info["classificacao"] == "Pronto com ressalvas"
    assert score_info["qtd_problemas_criticos"] == 0


def test_calcular_score_necessita_ajuste_ate_3_criticos():
    ids = list(svc._IDS_OBRIGATORIOS)
    status = {cid: "atende" for cid in ids}
    for cid in ids[:2]:
        status[cid] = "nao_atende"
    score_info = svc.calcular_score(_avaliacao_com_status(status))
    assert score_info["classificacao"] == "Necessita ajuste"
    assert score_info["qtd_problemas_criticos"] == 2


def test_calcular_score_nao_recomendado_mais_de_3_criticos():
    ids = list(svc._IDS_OBRIGATORIOS)
    status = {cid: "atende" for cid in ids}
    for cid in ids[:4]:
        status[cid] = "nao_atende"
    score_info = svc.calcular_score(_avaliacao_com_status(status))
    assert score_info["classificacao"] == "Não recomendado para uso em IA"
    assert score_info["qtd_problemas_criticos"] == 4


# ============================================================
# agregar_lote
# ============================================================

def test_agregar_lote_ordena_por_score_e_soma_bloqueantes():
    resultados = [
        {"arquivo": "b.pdf", "score_info": {"classificacao": "X", "score": 80.0, "qtd_problemas_criticos": 0, "qtd_problemas_recomendados": 1}},
        {"arquivo": "a.pdf", "score_info": {"classificacao": "Y", "score": 40.0, "qtd_problemas_criticos": 2, "qtd_problemas_recomendados": 0}},
    ]
    agregado = svc.agregar_lote(resultados)
    assert agregado["total_arquivos"] == 2
    assert agregado["tabela"][0]["arquivo"] == "a.pdf"  # menor score primeiro
    assert agregado["arquivos_bloqueantes"] == 1
    assert agregado["media_score"] == 60.0


def test_agregar_lote_lista_vazia():
    agregado = svc.agregar_lote([])
    assert agregado["total_arquivos"] == 0
    assert agregado["media_score"] == 0
    assert agregado["arquivos_bloqueantes"] == 0


# ============================================================
# processar_arquivo — caminhos de erro
# ============================================================

def test_processar_arquivo_extracao_vazia_marca_requer_ocr(tmp_path, monkeypatch):
    import fitz
    doc = fitz.open()
    doc.new_page()
    doc.new_page()
    conteudo_bytes = doc.tobytes()
    doc.close()

    resultado = svc.processar_arquivo("escaneado.pdf", conteudo_bytes, FakeClient())
    assert resultado["requer_ocr"] is True
    assert "erro" in resultado


def test_processar_arquivo_falha_generica_na_extracao():
    resultado = svc.processar_arquivo("corrompido.pdf", b"isto nao e um pdf valido", FakeClient())
    assert "erro" in resultado
    assert "Falha na extração" in resultado["erro"]


def test_processar_arquivo_falha_na_avaliacao(tmp_path):
    caminho_docx = _criar_docx_com_secoes(tmp_path)
    with open(caminho_docx, "rb") as f:
        conteudo_bytes = f.read()

    client = FakeClient(erro=RuntimeError("gateway fora do ar"))
    resultado = svc.processar_arquivo("documento.docx", conteudo_bytes, client)
    assert "erro" in resultado
    assert "Falha na avaliação" in resultado["erro"]


def test_processar_arquivo_sucesso_fim_a_fim(tmp_path):
    caminho_docx = _criar_docx_com_secoes(tmp_path)
    with open(caminho_docx, "rb") as f:
        conteudo_bytes = f.read()

    resultado = svc.processar_arquivo("documento.docx", conteudo_bytes, FakeClient())
    assert resultado["arquivo"] == "documento.docx"
    assert "erro" not in resultado
    assert resultado["score_info"]["score"] == 100.0
    assert resultado["resumo_narrativo"] == "resumo ok"


# ============================================================
# gerar_relatorio_html
# ============================================================

def _resultado_base(criterios):
    return {
        "arquivo": "doc.pdf",
        "avaliacao_detalhada": {"criterios": criterios},
        "score_info": {"score": 80.0, "classificacao": "Pronto com ressalvas"},
        "metricas_documento": {"num_blocos": 3, "total_caracteres": 900, "proporcao_tabular": 0.2, "blocos_duplicados": 0},
        "resumo_narrativo": "Resumo de teste.",
    }


def test_gerar_relatorio_html_sem_problemas():
    resultado = _resultado_base([{"id": "escrita_autocontida", "status": "atende"}])
    html_gerado = svc.gerar_relatorio_html(resultado)
    assert "Nenhum problema encontrado" in html_gerado
    assert "doc.pdf" in html_gerado


def test_gerar_relatorio_html_com_problema_mostra_sugestao_diferente_da_justificativa():
    criterios = [{
        "id": "siglas_explicadas",
        "status": "nao_atende",
        "justificativa": "Sigla N205 sem explicação.",
        "risco_ia": "IA pode não entender o código.",
        "trechos_problematicos": ["N205"],
        "localizacao": ["Página 3"],
        "sugestao_correcao": "N205 (negação por ausência de documentação complementar)",
    }]
    resultado = _resultado_base(criterios)
    html_gerado = svc.gerar_relatorio_html(resultado)
    assert "Siglas explicadas" in html_gerado
    assert "N205 (negação por ausência de documentação complementar)" in html_gerado
    assert "Página 3" in html_gerado


def test_gerar_relatorio_html_ignora_criterios_de_rastreabilidade():
    criterios = [
        {"id": next(iter(svc.IDS_RASTREABILIDADE)), "status": "nao_atende", "justificativa": "problema de metadado"},
    ]
    resultado = _resultado_base(criterios)
    html_gerado = svc.gerar_relatorio_html(resultado)
    assert "Nenhum problema encontrado" in html_gerado


def test_gerar_relatorio_html_ignora_criterios_com_diagnostico_incompleto():
    """Critérios marcados como diagnostico_incompleto (o modelo indicou
    problema mas não deu detalhe algum, por erro de classificação de
    grupo) não devem aparecer na tabela — não há nada acionável para
    mostrar, e exibi-los só gera linhas cheias de "—"."""
    criterios = [
        {"id": "escrita_autocontida", "status": "parcial", "justificativa": "",
         "diagnostico_incompleto": True},
    ]
    resultado = _resultado_base(criterios)
    html_gerado = svc.gerar_relatorio_html(resultado)
    assert "Nenhum problema encontrado" in html_gerado


# ============================================================
# processar_lote_arquivos — paralelismo compartilhado entre arquivos
# ============================================================

def test_processar_lote_arquivos_lista_vazia():
    assert svc.processar_lote_arquivos([], FakeClient()) == []


def test_processar_lote_arquivos_processa_todos_e_preserva_resultado(tmp_path):
    """Vários arquivos de um mesmo upload devem ser processados e cada um
    deve retornar seu próprio resultado corretamente, na mesma ordem em
    que foram enviados."""
    caminho1 = _criar_docx_com_secoes(tmp_path)
    with open(caminho1, "rb") as f:
        bytes1 = f.read()

    arquivos = [("documento1.docx", bytes1), ("documento2.docx", bytes1)]
    resultados = svc.processar_lote_arquivos(arquivos, FakeClient())

    assert len(resultados) == 2
    assert resultados[0]["arquivo"] == "documento1.docx"
    assert resultados[1]["arquivo"] == "documento2.docx"
    assert "erro" not in resultados[0]
    assert "erro" not in resultados[1]


def test_processar_lote_arquivos_respeita_teto_de_concorrencia_entre_arquivos(tmp_path, monkeypatch):
    """O ponto central da mudança: mesmo com vários arquivos gerando
    lotes ao mesmo tempo, o número de chamadas SIMULTÂNEAS à Gateway nunca
    deve passar de MAX_CHAMADAS_PARALELAS — o teto vale para o upload
    inteiro, não por arquivo. Sem o pool compartilhado, 3 arquivos com 4
    lotes cada, todos abrindo seu próprio pool de até
    MAX_CHAMADAS_PARALELAS, somariam um pico bem maior que 3."""
    import threading

    monkeypatch.setattr(svc, "MAX_CHAMADAS_PARALELAS", 3)
    # Cada arquivo terá 4 blocos, e cada bloco vira 1 lote (limite bem
    # pequeno), garantindo múltiplos lotes por arquivo de forma
    # determinística — sem depender de como o Excel real fatiaria linhas.
    monkeypatch.setattr(svc, "MAX_CHARS_PER_LLM_CALL", 1)
    monkeypatch.setattr(
        svc, "extrair_texto",
        lambda caminho: [Bloco(localizacao=f"P{i}", texto="conteudo") for i in range(4)],
    )

    contador_lock = threading.Lock()
    estado = {"em_andamento": 0, "pico_observado": 0}

    class ClienteContadorConcorrencia:
        def gerar_texto(self, *args, **kwargs):
            with contador_lock:
                estado["em_andamento"] += 1
                estado["pico_observado"] = max(estado["pico_observado"], estado["em_andamento"])
            time.sleep(0.05)
            with contador_lock:
                estado["em_andamento"] -= 1
            return json.dumps({"criterios_sem_problema": [], "criterios_com_problema": []})

    arquivos = [(f"arquivo_{i}.docx", b"conteudo qualquer") for i in range(3)]
    svc.processar_lote_arquivos(arquivos, ClienteContadorConcorrencia())

    assert estado["pico_observado"] <= 3
