"""
curadoria_service.py

Este é o coração do agente. É aqui que o documento enviado (PDF, Word,
Excel ou PowerPoint) é lido, avaliado contra a rubrica de qualidade da
Governança de IA, tem uma nota calculada, e vira um relatório em HTML.
Resumindo em 3 passos: 1) extrai o texto do arquivo sem usar IA, 2)
manda esse texto para o modelo avaliar critério por critério, 3)
transforma essa avaliação numa nota e numa classificação final.
"""

import concurrent.futures
import html
import json
import os
import re
import time
import traceback
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from src.services.gateway_apim_service import GatewayClient

CURATION_MODEL = os.getenv("CURATION_MODEL", "deployment-gpt-4.1-mini")
MAX_CHARS_PER_LLM_CALL = int(os.getenv("MAX_CHARS_PER_LLM_CALL", "60000"))
# Documento grande = muitos lotes = muitas chamadas sequenciais à Gateway.
# Sem retry, bastava 1 lote falhar (timeout pontual, rede instável) para
# o arquivo inteiro ser descartado. Com retry, uma falha isolada em 1
# lote não derruba os outros 20-30 que funcionaram.
TENTATIVAS_POR_LOTE = int(os.getenv("TENTATIVAS_POR_LOTE", "3"))
ESPERA_ENTRE_TENTATIVAS_SEGUNDOS = int(os.getenv("ESPERA_ENTRE_TENTATIVAS_SEGUNDOS", "5"))

# Cada lote = 1 chamada (com retry) ao workflow da Gateway, com custo de
# tokens fixo (todos os 24 critérios + few-shot examples) somado a cada
# chamada. As chamadas são feitas em paralelo (ver _MAX_CHAMADAS_PARALELAS)
# até o limite de concorrência que a própria Gateway suporta — hoje 10
# requisições simultâneas por vez. Sem limite de lotes, um PDF de 5000
# páginas ou um Excel de 1 milhão de linhas geraria centenas ou milhares de
# lotes — mesmo em paralelo, isso ultrapassaria uma meta de tempo de
# resposta de ~10 minutos por documento.
#
# Conta usada para calibrar LIMITE_LOTES_PROCESSAVEIS: orçamento de 600s
# (10 min), reservando ~60s para a chamada final do resumo narrativo =
# 540s para a fase de avaliação dos lotes. Com paralelismo de 10 e ~90s de
# latência média por lote (geração + polling + eventual retry), o tempo
# total é aproximadamente (n_lotes / 10) * 90s. Resolvendo para 540s:
# n_lotes ≈ 60. Se o tempo real por lote ou o limite de paralelismo da
# Gateway mudar, esse número deve ser recalculado.
#
# Acima de LIMITE_LOTES_PROCESSAVEIS, avalia só os primeiros lotes e avisa
# que o resto foi truncado (nunca corta um lote no meio, só omite lotes
# inteiros do fim). Acima de LIMITE_LOTES_BLOQUEIO (3x o processável, ou
# seja, cobertura mínima de ~33% do documento), recusa o arquivo inteiro
# ANTES de qualquer chamada à Gateway (custo de IA = zero nesse caso) — a
# extração de texto que gera os blocos é local e sem custo, então dá para
# saber quantos lotes seriam necessários sem gastar 1 token.
LIMITE_LOTES_PROCESSAVEIS = int(os.getenv("LIMITE_LOTES_PROCESSAVEIS", "60"))
LIMITE_LOTES_BLOQUEIO = int(os.getenv("LIMITE_LOTES_BLOQUEIO", "180"))

# Máximo de chamadas simultâneas ao workflow da Gateway. A própria Gateway
# suporta no máximo 10 requisições em paralelo por vez — nunca deixamos
# passar disso, mesmo que MAX_CHAMADAS_PARALELAS seja configurado com um
# valor maior por engano.
_LIMITE_CONCORRENCIA_GATEWAY = 10
MAX_CHAMADAS_PARALELAS = min(
    int(os.getenv("MAX_CHAMADAS_PARALELAS", "10")), _LIMITE_CONCORRENCIA_GATEWAY
)


class DocumentoMuitoGrandeError(Exception):
    """Levantado quando o documento geraria mais lotes de avaliação do que
    LIMITE_LOTES_BLOQUEIO permite. Levantado ANTES de qualquer chamada ao
    modelo — recusar aqui custa zero tokens."""
MAX_TOKENS_AVALIACAO = int(os.getenv("MAX_TOKENS_AVALIACAO", "8192"))
MAX_TOKENS_RESUMO = int(os.getenv("MAX_TOKENS_RESUMO", "1024"))
GATEWAY_CHUNK_SIZE_TOKENS = int(os.getenv("GATEWAY_CHUNK_SIZE_TOKENS", "800"))
GATEWAY_CHUNK_OVERLAP_TOKENS = int(os.getenv("GATEWAY_CHUNK_OVERLAP_TOKENS", "300"))
CHARS_POR_TOKEN_APROX = int(os.getenv("CHARS_POR_TOKEN_APROX", "4"))
LIMITE_CHUNKS_POR_DOCUMENTO_RECOMENDACAO = int(os.getenv("LIMITE_CHUNKS_POR_DOCUMENTO_RECOMENDACAO", "150"))
LIMITE_LINHAS_EXCEL_RECOMENDACAO = int(os.getenv("LIMITE_LINHAS_EXCEL_RECOMENDACAO", "10000"))

# Mesmo tamanho de lote (100 linhas) usado pela indexação real da Gateway
# para Excel — ver _extrair_excel() para o motivo de existir.
LINHAS_POR_BLOCO_EXCEL = int(os.getenv("LINHAS_POR_BLOCO_EXCEL", "100"))


# ============================================================
# EXTRAÇÃO LOCAL DE TEXTO (sem IA, sem embeddings, sem indexação)
# ============================================================

@dataclass
class Bloco:
    localizacao: str 
    texto: str


class ExtracaoVaziaError(Exception):
    """Levantado quando não foi possível extrair texto (ex: PDF escaneado)."""


def extrair_texto(caminho_arquivo: str) -> list[Bloco]:
    ext = Path(caminho_arquivo).suffix.lower()
    if ext == ".pdf":
        blocos = _extrair_pdf(caminho_arquivo)
    elif ext == ".docx":
        blocos = _extrair_docx(caminho_arquivo)
    elif ext in (".xlsx", ".xlsm"):
        blocos = _extrair_excel(caminho_arquivo)
    elif ext == ".pptx":
        blocos = _extrair_pptx(caminho_arquivo)
    else:
        raise ValueError(f"Formato não suportado: {ext}")

    return blocos


def _extrair_pdf(caminho: str) -> list[Bloco]:
    """1 bloco por página com texto. Páginas em branco não viram bloco."""
    import fitz 

    blocos, total_chars = [], 0
    doc = fitz.open(caminho)
    for i, pagina in enumerate(doc, start=1):
        texto = pagina.get_text().strip()
        total_chars += len(texto)
        if texto:
            blocos.append(Bloco(localizacao=f"Página {i}", texto=texto))
    n_paginas = len(doc)
    doc.close()

 
    if n_paginas > 0 and total_chars < 20 * n_paginas:
        raise ExtracaoVaziaError(
            "PDF aparenta estar escaneado (pouco ou nenhum texto extraído). "
            "Necessário OCR (Azure Document Intelligence) antes da curadoria."
        )
    return blocos


def _extrair_docx(caminho: str) -> list[Bloco]:
    """1 bloco por seção (do título "Heading" ou "Título" até o próximo
    título) e mais 1 bloco por tabela. Word não tem conceito de página
    no arquivo, só de estilo de parágrafo, por isso agrupamos por
    título em vez de por página."""
    import docx

    doc = docx.Document(caminho)
    partes, secao_atual, buffer = [], "Início do documento", []

    def flush():
        # Fecha a seção atual e salva o texto acumulado como 1 bloco.
        texto = "\n".join(buffer).strip()
        if texto:
            partes.append(Bloco(localizacao=secao_atual, texto=texto))

    for paragrafo in doc.paragraphs:
        estilo = (paragrafo.style.name or "").lower()
        if estilo.startswith("heading") or estilo.startswith("título"):
            flush()
            buffer.clear()
            secao_atual = paragrafo.text.strip() or secao_atual
        elif paragrafo.text.strip():
            buffer.append(paragrafo.text.strip())
    flush()

    for idx, tabela in enumerate(doc.tables, start=1):
        linhas = [" | ".join(cell.text.strip() for cell in row.cells) for row in tabela.rows]
        if linhas:
            partes.append(Bloco(localizacao=f"Tabela {idx}", texto="\n".join(linhas)))
    return partes


def _extrair_excel(caminho: str) -> list[Bloco]:
    """1 bloco por aba, a menos que a aba tenha mais linhas que
    LINHAS_POR_BLOCO_EXCEL — nesse caso a aba é dividida em vários
    blocos de LINHAS_POR_BLOCO_EXCEL linhas cada, seguindo o mesmo
    padrão de loteamento fixo (100 em 100 linhas) que a própria
    plataforma Gateway usa na indexação de Excel. Sem isso, uma aba com
    milhares de linhas (ex: tabela TUSS) virava 1 bloco só e gigante,
    que _dividir_em_lotes não conseguia quebrar (bloco maior que o
    limite fica sozinho), estourando o payload enviado à Gateway e
    derrubando o workflow (WF_COMPLETED_ERROR) antes de sequer avaliar
    o conteúdo. Cada linha da planilha vira uma linha de texto dentro
    do bloco, com as colunas separadas por " | ". data_only=True lê o
    valor calculado da célula, não a fórmula."""
    import openpyxl

    wb = openpyxl.load_workbook(caminho, data_only=True)
    blocos = []
    for aba in wb.sheetnames:
        ws = wb[aba]
        linhas_texto = []
        for row in ws.iter_rows(values_only=True):
            valores = [str(c) if c is not None else "" for c in row]
            if any(v.strip() for v in valores):  # pula linha totalmente vazia
                linhas_texto.append(" | ".join(valores))

        if not linhas_texto:
            continue

        if len(linhas_texto) <= LINHAS_POR_BLOCO_EXCEL:
            blocos.append(Bloco(localizacao=f"Aba: {aba}", texto="\n".join(linhas_texto)))
            continue

        for inicio in range(0, len(linhas_texto), LINHAS_POR_BLOCO_EXCEL):
            lote_linhas = linhas_texto[inicio:inicio + LINHAS_POR_BLOCO_EXCEL]
            fim = inicio + len(lote_linhas)
            localizacao = f"Aba: {aba} (linhas {inicio + 1}-{fim})"
            blocos.append(Bloco(localizacao=localizacao, texto="\n".join(lote_linhas)))
    return blocos


def _extrair_pptx(caminho: str) -> list[Bloco]:
    """1 bloco por slide com conteúdo. Marca no texto quando o slide tem
    imagem, pra rubrica avaliar se ela é só decorativa ou informativa."""
    from pptx import Presentation

    prs = Presentation(caminho)
    blocos = []
    for i, slide in enumerate(prs.slides, start=1):
        textos, tem_imagem = [], False
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                textos.append(shape.text_frame.text.strip())
            if shape.shape_type == 13:  # 13 = tipo "Picture" no python-pptx
                tem_imagem = True
        conteudo = "\n".join(textos)
        if tem_imagem:
            conteudo += "\n[SLIDE CONTÉM IMAGEM — verificar se é informativa ou decorativa]"
        if conteudo.strip():
            blocos.append(Bloco(localizacao=f"Slide {i}", texto=conteudo))
    return blocos


def _concatenar_blocos(blocos: list[Bloco]) -> str:
    return "\n\n".join(f"### [{b.localizacao}]\n{b.texto}" for b in blocos)


REFERENCIA_CHUNK_SIZE = GATEWAY_CHUNK_SIZE_TOKENS * CHARS_POR_TOKEN_APROX

_REGEX_CARACTERES_RUIDO = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]|\s{4,}")


def calcular_metricas_documento(blocos: list[Bloco]) -> dict:
    num_blocos = len(blocos)
    tamanhos = [len(b.texto) for b in blocos]
    total_chars = sum(tamanhos)
    tamanho_medio = round(total_chars / num_blocos, 1) if num_blocos else 0

    # 1. Fragmentação. Muitos blocos pequenos indicam título ou frase
    # solta em vez de parágrafo coeso.
    blocos_curtos = sum(1 for t in tamanhos if t < 120)
    proporcao_blocos_curtos = round(blocos_curtos / num_blocos, 2) if num_blocos else 0
    alerta_fragmentacao = num_blocos >= 4 and proporcao_blocos_curtos >= 0.5

    # 2. Ruído estrutural. Caracteres de controle ou espaço em excesso
    # são sinal de extração suja (ex: OCR malformado).
    caracteres_ruido = sum(len(m) for b in blocos for m in _REGEX_CARACTERES_RUIDO.findall(b.texto))
    proporcao_ruido = round(caracteres_ruido / total_chars, 3) if total_chars else 0

    # 3. Proporção tabela e texto. Mede quanto do documento é tabela e
    # quanto é texto corrido.
    chars_tabulares = sum(
        len(b.texto) for b in blocos
        if b.localizacao.startswith("Tabela") or b.localizacao.startswith("Aba:")
    )
    proporcao_tabular = round(chars_tabulares / total_chars, 2) if total_chars else 0

    # 4. Repetição interna. Blocos com texto idêntico (já normalizado)
    # em mais de um lugar do mesmo documento sinalizam boilerplate
    # repetido ou conteúdo duplicado dentro do próprio arquivo.
    normalizados = [" ".join(b.texto.lower().split()) for b in blocos if len(b.texto.strip()) > 20]
    contagem = Counter(normalizados)
    blocos_duplicados = sum(c - 1 for c in contagem.values() if c > 1)

    # 5. Compatibilidade com o chunk_size real do indexador.
    if tamanho_medio == 0:
        compatibilidade_chunking = "sem_dados"
    elif tamanho_medio < REFERENCIA_CHUNK_SIZE * 0.25:
        compatibilidade_chunking = "blocos_muito_pequenos"
    elif tamanho_medio > REFERENCIA_CHUNK_SIZE * 3:
        compatibilidade_chunking = "blocos_muito_grandes"
    else:
        compatibilidade_chunking = "adequado"

    return {
        "num_blocos": num_blocos,
        "total_caracteres": total_chars,
        "tamanho_medio_bloco": tamanho_medio,
        "proporcao_blocos_curtos": proporcao_blocos_curtos,
        "alerta_fragmentacao": alerta_fragmentacao,
        "proporcao_ruido": proporcao_ruido,
        "proporcao_tabular": proporcao_tabular,
        "blocos_duplicados": blocos_duplicados,
        "compatibilidade_chunking": compatibilidade_chunking,
    }


# Porte do documento, usado só para recomendar divisão quando o arquivo
# é grande. Nunca bloqueia nem trunca nada: a avaliação sempre cobre o
# documento inteiro via map-reduce (ver avaliar_documento).
_AVANCO_LIQUIDO_CHARS_POR_CHUNK = max(
    1, (GATEWAY_CHUNK_SIZE_TOKENS - GATEWAY_CHUNK_OVERLAP_TOKENS) * CHARS_POR_TOKEN_APROX
)


def _unidade_nativa(blocos: list[Bloco], ext: str) -> tuple[str, int]:
    """Conta o porte do documento na unidade que faz sentido pra quem não
    é técnico. PDF e PowerPoint geram um bloco por página/slide com
    conteúdo. Excel não tem um bloco por linha (cada bloco é uma aba
    inteira), então as linhas são contadas dentro do texto de cada aba.
    Word (.docx) não tem contagem nativa de página, então usa total de
    caracteres."""
    if ext == ".pdf":
        return "páginas", len(blocos)
    if ext == ".pptx":
        return "slides", len(blocos)
    if ext in (".xlsx", ".xlsm"):
        linhas = sum(b.texto.count("\n") + 1 for b in blocos if b.localizacao.startswith("Aba:"))
        return "linhas", linhas
    return "caracteres", sum(len(b.texto) for b in blocos)


def calcular_porte_documento(blocos: list[Bloco], ext: str) -> dict:
    """Estima quantos chunks o documento geraria no índice real da Gateway
    (critério que decide se é "grande demais") e conta também a unidade
    nativa (páginas/slides/linhas) só para exibição na recomendação."""
    unidade, quantidade = _unidade_nativa(blocos, ext)

    if ext in (".xlsx", ".xlsm"):
        num_chunks_estimado = quantidade  # 1 linha = 1 chunk
        limite = LIMITE_LINHAS_EXCEL_RECOMENDACAO
    else:
        total_chars = sum(len(b.texto) for b in blocos)
        num_chunks_estimado = -(-total_chars // _AVANCO_LIQUIDO_CHARS_POR_CHUNK)  # ceil sem importar math
        limite = LIMITE_CHUNKS_POR_DOCUMENTO_RECOMENDACAO

    return {
        "unidade": unidade,
        "quantidade": quantidade,
        "num_chunks_estimado": num_chunks_estimado,
        "limite_chunks_recomendado": limite,
        "documento_grande": num_chunks_estimado > limite,
    }


CRITERIOS_OBRIGATORIOS = [
    {"id": "escrita_autocontida", "descricao": "Frases e parágrafos fazem sentido sozinhos, sem depender de contexto externo ao trecho. Sem referências vagas ('acima', 'abaixo', 'conforme mencionado', 'o item anterior') sem nomear explicitamente o que é referenciado."},
    {"id": "paragrafos_coesos", "descricao": "Não há fragmentação excessiva: título seguido de frases soltas em linhas separadas que deveriam formar um único parágrafo coeso. Ideias relacionadas estão unidas, não quebradas em múltiplas linhas sem necessidade."},
    {"id": "titulos_descritivos", "descricao": "Títulos e subtítulos de seção são específicos e contêm palavras-chave do conteúdo (ex: 'Prazo de Análise de Reembolso'), não genéricos (ex: 'Seção 3', 'Informações Gerais')."},
    {"id": "conteudo_contraditorio_tratado", "descricao": "Não há regras conflitantes apresentadas lado a lado sem indicação clara de qual prevalece ou em que situação cada uma se aplica."},
    {"id": "textos_revogados_tratados", "descricao": "Não há texto tachado sem tratamento nem conteúdo antigo misturado ao vigente sem sinalização clara (ex: '[REVOGADO]' / '[REGRA ATUAL]', ou 'válido até'/'válido a partir de')."},
    {"id": "cores_convertidas_em_texto", "descricao": "Nenhuma informação depende apenas de cor, marca-texto ou destaque visual para ser entendida; todo significado relevante também está escrito em texto."},
    {"id": "imagens_tratadas_corretamente", "descricao": "Imagens, prints e fluxos que carregam informação (não decorativos) têm descrição textual próxima. Imagens puramente decorativas ou de layout (logos, ícones, linhas divisórias) não comprometem a leitura caso estejam presentes."},
    {"id": "tabelas_revisadas", "descricao": "Tabelas são legíveis, sem células mescladas, sem células vazias sem explicação ('não se aplica', 'não informado'), sem abreviações não explicadas nos cabeçalhos, e não excessivamente complexas ou densas para o formato. CABEÇALHOS DEVEM SER COMPLETOS E ESPECÍFICOS — regra geral, aplicável a QUALQUER cabeçalho, não só aos exemplos citados aqui: um cabeçalho de uma palavra só e genérica (ex: 'Código', 'Descrição', 'Status', 'Nome', 'Valor', 'Data', 'Tipo', 'Área', 'Responsável', entre outros) sem dizer do que se trata é um problema mesmo não sendo sigla ou abreviação. Teste a aplicar em CADA cabeçalho da tabela: uma pessoa sem nenhum contexto além do próprio cabeçalho entenderia exatamente o que a coluna representa? Se não, o cabeçalho precisa ser mais específico (ex: 'Código da Regra' em vez de só 'Código'; 'Data de Vigência' em vez de só 'Data'; 'Área Responsável pela Aprovação' em vez de só 'Área'). REGRA ESPECÍFICA PARA ARQUIVOS EXCEL (.xlsx/.xlsm): os cabeçalhos de coluna precisam estar sem espaços (usar underscore '_' no lugar de espaço) e sem acentuação — esse é um requisito técnico do formato de indexação da plataforma Gateway, não apenas uma boa prática de clareza. Em tabelas de outros formatos (PDF, Word, PowerPoint), essa regra técnica NÃO se aplica — o que importa ali é só a clareza do cabeçalho (pode ter espaço e acento normalmente)."},
    {"id": "siglas_explicadas", "descricao": "Siglas, códigos e abreviações têm o significado explicado na primeira ocorrência (ou próximo dela), não apenas em um glossário separado e distante."},
    {"id": "densidade_informativa", "descricao": "O texto é direto e objetivo, sem repetição desnecessária da mesma informação em frases diferentes só para preencher espaço. Cada frase agrega algo novo."},
    {"id": "linguagem_informal_evitada", "descricao": "Não há expressões de conversa, reunião ou mensagem interna (ex: 'conforme falado', 'como combinado', 'verificar com o fulano', 'validar depois', 'regra antiga/nova') sem substituição por linguagem formal e datada (ex: 'conforme regra vigente desde 01/01/2026')."},
    {"id": "regras_de_negocio_completas", "descricao": "Regras de negócio não têm ambiguidade real que impediria a aplicação correta. Nem toda regra precisa mencionar prazo, exceção, documento comprobatório E área responsável — avalie caso a caso quais desses elementos são relevantes para aquela regra específica. Só marque como problema o elemento que está genuinamente faltando e geraria dúvida real (ex: 'entrar em contato com o setor responsável' sem dizer qual setor é isso; 'o prazo pode ser prorrogado em casos excepcionais' sem dizer quais casos). Não exija a lista completa de elementos em toda regra."},
    {"id": "notas_incorporadas_ao_texto", "descricao": "Informações relevantes (regras, exceções, prazos, condições) não estão apenas em comentários, rodapés, notas soltas, caixas laterais ou anotações fora do fluxo principal do texto ou da tabela — estão incorporadas ao conteúdo principal, próximas da regra a que se referem."},
    {"id": "termos_padronizados", "descricao": "O mesmo conceito é representado sempre com o mesmo termo ou status ao longo do documento (ex: não varia entre 'OK', 'Aprovada', 'Validado' e 'Sim' para dizer a mesma coisa em lugares diferentes)."},
]

CRITERIOS_DESEJAVEIS = [
    {"id": "documento_vigente_explicito", "descricao": "O documento indica explicitamente que está na versão vigente/atual (não apenas implícito pela ausência de menção contrária)."},
    {"id": "responsavel_definido", "descricao": "Existe área ou pessoa responsável pelo conteúdo identificada no próprio documento (cabeçalho, rodapé ou seção de autoria)."},
    {"id": "data_atualizacao", "descricao": "A data da última atualização está identificada no documento."},
    {"id": "versao_identificada", "descricao": "A versão do documento está claramente identificada (ex: 'v2.1', 'Revisão 3')."},
    {"id": "status_documento", "descricao": "O documento possui status explícito (vigente, obsoleto, em revisão etc.)."},
    {"id": "historico_alteracoes", "descricao": "O documento possui registro das principais mudanças (changelog)."},
    {"id": "faq", "descricao": "O documento possui seção de perguntas frequentes, quando aplicável ao tipo de conteúdo."},
    {"id": "exemplos_aplicacao", "descricao": "O documento apresenta exemplos práticos das regras descritas."},
    {"id": "metadados_enriquecidos", "descricao": "O documento possui categoria, processo, público-alvo, vigência e área responsável explícitos."},
    {"id": "revisao_periodica_definida", "descricao": "Existe frequência de revisão prevista para o conteúdo."},
]

_FEW_SHOT_EXAMPLES = """
Exemplo 1 — Ambiguidade / referência vaga (critério: escrita_autocontida)
Ruim: "Seguir regra anterior, exceto nos casos novos."
Recomendado: "Para solicitações abertas até 31/12/2025, deve ser aplicada a
regra anterior. Para solicitações abertas a partir de 01/01/2026, deve ser
aplicada a nova regra descrita na seção 4."

Exemplo 2 — Cor sem explicação textual (critério: cores_convertidas_em_texto)
Ruim: um trecho aparece apenas destacado em amarelo, sem legenda no texto.
Recomendado: "[Categoria: Sinistro] O cliente deve apresentar a documentação
obrigatória para abertura do processo."

Exemplo 3 — Tabela mal formatada (critério: tabelas_revisadas)
Ruim: colunas "Cod | Desc | Obs" com abreviações e células como "Ver regra".
Recomendado: colunas "Código da regra | Descrição da regra | Observação
complementar", com texto completo em cada célula.

Exemplo 4 — Sigla sem definição (critério: siglas_explicadas)
Ruim: "N205" aparece no texto sem nenhuma explicação próxima.
Recomendado: "N205 (código de negação por ausência de documentação
complementar)".

Exemplo 5 — Linguagem informal (critério: linguagem_informal_evitada)
Ruim: "conforme falado", "validar depois", "regra antiga", "aquele caso".
Recomendado: "conforme regra vigente desde 01/01/2026", "validar com a área
responsável pelo processo", "regra válida até 31/12/2025".

Exemplo 6 — Regra de negócio incompleta (critério: regras_de_negocio_completas)
Ruim: "Solicitar nota fiscal quando necessário."
Recomendado: "A nota fiscal deve ser solicitada quando o documento apresentado
não possuir as informações fiscais obrigatórias para validação do reembolso.
Quando a informação já constar na base oficial, a solicitação da nota fiscal
poderá ser dispensada, conforme regra vigente da área responsável."

Exemplo 7 — Nota/observação fora do fluxo principal (critério: notas_incorporadas_ao_texto)
Ruim: regra no texto principal ("O prazo de análise é de 5 dias úteis.") e uma
nota solta no rodapé ("Para solicitações com pendência, o prazo fica suspenso.").
Recomendado: "O prazo de análise é de 5 dias úteis. Para solicitações com
pendência documental, esse prazo fica suspenso até o recebimento da
documentação completa."

Exemplo 8 — Termos inconsistentes para o mesmo status (critério: termos_padronizados)
Ruim: a mesma condição de "aprovado" aparece como "OK", "Aprovada", "Validado"
e "Sim" em linhas diferentes da mesma tabela.
Recomendado: usar sempre o mesmo termo ("Aprovado") para o mesmo significado
em todo o documento.

Exemplo 9 — Como escrever o campo "risco_ia" (aplica-se a qualquer critério)
Para o problema do Exemplo 1 (referência vaga "exceto nos casos novos"):
risco_ia: "A IA pode não saber quando aplicar a regra nova, e responder
com o prazo antigo mesmo para casos que já mudaram."
Para o problema do Exemplo 4 (sigla "N205" sem explicação):
risco_ia: "A IA pode não entender o que significa o código e deixar de
mencionar essa negativa numa resposta sobre o assunto."

Exemplo 10 — Recomendação com correção JÁ APLICADA, apenas 1 exemplo (critério: cores_convertidas_em_texto)
Documento real: "Observação: As críticas grifadas de amarelo são
utilizadas pela equipe do RSMH; As críticas grifadas de vermelho são
utilizadas pela equipe da Enfermagem; As críticas grifadas de azul são
utilizadas pela equipe da Consultoria."
Ruim (sugestão genérica, não mostra como aplicar): "Incluir a informação
das cores também em texto, adicionando etiquetas textuais junto às
críticas para indicar a equipe responsável."
Ruim também (exagera reescrevendo todos os casos): mostrar a
transformação para amarelo, vermelho E azul — não é necessário.
Recomendado (1 exemplo aplicado, seguindo o padrão de tag do Guia, para a
pessoa replicar nos demais casos): "Substituir a marcação por cor por uma
tag textual antes de cada crítica, por exemplo: '[Equipe: RSMH] Crítica
X.' em vez de apenas grifar de amarelo. Aplicar o mesmo padrão para as
demais equipes (Enfermagem, Consultoria)."

Exemplo 11 — Regra de negócio: apontar só o que falta de verdade, não a
lista inteira (critério: regras_de_negocio_completas)
Ruim (exige tudo sempre): "Incluir informações completas sobre aplicação,
exceções, prazos, documentos necessários e responsáveis para cada regra."
— isso é genérico e pode não fazer sentido se a regra já é clara e
simplesmente não tem prazo ou exceção.
Recomendado (aponta a lacuna real e específica): se a regra diz "entrar em
contato com o setor responsável" sem dizer qual setor, a recomendação
deve ser: "Especificar qual é o setor responsável (ex: 'entrar em contato
com o setor de Atendimento ao Beneficiário') em vez de deixar genérico."
Se a regra já é clara sobre o quê e quando, mas realmente não tem prazo
aplicável ao caso, isso NÃO é um problema — não invente uma exigência de
prazo que não faz sentido para aquela regra específica.

Exemplo 12 — Regra de cabeçalho de tabela: só vale para Excel (critério: tabelas_revisadas)
FORMATO DO ARQUIVO: Excel (.xlsx)
Ruim: cabeçalho de coluna "Código da Regra" (com espaço e acento).
Recomendado: "Codigo_da_Regra" (sem espaço, sem acento, usando underscore) —
isso é obrigatório para arquivos Excel, por exigência técnica da
plataforma Gateway de indexação.

FORMATO DO ARQUIVO: PDF
Mesma tabela, cabeçalho claro e completo, exemplo "Código da Regra" (com espaço e acento) —
aqui isso NÃO é problema, porque a regra de underscore/sem acento é
específica de Excel. Em PDF, Word ou PowerPoint, o cabeçalho só precisa
ser claro e descritivo, pode ter espaço e acentuação normalmente.

""".strip()

_SYSTEM_PROMPT = """Você é um especialista em curadoria de documentos para uso em
agentes de IA Generativa com RAG (Retrieval-Augmented Generation).

Sua função é avaliar um documento contra uma lista fixa de critérios de
qualidade documental, definidos pela área de Governança de IA da empresa.
Você NÃO responde perguntas sobre o conteúdo do documento, NÃO resume o
documento e NÃO opina sobre o mérito das regras de negócio nele contidas.
Sua única tarefa é apontar problemas de FORMA (estrutura, clareza,
completude, rastreabilidade) que prejudicariam a leitura desse documento
por um mecanismo de busca e por um modelo de linguagem.

IMPORTANTE sobre rigor e visão holística: avalie o documento como um todo,
nunca uma frase isolada fora do contexto ao redor dela. Só marque
"nao_atende" ou "parcial" se o problema REALMENTE comprometeria uma
resposta do agente de IA (informação errada, ambígua ou impossível de
recuperar) — não marque só porque uma frase, isolada, poderia em teoria
ser reescrita de forma diferente. Isso é especialmente importante nestes
três casos, que são erros comuns de excesso de rigor:

1. Listas bem formadas NÃO são fragmentação. Uma frase de introdução,
   seguida de itens em lista, seguida de uma frase de fechamento, é uma
   estrutura clara e válida — não marque "paragrafos_coesos" como
   problema só porque o conteúdo está organizado em lista. Exemplo que
   NÃO é problema: "O primeiro passo é definir X: • Opção A; • Opção B;
   • Opção C. Cada escolha leva a um caminho específico." Isso é uma
   lista coesa e completa, não frases soltas.

2. Estruturas condicionais completas ("Se X, então Y. Se não X, então
   Z.") são autocontidas, não são referências vagas. Só marque
   "escrita_autocontida" como problema quando a referência realmente não
   diz o que está sendo referenciado (ex: "conforme mencionado acima",
   "isso", "aquele caso" sem dizer o quê). "Se o usuário confirma, o
   fluxo segue para X. Se o usuário recusa, o sistema informa Y." NÃO é
   vago — as duas condições e os dois resultados estão explícitos.

3. Diferenças de tom/formalidade não são necessariamente um problema.
   Só marque "linguagem_informal_evitada" quando a expressão realmente
   parecer uma nota pessoal/interna que só faz sentido para quem
   participou da conversa original (ex: "conforme combinado com o
   fulano", "validar depois", "aquele caso de ontem") — não quando é
   apenas uma frase direta e imperativa mas ainda clara e profissional
   (ex: "faça sua configuração pelo aplicativo" é direto, não informal
   a ponto de confundir a IA ou o leitor).

4. Para "regras_de_negocio_completas", NÃO exija a lista inteira (o quê,
   quando se aplica, quando não se aplica, prazo, exceção, documento,
   área responsável) em toda regra do documento — nem toda regra precisa
   de todos esses elementos. Avalie regra por regra e aponte APENAS o
   elemento que está genuinamente faltando e geraria dúvida real. Exemplo
   de problema real: a regra diz "entrar em contato com o setor
   responsável" sem dizer qual setor — isso sim é uma lacuna concreta.
   Exemplo que NÃO é problema: uma regra que explica claramente o quê e
   quando, mas não menciona prazo, porque aquele tipo de regra
   simplesmente não tem prazo (não invente uma exigência que não existe
   para aquele caso).

Priorize sua atenção nos problemas que MAIS afetam a qualidade de um
agente de IA: tabelas mal formatadas, siglas/códigos sem explicação,
conteúdo contraditório, informação decisiva realmente ausente (a IA não
teria como saber a resposta), e regras de negócio que omitem uma condição
crítica (não apenas por serem sucintas). Um documento pode ser sucinto e
direto e ainda assim ser perfeitamente adequado para um agente de IA — a
meta não é "reescrever tudo de forma mais elaborada", é garantir que a
informação necessária esteja presente e sem ambiguidade real.

IMPORTANTE sobre o campo "sugestao_correcao": sempre que possível, não
escreva só uma instrução genérica — mostre a correção JÁ APLICADA a UM
trecho real encontrado no documento, como um único "antes/depois"
concreto que a pessoa pode usar como modelo. Mostre APENAS 1 exemplo
aplicado, mesmo que o mesmo problema se repita em várias partes do
documento ou envolva várias categorias/casos — a pessoa consegue replicar
o padrão sozinha a partir de 1 exemplo claro; não é necessário (nem
desejável) reescrever cada ocorrência uma por uma. Isso vale especialmente
para:
- "cores_convertidas_em_texto": não diga apenas "adicione uma tag de
  texto" — reescreva UM trecho real substituindo a dependência de cor por
  uma tag textual, seguindo o padrão do Guia (ex: "[Categoria: X]" antes
  do trecho). Se o documento tem "críticas em amarelo = equipe X, em
  vermelho = equipe Y", mostre a transformação aplicada a apenas UM desses
  casos como exemplo — não é preciso reescrever para X, Y e Z.
- "siglas_explicadas": mostre UMA frase com a sigla já seguida do
  significado entre parênteses, usando a sigla real encontrada.
- Qualquer critério onde reescrever o trecho específico é possível: prefira
  sempre mostrar o resultado já pronto (1 exemplo só) em vez de somente
  descrever o que fazer.

IMPORTANTE sobre os critérios de rastreabilidade/metadados (documento_vigente_explicito,
responsavel_definido, data_atualizacao, versao_identificada, status_documento,
historico_alteracoes, faq, exemplos_aplicacao, metadados_enriquecidos,
revisao_periodica_definida): a maioria dos documentos operacionais reais NÃO
tem esses campos declarados formalmente, e isso é normal, não é um defeito
grave. Para esses critérios especificamente:
- Aceite evidência indireta ou inferida (nome de área citado no texto,
  cabeçalho, rodapé, nome do arquivo, menção a "atualizado recentemente" etc.)
  como suficiente para "atende" ou "parcial".
- Só marque "nao_atende" se a ausência total do elemento realmente
  compromete o entendimento do conteúdo, não apenas porque o campo
  formal não existe.
- Seja generoso: esses critérios existem para RECOMENDAR boas práticas de
  governança, não para reprovar documentos operacionais legítimos.

Já os critérios de conteúdo e formatação (escrita_autocontida, paragrafos_coesos,
titulos_descritivos, conteudo_contraditorio_tratado, textos_revogados_tratados,
cores_convertidas_em_texto, imagens_tratadas_corretamente, tabelas_revisadas,
siglas_explicadas, densidade_informativa, linguagem_informal_evitada,
regras_de_negocio_completas, notas_incorporadas_ao_texto, termos_padronizados)
devem ser avaliados com atenção real ao impacto (conforme a orientação de
rigor e visão holística acima), pois afetam a qualidade de recuperação e
leitura por IA quando o problema é genuíno — mas sem exagerar em nuances
de estilo que não mudam o entendimento.

Para cada critério, primeiro analise o texto e só depois decida o status.
Se o status for "nao_atende" ou "parcial", explique seu raciocínio no
campo "justificativa" antes de escrever os demais campos — nunca decida
o texto da justificativa antes de ter clareza do problema real.

IMPORTANTE sobre repetição de páginas/locais: se o MESMO problema (mesma
causa, mesmo diagnóstico) ocorre em várias páginas, seções, abas ou slides,
consolide em UM item só, listando TODAS as localizações no campo
"localizacao" (ex: ["Página 2", "Página 5", "Página 9"]).
Porém, se o mesmo critério tem problemas DIFERENTES em partes diferentes do
documento (diagnóstico ou recomendação diferentes — ex: uma tabela mal
formatada na página 2 e uma sigla sem explicação também relacionada a
tabelas na página 7), reporte como itens SEPARADOS, mesmo compartilhando o
mesmo "id" de critério. Nunca junte diagnósticos diferentes num único item
só para evitar repetir o id.

IMPORTANTE sobre critérios sem ocorrência no documento: se um critério
simplesmente não tem nenhum elemento aplicável no documento (ex: o critério
é sobre tabelas e o documento não tem nenhuma tabela), marque status
"nao_aplicavel". Isso não será exibido como problema.

IMPORTANTE sobre o campo "risco_ia": para cada problema encontrado (status
"nao_atende" ou "parcial"), explique em 1 frase curta, em linguagem simples
e não técnica (para alguém da área de negócio, sem jargão de TI/RAG), o que
ESSE problema específico pode causar na resposta do agente de IA. Não seja
genérico sobre o critério — seja específico sobre o caso encontrado. Ex:
"A IA pode misturar o prazo antigo com o novo e responder com o valor
errado." em vez de "Isso pode prejudicar a recuperação de informação."

IMPORTANTE sobre economia de tokens na resposta: o grupo em que cada
critério entra é definido ESTRITAMENTE pelo status, sem excecão:
- status "atende" ou "nao_aplicavel" → grupo "criterios_sem_problema"
  (só id e status, sem justificativa nem nenhum outro campo).
- status "nao_atende" ou "parcial" → grupo "criterios_com_problema" (com
  justificativa, risco_ia, trechos, localizacao e sugestão completos).
NUNCA coloque um critério com status "parcial" ou "nao_atende" no grupo
"criterios_sem_problema", mesmo que o problema pareça leve, pequeno ou
seja sobre um critério de rastreabilidade/metadado — "parcial" SEMPRE
exige a justificativa completa no grupo "criterios_com_problema", sem
exceção. Se está em dúvida se o problema é leve o bastante para não
justificar, marque "atende" (não "parcial") em vez de deixar um "parcial"
sem justificativa.
Isso significa que TODO critério da lista deve aparecer em exatamente 1
dos dois grupos — nunca repita o mesmo id nos dois grupos, e nunca omita
um critério por completo.

Responda SOMENTE em JSON válido, seguindo exatamente o schema abaixo,
sem nenhum texto antes ou depois do JSON:

{
  "criterios_sem_problema": [
    {"id": "<id do critério>", "status": "atende" | "nao_aplicavel"}
  ],
  "criterios_com_problema": [
    {
      "id": "<id do critério>",
      "status": "nao_atende" | "parcial",
      "justificativa": "<explicação breve, 1-2 frases>",
      "risco_ia": "<1 frase simples: o que esse problema específico pode causar na resposta do agente de IA>",
      "trechos_problematicos": ["<trecho curto do documento, se houver problema>"],
      "localizacao": ["<ex: Página 4, Aba: Prazos>"],
      "sugestao_correcao": "<correção já aplicada ao trecho real, no formato antes/depois quando possível — não apenas uma instrução genérica>"
    }
  ]
}
"""


def _montar_prompt_usuario(texto_documento: str, formato_arquivo: str) -> str:
    """Monta o prompt enviado ao modelo: lista os critérios da rubrica,
    os exemplos de referência e o texto do documento a avaliar."""
    criterios_texto = "\n".join(
        f"- {c['id']}: {c['descricao']}"
        for c in CRITERIOS_OBRIGATORIOS + CRITERIOS_DESEJAVEIS
    )
    return f"""FORMATO DO ARQUIVO SENDO AVALIADO: {formato_arquivo}
(Use essa informação para aplicar corretamente a regra de cabeçalho de
tabela do critério "tabelas_revisadas" — a regra de underscore/sem acento
só vale para Excel, não para os demais formatos.)

CRITÉRIOS A AVALIAR (id: descrição):
{criterios_texto}

EXEMPLOS DE REFERÊNCIA (padrão já definido pela Governança de IA):
{_FEW_SHOT_EXAMPLES}

DOCUMENTO A AVALIAR:
{texto_documento}

Avalie CADA critério da lista acima, um por um, e responda no formato JSON
especificado, separando em "criterios_sem_problema" (atende/nao_aplicavel,
sem justificativa) e "criterios_com_problema" (nao_atende/parcial, com
justificativa completa).
"""


# ============================================================
# CHAMADA AO MODELO (workflow "Gerador de Texto" da Gateway)
# ============================================================

_STATUS_SEM_PROBLEMA = {"atende", "nao_aplicavel"}
_STATUS_COM_PROBLEMA = {"nao_atende", "parcial"}


def _normalizar_resposta_llm(parsed: dict) -> dict:
    """Expande o formato compacto que o modelo devolve (separado em
    "criterios_sem_problema", só id+status, e "criterios_com_problema",
    com justificativa completa) de volta para o formato interno único
    {"criterios": [...]} que o resto do pipeline (score, redução de
    lotes, relatório HTML) já espera. Isso existe só para economizar
    tokens de SAÍDA na resposta do modelo — que custam bem mais caro que
    tokens de entrada — sem exigir mudança em nenhuma outra função.

    Nem sempre o modelo classifica um critério no grupo certo (ex:
    manda um status "parcial" dentro de "criterios_sem_problema", que só
    tem id+status) — isso gerava linhas na tela com status "Parcial" mas
    todo o resto em branco (diagnóstico, risco, exemplo, recomendação),
    porque o item vinha sem justificativa nenhuma. Por isso, em vez de
    confiar na posição em que o item veio, cada item é reclassificado
    aqui pelo STATUS real: se o status indica problema (nao_atende ou
    parcial) mas o item veio sem justificativa (porque caiu no grupo
    errado), ele ainda é tratado como "com problema" — só que sem os
    detalhes que o modelo não forneceu, em vez de aparecer como se
    tivesse justificativa vazia por design.

    Se a resposta já vier no formato antigo (com "criterios" direto —
    ex: em testes ou numa resposta inesperada do modelo), devolve como
    está, sem quebrar."""
    if "criterios" in parsed:
        return parsed

    criterios = []
    for item in parsed.get("criterios_sem_problema", []):
        status = item.get("status", "atende")
        if status in _STATUS_COM_PROBLEMA:
            # Modelo classificou errado: status indica problema, mas o
            # item veio no grupo "sem_problema" (sem justificativa). Não
            # descartamos o status real nem inventamos justificativa —
            # só sinalizamos a lacuna de forma visível, em vez de deixar
            # tudo em branco silenciosamente.
            print(
                f"[curadoria] ⚠️ Critério '{item.get('id')}' veio com status "
                f"'{status}' no grupo sem_problema (sem justificativa do modelo)."
            )
            criterios.append({
                "id": item["id"],
                "status": status,
                "justificativa": "",
                "risco_ia": "",
                "trechos_problematicos": [],
                "localizacao": [],
                "sugestao_correcao": "",
                # Sinaliza que este item não tem detalhe algum (o modelo
                # classificou errado e não forneceu justificativa) — usado
                # para NÃO exibir a linha na tabela/relatório (não há nada
                # acionável para mostrar), mas o status real é preservado
                # e continua contando normalmente no cálculo do score.
                "diagnostico_incompleto": True,
            })
        else:
            criterios.append({
                "id": item["id"],
                "status": status,
                "justificativa": "",
                "risco_ia": "",
                "trechos_problematicos": [],
                "localizacao": [],
                "sugestao_correcao": "",
            })
    for item in parsed.get("criterios_com_problema", []):
        criterios.append({
            "id": item["id"],
            "status": item.get("status", "parcial"),
            "justificativa": item.get("justificativa", ""),
            "risco_ia": item.get("risco_ia", ""),
            "trechos_problematicos": item.get("trechos_problematicos", []),
            "localizacao": item.get("localizacao", []),
            "sugestao_correcao": item.get("sugestao_correcao", ""),
        })
    return {"criterios": criterios}


def _chamar_llm(prompt_usuario: str, client: GatewayClient) -> dict:
    """Faz 1 chamada ao modelo com o prompt já pronto e devolve o JSON
    de avaliação já decodificado e normalizado. client é injetado pelo
    controller (Depends(GatewayClient)), o que permite substituí-lo por
    um fake nos testes."""
    try:
        conteudo = client.gerar_texto(
            _SYSTEM_PROMPT, prompt_usuario, CURATION_MODEL,
            temperature=0,
            max_tokens=MAX_TOKENS_AVALIACAO,
            response_format={"type": "json_object"},
        )
    except Exception:
        print("[curadoria] ❌ ERRO na chamada ao modelo via Gateway:")
        print(traceback.format_exc())
        raise

    try:
        parsed = json.loads(conteudo)
    except json.JSONDecodeError:
        print("[curadoria] ❌ ERRO ao interpretar JSON da resposta. Conteúdo bruto recebido:")
        print(conteudo[:3000])
        raise

    return _normalizar_resposta_llm(parsed)


# ============================================================
# RESUMO NARRATIVO
# ============================================================

_RESUMO_SYSTEM_PROMPT = """Você é um especialista em curadoria de documentos para
RAG. Você recebe um diagnóstico JÁ PRONTO (lista de critérios avaliados, com
status e justificativa) sobre um documento, e escreve um parecer em texto
corrido — um único parágrafo, sem bullets, sem markdown, sem repetir a lista
item por item.

O parágrafo deve conter, nesta ordem:
1. Um veredito de abertura, situando a classificação geral e o motivo central.
2. O que está bom no documento — os critérios que atendem, resumidos de forma
   agregada (não liste um por um, agrupe por tema).
3. O que precisa de atenção — os critérios com problema, agrupados por tema,
   mencionando as localizações mais relevantes de forma natural na frase.
4. Uma frase final objetiva sobre o que fazer a seguir.

Você NÃO tem acesso ao texto original do documento — trabalhe apenas com o
diagnóstico estruturado fornecido. Não invente detalhes que não estão no
diagnóstico. Responda em português, em um único parágrafo corrido, sem título.
"""


def gerar_resumo_narrativo(
    avaliacao: dict,
    score_info: dict,
    client: GatewayClient,
    executor: concurrent.futures.ThreadPoolExecutor | None = None,
) -> str:
    """Gera o parecer em texto corrido a partir do JSON já avaliado.
    Não reprocessa o documento original, só resume o diagnóstico.

    Se "executor" for passado, a chamada à Gateway é submetida a esse
    pool compartilhado em vez de ser feita direto — importante quando
    vários arquivos de um mesmo upload terminam a avaliação quase juntos
    (ver processar_lote_arquivos): sem isso, cada arquivo dispararia seu
    resumo simultaneamente, somando concorrência além do teto de
    MAX_CHAMADAS_PARALELAS que o pool de lotes já respeita."""
    diagnostico_compacto = {
        "classificacao": score_info["classificacao"],
        "score": score_info["score"],
        "criterios": [
            {"id": c["id"], "status": c["status"], "justificativa": c.get("justificativa", "")}
            for c in avaliacao.get("criterios", [])
            if c["id"] in _IDS_OBRIGATORIOS  # resumo cobre conteúdo/formatação, não metadados
        ],
    }
    entrada_usuario = json.dumps(diagnostico_compacto, ensure_ascii=False)

    try:
        chamada = lambda: client.gerar_texto(
            _RESUMO_SYSTEM_PROMPT, entrada_usuario, CURATION_MODEL,
            temperature=0.3,
            max_tokens=MAX_TOKENS_RESUMO,
        )
        texto = executor.submit(chamada).result() if executor is not None else chamada()
        return texto.strip()
    except Exception:
        print("[curadoria] ⚠️ Falha ao gerar resumo narrativo (não bloqueia o restante):")
        print(traceback.format_exc())
        return ""


def _dividir_em_lotes(blocos: list[Bloco], limite_chars: int) -> list[list[Bloco]]:
    """Junta blocos num lote até quase estourar limite_chars, aí começa
    um lote novo. Nunca quebra um bloco no meio: um bloco inteiro pode
    passar do limite sozinho se ele já for maior que isso."""
    lotes, atual, tamanho = [], [], 0
    for bloco in blocos:
        tamanho_bloco = len(bloco.texto)
        if atual and tamanho + tamanho_bloco > limite_chars:
            lotes.append(atual)
            atual, tamanho = [], 0
        atual.append(bloco)
        tamanho += tamanho_bloco
    if atual:
        lotes.append(atual)
    return lotes


def _remover_duplicatas(itens: list[str]) -> list[str]:
    """Remove duplicatas de uma lista mantendo a ordem original."""
    vistos = set()
    resultado = []
    for item in itens:
        if item not in vistos:
            vistos.add(item)
            resultado.append(item)
    return resultado


def _reduzir_avaliacoes(avaliacoes_por_lote: list[dict]) -> dict:
    """Combina os critérios de todos os lotes. Só mescla itens que são
    duplicata exata (mesmo id e mesma justificativa)."""
    todos, vistos = [], {}
    for avaliacao in avaliacoes_por_lote:
        for item in avaliacao.get("criterios", []):
            chave = (item["id"], (item.get("justificativa") or "").strip().lower())
            if chave in vistos:
                idx = vistos[chave]
                todos[idx]["trechos_problematicos"] = _remover_duplicatas(
                    list(todos[idx].get("trechos_problematicos", []))
                    + list(item.get("trechos_problematicos", []))
                )
                todos[idx]["localizacao"] = _remover_duplicatas(
                    list(todos[idx].get("localizacao", []))
                    + list(item.get("localizacao", []))
                )
            else:
                novo_item = dict(item)
                novo_item["localizacao"] = _remover_duplicatas(
                    list(novo_item.get("localizacao", []))
                )
                novo_item["trechos_problematicos"] = _remover_duplicatas(
                    list(novo_item.get("trechos_problematicos", []))
                )
                vistos[chave] = len(todos)
                todos.append(novo_item)
    return {"criterios": todos}


_ROTULO_FORMATO = {
    ".pdf": "PDF",
    ".docx": "Word (.docx)",
    ".xlsx": "Excel (.xlsx)",
    ".xlsm": "Excel (.xlsm)",
    ".pptx": "PowerPoint (.pptx)",
}


def _chamar_llm_com_retry(prompt: str, client: GatewayClient) -> dict:
    """Repete a chamada ao modelo até TENTATIVAS_POR_LOTE vezes antes de
    desistir do lote. Cobre falhas transitórias (timeout de polling da
    Gateway, instabilidade de rede) que não se repetem numa nova
    tentativa — evita que 1 lote lento derrube o arquivo inteiro."""
    ultimo_erro = None
    for tentativa in range(1, TENTATIVAS_POR_LOTE + 1):
        try:
            return _chamar_llm(prompt, client)
        except Exception as e:
            ultimo_erro = e
            if tentativa < TENTATIVAS_POR_LOTE:
                print(
                    f"[curadoria] ⚠️ Tentativa {tentativa}/{TENTATIVAS_POR_LOTE} "
                    f"falhou para este lote ({e}). Tentando novamente..."
                )
                time.sleep(ESPERA_ENTRE_TENTATIVAS_SEGUNDOS)
    raise ultimo_erro


def avaliar_documento(
    blocos: list[Bloco],
    formato_arquivo: str,
    client: GatewayClient,
    executor: concurrent.futures.ThreadPoolExecutor | None = None,
) -> dict:
    """Map-reduce de avaliação. Se o documento for grande, divide em lotes
    (por página, seção, slide ou aba) e avalia até MAX_CHAMADAS_PARALELAS
    lotes ao mesmo tempo (com retry por lote), agregando o resultado no
    final — sem cortar um bloco no meio.

    O parâmetro "executor" permite compartilhar 1 único pool de threads
    entre VÁRIOS arquivos de um mesmo upload (ver processar_arquivo e o
    controller de curadoria): sem isso, cada arquivo abriria seu próprio
    pool de até 10 threads, e um upload de 5 arquivos poderia gerar até
    50 chamadas simultâneas à Gateway — 5x acima do limite de concorrência
    que a Gateway suporta. Compartilhando o mesmo pool entre todos os
    arquivos do lote, o teto de 10 chamadas simultâneas vale para o
    upload inteiro, não por arquivo. Se "executor" não for passado (uso
    isolado, como nos testes), a função cria e fecha seu próprio pool
    local, mantendo o comportamento de antes.

    Processar em paralelo em vez de sequencial não reduz o custo (mesmo
    total de tokens), mas reduz o tempo de espera do usuário na proporção
    do paralelismo: 60 lotes a ~90s cada, um atrás do outro, levam ~90min;
    em paralelo (10 por vez), levam ~9min — essencial para caber num
    orçamento de resposta de ~10 minutos por documento.

    Antes de chamar a Gateway, confere quantos lotes o documento geraria:
    - acima de LIMITE_LOTES_BLOQUEIO: recusa o arquivo inteiro (nenhuma
      chamada ao modelo é feita, custo de IA = zero).
    - acima de LIMITE_LOTES_PROCESSAVEIS (mas dentro do bloqueio): avalia
      só os primeiros LIMITE_LOTES_PROCESSAVEIS lotes e marca o restante
      como "não avaliado por limite de tamanho", sem gastar tokens com o
      excedente.

    Um documento grande pode gerar dezenas de lotes/chamadas à Gateway —
    se cada uma tiver, digamos, 1% de chance de falhar, a chance de pelo
    menos 1 falhar num lote de 30 chamadas já passa de 25%. Por isso um
    lote que falhar mesmo após todas as tentativas de retry NÃO derruba
    os demais: ele é marcado como "não avaliado" e o restante do
    documento segue sendo avaliado normalmente. Só levanta exceção (e
    falha o arquivo inteiro) se NENHUM lote conseguir ser avaliado."""
    lotes = _dividir_em_lotes(blocos, MAX_CHARS_PER_LLM_CALL)

    if len(lotes) > LIMITE_LOTES_BLOQUEIO:
        raise DocumentoMuitoGrandeError(
            f"Documento grande demais para avaliação: geraria {len(lotes)} lotes de "
            f"análise, acima do limite de {LIMITE_LOTES_BLOQUEIO} suportado. "
            "Divida o arquivo em partes menores antes de enviar (ex: por seção, "
            "capítulo ou aba) e envie cada parte separadamente."
        )

    lotes_truncados_localizacoes = []
    if len(lotes) > LIMITE_LOTES_PROCESSAVEIS:
        lotes_truncados = lotes[LIMITE_LOTES_PROCESSAVEIS:]
        lotes = lotes[:LIMITE_LOTES_PROCESSAVEIS]
        lotes_truncados_localizacoes = [[b.localizacao for b in lote] for lote in lotes_truncados]
        print(
            f"[curadoria] ⚠️ Documento truncado: {len(lotes_truncados)} lote(s) "
            f"além do limite de {LIMITE_LOTES_PROCESSAVEIS} não serão avaliados."
        )

    # Cada tarefa processa 1 lote inteiro (monta o prompt e chama
    # _chamar_llm_com_retry, que já tem seu próprio retry sequencial
    # interno). Submetidas ao "executor" compartilhado (ou a um local, se
    # nenhum foi passado) — o teto de MAX_CHAMADAS_PARALELAS (limitado a
    # _LIMITE_CONCORRENCIA_GATEWAY) nunca é excedido, mesmo somando lotes
    # de vários arquivos processados ao mesmo tempo.
    def _avaliar_lote(lote: list[Bloco]) -> tuple[list[str], dict | None, Exception | None]:
        localizacoes = [b.localizacao for b in lote]
        prompt = _montar_prompt_usuario(_concatenar_blocos(lote), formato_arquivo)
        try:
            return localizacoes, _chamar_llm_com_retry(prompt, client), None
        except Exception as e:
            return localizacoes, None, e

    def _executar(pool: concurrent.futures.ThreadPoolExecutor):
        avaliacoes, lotes_com_falha = [], []
        for localizacoes, avaliacao, erro in pool.map(_avaliar_lote, lotes):
            if erro is not None:
                print(f"[curadoria] ❌ Lote não avaliado após {TENTATIVAS_POR_LOTE} tentativas ({localizacoes}): {erro}")
                lotes_com_falha.append(localizacoes)
            else:
                avaliacoes.append(avaliacao)
        return avaliacoes, lotes_com_falha

    if executor is not None:
        avaliacoes, lotes_com_falha = _executar(executor)
    else:
        with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_CHAMADAS_PARALELAS) as pool_local:
            avaliacoes, lotes_com_falha = _executar(pool_local)

    if not avaliacoes and not lotes_truncados_localizacoes:
        raise RuntimeError(
            f"Nenhum dos {len(lotes)} lote(s) do documento pôde ser avaliado "
            f"(todas as tentativas falharam em todos os lotes)."
        )

    resultado = _reduzir_avaliacoes(avaliacoes)
    if lotes_com_falha:
        resultado["lotes_nao_avaliados"] = lotes_com_falha
    if lotes_truncados_localizacoes:
        resultado["lotes_truncados"] = lotes_truncados_localizacoes
    return resultado


_IDS_OBRIGATORIOS = {c["id"] for c in CRITERIOS_OBRIGATORIOS}
_IDS_DESEJAVEIS = {c["id"] for c in CRITERIOS_DESEJAVEIS}
IDS_RASTREABILIDADE = _IDS_DESEJAVEIS

_PRIORIDADE_STATUS = {"nao_atende": 0, "parcial": 1, "atende": 2, "nao_aplicavel": 3}


def _pior_status_por_criterio(criterios: list[dict]) -> dict:
    """Colapsa a lista de critérios (que pode ter o mesmo id repetido com
    diagnósticos diferentes) para um único status por id: o pior
    encontrado. Usada só para pontuação, nunca para exibição."""
    pior = {}
    for c in criterios:
        atual = pior.get(c["id"])
        if atual is None or _PRIORIDADE_STATUS.get(c["status"], 9) < _PRIORIDADE_STATUS.get(atual, 9):
            pior[c["id"]] = c["status"]
    return pior


def _taxa_atendimento(status_por_criterio: dict, ids: set[str]) -> float:
    """Percentual de atendimento dos critérios do grupo, de 0 a 1: atende
    vale 1, parcial vale 0.5, nao_atende vale 0. Ignora nao_aplicavel."""
    relevantes = [s for cid, s in status_por_criterio.items() if cid in ids and s != "nao_aplicavel"]
    if not relevantes:
        return 1.0
    pontos = {"atende": 1.0, "parcial": 0.5, "nao_atende": 0.0}
    return sum(pontos.get(s, 0.0) for s in relevantes) / len(relevantes)


def calcular_score(avaliacao: dict) -> dict:
    criterios = avaliacao.get("criterios", [])
    status_por_criterio = _pior_status_por_criterio(criterios)

    taxa_obrig = _taxa_atendimento(status_por_criterio, _IDS_OBRIGATORIOS)
    taxa_desej = _taxa_atendimento(status_por_criterio, _IDS_DESEJAVEIS)

    # Score geral = 100% Conteúdo e Formatação. Rastreabilidade não pesa aqui.
    score = round(taxa_obrig * 100, 1)

    problemas_criticos = [
        cid for cid, s in status_por_criterio.items() if cid in _IDS_OBRIGATORIOS and s == "nao_atende"
    ]
    problemas_recomendados = [
        cid for cid, s in status_por_criterio.items()
        if (s in ("nao_atende", "parcial") and cid not in _IDS_OBRIGATORIOS)
        or (cid in _IDS_OBRIGATORIOS and s == "parcial")
    ]

    if taxa_obrig == 1.0:
        classificacao = "Pronto para uso em IA"
    elif len(problemas_criticos) == 0:
        classificacao = "Pronto com ressalvas"
    elif len(problemas_criticos) <= 3:
        classificacao = "Necessita ajuste"
    else:
        classificacao = "Não recomendado para uso em IA"

    qtd_linhas_problema = len([
        c for c in criterios
        if c["id"] in _IDS_OBRIGATORIOS and c["status"] not in ("atende", "nao_aplicavel")
    ])

    return {
        "score": score,
        "score_conteudo_formato": round(taxa_obrig * 100, 1),
        "score_rastreabilidade": round(taxa_desej * 100, 1),
        "classificacao": classificacao,
        "qtd_problemas_criticos": len(problemas_criticos),
        "qtd_problemas_recomendados": len(problemas_recomendados),
        "qtd_linhas_problema": qtd_linhas_problema,
    }


def processar_lote_arquivos(
    arquivos: list[tuple[str, bytes]], client: GatewayClient
) -> list[dict]:
    """Processa vários arquivos de um mesmo upload, compartilhando 1 único
    pool de até MAX_CHAMADAS_PARALELAS chamadas simultâneas à Gateway entre
    TODOS os arquivos — não 1 pool por arquivo.

    Sem isso, processar N arquivos de um upload teria 2 problemas:
    1. Sequencial: cada arquivo (com seu próprio paralelismo interno de
       lotes) só começa depois que o anterior termina por completo — um
       upload de 5 arquivos de ~9min cada levaria ~45min, muito acima do
       orçamento de ~10min pensado por arquivo.
    2. Se em vez disso cada arquivo abrisse seu próprio pool de até 10
       chamadas paralelas, rodando os 5 arquivos ao mesmo tempo somaria
       até 50 chamadas simultâneas à Gateway — 5x acima do limite de
       concorrência que a Gateway suporta.

    A solução usa 2 pools com papéis diferentes:
    - "pool_gateway": único pool com no máximo MAX_CHAMADAS_PARALELAS
      threads, usado por TODOS os lotes de TODOS os arquivos (e pelos
      resumos narrativos) — é o único lugar que efetivamente conversa com
      a Gateway, então o teto de concorrência vale para o upload inteiro.
    - "pool_arquivos": pool de orquestração, 1 thread por arquivo. Essas
      threads não chamam a Gateway diretamente — elas só chamam
      processar_arquivo(), que por sua vez submete o trabalho real ao
      pool_gateway e espera o resultado. Por serem pools diferentes, não
      há risco de deadlock (threads de orquestração bloqueadas não
      competem pelas mesmas vagas que os lotes precisam)."""
    if not arquivos:
        return []

    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_CHAMADAS_PARALELAS) as pool_gateway:
        with concurrent.futures.ThreadPoolExecutor(max_workers=len(arquivos)) as pool_arquivos:
            futuros = [
                pool_arquivos.submit(processar_arquivo, nome, conteudo, client, executor=pool_gateway)
                for nome, conteudo in arquivos
            ]
            return [f.result() for f in futuros]


def agregar_lote(resultados_por_arquivo: list[dict]) -> dict:
    """Tabela consolidada de múltiplos arquivos. Só código, sem IA."""
    tabela = [
        {
            "arquivo": r["arquivo"],
            "classificacao": r["score_info"]["classificacao"],
            "score": r["score_info"]["score"],
            "problemas_criticos": r["score_info"]["qtd_problemas_criticos"],
            "problemas_recomendados": r["score_info"]["qtd_problemas_recomendados"],
        }
        for r in resultados_por_arquivo
    ]
    tabela.sort(key=lambda x: x["score"])
    return {
        "total_arquivos": len(tabela),
        "media_score": round(sum(t["score"] for t in tabela) / len(tabela), 1) if tabela else 0,
        "arquivos_bloqueantes": sum(1 for t in tabela if t["problemas_criticos"] > 0),
        "tabela": tabela,
    }


# ============================================================
# PIPELINE
# ============================================================

def processar_arquivo(
    nome_arquivo: str,
    conteudo_bytes: bytes,
    client: GatewayClient,
    executor: concurrent.futures.ThreadPoolExecutor | None = None,
) -> dict:
    """Recebe o nome e os bytes do arquivo (já lidos pelo controller) e
    roda a pipeline inteira: primeiro a extração local, depois a
    avaliação pelo LLM, e por fim o cálculo do score.

    O parâmetro "executor" é repassado para avaliar_documento() — permite
    que vários arquivos de um mesmo upload (ver processar_lote_arquivos)
    compartilhem o mesmo teto de MAX_CHAMADAS_PARALELAS chamadas
    simultâneas à Gateway, em vez de cada arquivo abrir seu próprio pool
    de até 10 chamadas paralelas."""
    import tempfile

    with tempfile.TemporaryDirectory() as tmpdir:
        caminho_temp = Path(tmpdir) / nome_arquivo
        caminho_temp.write_bytes(conteudo_bytes)

        try:
            blocos = extrair_texto(str(caminho_temp))
        except ExtracaoVaziaError as e:
            return {"arquivo": nome_arquivo, "erro": str(e), "requer_ocr": True}
        except Exception as e:
            print(f"[curadoria] ❌ ERRO na extração de '{nome_arquivo}':")
            print(traceback.format_exc())
            return {"arquivo": nome_arquivo, "erro": f"Falha na extração: {e}"}

        ext = Path(nome_arquivo).suffix.lower()

        metricas_documento = calcular_metricas_documento(blocos)
        porte = calcular_porte_documento(blocos, ext)
        metricas_documento.update(porte)

        formato_arquivo = _ROTULO_FORMATO.get(ext, ext)

        try:
            avaliacao = avaliar_documento(blocos, formato_arquivo, client, executor=executor)
        except DocumentoMuitoGrandeError as e:
            # Recusado ANTES de qualquer chamada à Gateway — custo de IA
            # zero. Erro diferente do genérico para o frontend poder
            # exibir de forma distinta (não é uma falha técnica, é um
            # limite de porte do documento).
            return {"arquivo": nome_arquivo, "erro": str(e), "documento_grande_demais": True}
        except Exception as e:
            print(f"[curadoria] ❌ ERRO na avaliação de '{nome_arquivo}':")
            print(traceback.format_exc())
            return {"arquivo": nome_arquivo, "erro": f"Falha na avaliação: {e}"}

    score_info = calcular_score(avaliacao)
    resumo_narrativo = gerar_resumo_narrativo(avaliacao, score_info, client, executor=executor)
    lotes_nao_avaliados = avaliacao.pop("lotes_nao_avaliados", None) or []
    lotes_truncados = avaliacao.pop("lotes_truncados", None) or []
    resultado = {
        "arquivo": nome_arquivo,
        "avaliacao_detalhada": avaliacao,
        "score_info": score_info,
        "metricas_documento": metricas_documento,
        "resumo_narrativo": resumo_narrativo,
    }
    if lotes_nao_avaliados or lotes_truncados:
        # Documento onde 1+ trechos não entraram na avaliação, seja por
        # falha de comunicação (mesmo após retry) ou por corte de limite
        # de tamanho (LIMITE_LOTES_PROCESSAVEIS). O arquivo NÃO é
        # descartado — os demais trechos foram avaliados normalmente —
        # mas quem está usando a tela precisa saber que o diagnóstico é
        # parcial e por qual motivo.
        partes_aviso = []
        if lotes_nao_avaliados:
            partes_aviso.append(
                f"{len(lotes_nao_avaliados)} trecho(s) não puderam ser avaliados "
                "(falha de comunicação com o modelo mesmo após novas tentativas)"
            )
        if lotes_truncados:
            partes_aviso.append(
                f"{len(lotes_truncados)} trecho(s) não foram avaliados por exceder o limite "
                f"de {LIMITE_LOTES_PROCESSAVEIS} lotes suportado por avaliação "
                "(documento muito extenso; considere dividir em partes menores)"
            )
        resultado["aviso_avaliacao_parcial"] = (
            ", e ".join(partes_aviso) + ". O restante do documento foi avaliado normalmente "
            "— o score e o diagnóstico abaixo cobrem apenas a parte avaliada com sucesso."
        )
        resultado["locais_nao_avaliados"] = [
            loc for grupo in (lotes_nao_avaliados + lotes_truncados) for loc in grupo
        ]
    return resultado


# ============================================================
# RELATÓRIO HTML
# ============================================================

STATUS_LABEL = {
    "atende": "✅ Atende",
    "parcial": "🟡 Parcial",
    "nao_atende": "🔴 Não atende",
    "nao_aplicavel": "⚪ Não aplicável",
}


def gerar_relatorio_html(resultado: dict) -> str:
    """Monta um HTML autocontido (CSS embutido, sem arquivo externo) com
    o diagnóstico de 1 documento — mesma informação exibida na tela
    Streamlit, formatada para impressão ou compartilhamento."""
    arquivo = html.escape(resultado["arquivo"])
    score_info = resultado["score_info"]
    metricas = resultado.get("metricas_documento") or {}
    resumo = html.escape(resultado.get("resumo_narrativo") or "Resumo não disponível.")
    gerado_em = datetime.now().strftime("%d/%m/%Y às %H:%M")

    criterios_com_problema = [
        c for c in resultado["avaliacao_detalhada"]["criterios"]
        if c["status"] not in ("atende", "nao_aplicavel")
        and c["id"] not in IDS_RASTREABILIDADE
        and not c.get("diagnostico_incompleto")
    ]

    if criterios_com_problema:
        linhas_tabela = ""
        for c in criterios_com_problema:
            justificativa = html.escape((c.get("justificativa") or "").strip())
            sugestao = html.escape((c.get("sugestao_correcao") or "").strip())
            mostrar_sugestao = bool(sugestao) and sugestao.lower() != justificativa.lower()
            locais = list(dict.fromkeys(c.get("localizacao", [])))
            local = html.escape(", ".join(locais) or "—")
            status = html.escape(STATUS_LABEL.get(c["status"], c["status"]))
            criterio = html.escape(c["id"].replace("_", " ").capitalize())
            risco = html.escape((c.get("risco_ia") or "").strip()) or "—"
            trechos = [t.strip() for t in c.get("trechos_problematicos", []) if t and t.strip()]
            exemplo = f'Trecho: "{html.escape(trechos[0])}"' if trechos else "—"
            linhas_tabela += f"""
<tr>
<td>{criterio}</td>
<td>{justificativa}</td>
<td>{risco}</td>
<td><i>{exemplo}</i></td>
<td>{sugestao if mostrar_sugestao else "—"}</td>
<td>{local}</td>
<td>{status}</td>
</tr>"""
        tabela_html = f"""
<table>
<thead>
<tr><th>Critério</th><th>Diagnóstico</th><th>Risco</th><th>Exemplo</th><th>Recomendação</th><th>Local</th><th>Status</th></tr>
</thead>
<tbody>{linhas_tabela}
</tbody>
</table>"""
    else:
        tabela_html = "<p>Nenhum problema encontrado — documento atende a todos os critérios de conteúdo e formatação.</p>"

    metricas_texto = (
        f"{metricas.get('num_blocos', 0)} blocos extraídos · "
        f"{metricas.get('total_caracteres', 0):,} caracteres · "
        f"{metricas.get('proporcao_tabular', 0)*100:.0f}% tabular · "
        f"{metricas.get('blocos_duplicados', 0)} bloco(s) duplicado(s)"
    ).replace(",", ".")

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>DSH - Relatório de Diagnóstico</title>
<style>
    body {{ font-family: 'Segoe UI', Arial, sans-serif; color: #2D3748; max-width: 860px; margin: 40px auto; padding: 0 24px; line-height: 1.6; }}
    .cabecalho {{ display: flex; align-items: center; gap: 14px; border-bottom: 3px solid #1E3A8A; padding-bottom: 16px; margin-bottom: 24px; }}
    .marca {{ width: 48px; height: 48px; border-radius: 12px; background: linear-gradient(135deg, #1E3A8A, #1E40AF); color: white; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 20px; flex-shrink: 0; overflow: hidden; }}
    .cabecalho h1 {{ font-size: 19px; margin: 0; }}
    .cabecalho .sub {{ font-size: 11px; color: #1E3A8A; text-transform: uppercase; letter-spacing: 0.06em; font-weight: 700; }}
    h2 {{ font-size: 15px; color: #1E3A8A; border-bottom: 1px solid #EDF2F7; padding-bottom: 6px; margin-top: 32px; }}
    h3 {{ font-size: 12px; text-transform: uppercase; letter-spacing: 0.05em; color: #718096; margin: 16px 0 4px; }}
    .meta {{ font-size: 13px; color: #718096; margin-bottom: 4px; }}
    .badges {{ margin: 10px 0 4px; }}
    .badge {{ display: inline-block; padding: 6px 14px; border-radius: 8px; font-weight: 700; font-size: 13px; margin-right: 8px; }}
    table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12.5px; }}
    th {{ background: #1E3A8A; color: white; text-align: left; padding: 8px 10px; font-size: 11px; text-transform: uppercase; letter-spacing: 0.03em; }}
    td {{ padding: 8px 10px; border-bottom: 1px solid #EDF2F7; vertical-align: top; }}
    ul {{ font-size: 13px; padding-left: 18px; }}
    .rodape {{ margin-top: 40px; padding-top: 14px; border-top: 1px solid #EDF2F7; font-size: 11px; color: #A0AEC0; }}
</style>
</head>
<body>

<div class="cabecalho">
<div class="marca">A</div>
<div>
<h1>Agente Orientador</h1>
<div class="sub">Curadoria de Documentos</div>
</div>
</div>

<div class="meta">Relatório de Diagnóstico Documental</div>
<div class="meta">Gerado em: {gerado_em}</div>

<h2>Origem e finalidade deste relatório</h2>
<h3>Objetivo</h3>
<p style="font-size:13px;">Este relatório foi gerado automaticamente pelo <b>Agente Orientador de Curadoria de Documentos</b>. As diretrizes e critérios de qualidade utilizados nesta avaliação são
definidos e mantidos pela equipe de <b>Data Science Holding (DSH) e
Governança de IA</b>.</p>
<h3>Limitações do Agente</h3>
<ul>
<li>Não altera o documento automaticamente, só aponta melhorias e faz recomendações.</li>
<li>Não substitui a validação humana de conteúdo.</li>
<li>As recomendações geradas são um ponto de partida, não uma decisão final, a área de negócio deve conferir se cada sugestão realmente se aplica ao contexto antes de adotá-la.</li>
</ul>
<h3>Responsabilidades</h3>
<ul>
<li><b>Área dona do conteúdo:</b> responsável por aplicar as recomendações da avaliação e manter a informação atualizada, vigente e aderente aos padrões definidos para uso em IA Generativa;</li>
<li><b>Governança de IA:</b> define os critérios mínimos de qualidade, diretrizes e apoia a disseminação das boas práticas para utilização da IA Generativa.</li>
</ul>

<h2>Diagnóstico do documento: {arquivo}</h2>
<div class="badges">
<span class="badge" style="background:#FBEFDA;color:#B8791A;">Score: {score_info['score']}/100</span>
<span class="badge" style="background:#E4F5EC;color:#1E8A5F;">{html.escape(score_info['classificacao'])}</span>
</div>
<p style="font-size:13px;color:#718096;">A nota (0 a 100) mede a qualidade do conteúdo e formatação. A classificação indica se o documento já está pronto para uso por agentes de IA.</p>

<h3>Resumo do diagnóstico</h3>
<p style="font-size:13.5px;">{resumo}</p>

<h3>Métricas técnicas</h3>
<p style="font-size:12.5px;color:#718096;">{metricas_texto}</p>

<h3>Pontos de conteúdo e formatação</h3>
{tabela_html}

<div class="rodape">
    Relatório gerado automaticamente pelo Agente Orientador de Curadoria de Documentos. As diretrizes de
    curadoria são de responsabilidade da equipe de Data Science Holding
    (DSH) e Governança de IA. Alterações no documento devem ser aplicadas
    pela área dona do conteúdo.
</div>

</body>
</html>"""
