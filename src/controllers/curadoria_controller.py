"""
curadoria_controller.py

Aqui ficam as rotas HTTP que o frontend chama para enviar documentos e
pedir o relatório. Este arquivo só cuida da "porta de entrada": recebe
o upload, confere se a quantidade/extensão/tamanho dos arquivos está
dentro do permitido, e delega o trabalho pesado (extrair texto, avaliar
com IA, calcular nota) para curadoria_service.
"""
import os
from typing import List

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from fastapi.responses import HTMLResponse
from pydantic import BaseModel

from src.services.gateway_apim_service import GatewayClient
from src.services import curadoria_service
from src.util import get_file_extension

curadoria_route = APIRouter()

MAX_FILES_PER_BATCH = int(os.getenv("MAX_FILES_PER_BATCH", "5"))
MAX_FILE_SIZE_MB = int(os.getenv("MAX_FILE_SIZE_MB", "20"))
ALLOWED_EXTENSIONS = {"pdf", "docx", "xlsx", "xlsm", "pptx"}


@curadoria_route.post(
    "/avaliar_documentos",
    tags=["Curadoria"],
    summary="Avalia documentos contra a rubrica de qualidade documental",
)
async def api_avaliar_documentos(
    files: List[UploadFile] = File(...),
    client: GatewayClient = Depends(GatewayClient),
):
    try:
        # Validações do lote inteiro primeiro (falha rápido, sem ler
        # bytes de nenhum arquivo) antes de validar arquivo por arquivo.
        if len(files) > MAX_FILES_PER_BATCH:
            raise HTTPException(
                status_code=400,
                detail=f"É permitido enviar no máximo {MAX_FILES_PER_BATCH} arquivos por vez.",
            )

        arquivos_lidos = []
        for file in files:
            filename = file.filename
            ext = get_file_extension(filename)

            if ext not in ALLOWED_EXTENSIONS:
                raise HTTPException(status_code=400, detail=f"Extensão não permitida: {filename}")

            file_bytes = await file.read()
            file_size_mb = len(file_bytes) / (1024 * 1024)

            if file_size_mb > MAX_FILE_SIZE_MB:
                raise HTTPException(
                    status_code=400,
                    detail=f"O arquivo '{filename}' excede o limite de {MAX_FILE_SIZE_MB} MB.",
                )

            arquivos_lidos.append((filename, file_bytes))

        # processar_arquivo nunca propaga exceção (ver docstring dela em
        # curadoria_service.py) — cada item do lote vira {"erro": ...}
        # em caso de falha, então 1 arquivo com problema não derruba os
        # demais nem o HTTP da rota, que segue 200 com o lote completo.
        # processar_lote_arquivos roda os arquivos do upload em paralelo,
        # compartilhando 1 único teto de chamadas simultâneas à Gateway
        # entre todos eles (em vez de cada arquivo abrir seu próprio pool
        # ou de processar 1 arquivo por vez).
        resultados = curadoria_service.processar_lote_arquivos(arquivos_lidos, client)

        return {"resultados": resultados}

    except HTTPException:
        # Repassa como está (400 de validação) sem cair no 500 genérico abaixo.
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao avaliar os documentos: {str(e)}")


class RelatorioRequest(BaseModel):
    resultado: dict


@curadoria_route.post(
    "/relatorio_html",
    tags=["Curadoria"],
    summary="Gera o relatório HTML de diagnóstico de 1 documento já avaliado",
)
def api_relatorio_html(body: RelatorioRequest):
    try:
        html_gerado = curadoria_service.gerar_relatorio_html(body.resultado)
        return HTMLResponse(content=html_gerado)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao gerar relatório: {str(e)}")
