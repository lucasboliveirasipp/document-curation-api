"""
main.py

Ponto de partida da aplicação: aqui a API FastAPI é criada, o CORS é
configurado e as rotas de curadoria e de health check são conectadas.
É o arquivo que sobe quando o servidor inicia.
"""

# Carrega variáveis do arquivo .env mais próximo (pasta do projeto)
def _carregar_env():
    from pathlib import Path
    try:
        from dotenv import load_dotenv
    except ImportError:
        return
    for pasta in Path(__file__).resolve().parents:
        if (pasta / ".env").exists():
            load_dotenv(pasta / ".env")
            break


_carregar_env()

import os
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from src.controllers.curadoria_controller import curadoria_route
from src.controllers.health_check_controller import health_check_route

CORS_ORIGINS = [origem.strip() for origem in os.getenv("CORS_ORIGINS", "").split(",") if origem.strip()]

app = FastAPI(
    title="BFF Agente Orientador Curadoria de Documentos",
    version=os.getenv("APP_VERSION", "1.0.0"),
    root_path=os.getenv("APP_ROOT_PATH", "/api-agente-orientador-curadoria"),
    docs_url=os.getenv("DOCS_URL", "/docs"),
    redoc_url=os.getenv("REDOC_URL", "/redoc"),
    openapi_url="/openapi.json"
)


app.add_middleware(CORSMiddleware,
                   allow_credentials=True,
                   allow_origins=CORS_ORIGINS,
                   allow_methods=["GET", "POST"],
                   allow_headers=["Authorization", "Content-Type"])

app.include_router(health_check_route)
app.include_router(curadoria_route)