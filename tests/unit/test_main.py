# tests/unit/test_main.py
"""
Confere se a aplicação sobe corretamente: nome e versão certos, CORS
configurado, e as páginas de documentação (/docs, /redoc) acessíveis.
"""
from fastapi.testclient import TestClient

from src.main import app


def test_app_metadata_e_openapi():
    assert app.title == "BFF Agente Orientador Curadoria de Documentos"
    assert app.openapi_url == "/openapi.json"

    client = TestClient(app)
    openapi = client.get("/openapi.json")
    assert openapi.status_code == 200
    assert openapi.json()["info"]["title"] == app.title


def test_docs_e_redoc_acessiveis():
    client = TestClient(app)
    assert client.get("/docs").status_code == 200
    assert client.get("/redoc").status_code == 200


def test_cors_middleware_configurado():
    assert any(m.cls.__name__ == "CORSMiddleware" for m in app.user_middleware)


def test_rotas_de_curadoria_e_health_incluidas():
    client = TestClient(app)
    assert client.get("/health/liveness").status_code == 200
    rotas = {route.path for route in app.routes}
    assert "/avaliar_documentos" in rotas
    assert "/relatorio_html" in rotas
