# tests/conftest.py
from __future__ import annotations
import os
import sys
from pathlib import Path
import pytest

# ==============================================================================
# 1) Descobrir a raiz do repositório e injetar caminhos no sys.path
# ==============================================================================
THIS_FILE = Path(__file__).resolve()
TESTS_DIR = THIS_FILE.parent
ROOT = TESTS_DIR.parent
SRC = ROOT / "src"

paths_to_add = []
if SRC.exists():
    paths_to_add.append(str(SRC))
paths_to_add.append(str(ROOT))

for p in paths_to_add:
    if p not in sys.path:
        sys.path.insert(0, p)

# ==============================================================================
# 2) Variáveis de ambiente padrão para modo de teste
# ==============================================================================
DEFAULT_ENV = {
    "ENV": os.getenv("ENV", "test"),
    "LOG_LEVEL": os.getenv("LOG_LEVEL", "INFO"),
    "CORS_ORIGINS": "https://test.example.com",
    "APIM_WORKFLOW_URL": "https://wf.test.example.com",
    "APIM_STATUS_URL": "https://status.test.example.com",
    "WORKFLOW_CONFIGURATION_CODE": "TEST_WRFL_TEXTGEN",
    "CURATION_MODEL": "test-deployment",
    "OPENAI_API_ENDPOINT": "https://seu-recurso.openai.azure.com/",
    "OPENAI_API_VERSION": "2024-02-01",
    "OPENAI_API_KEY": "test-key",
    "MAX_FILES_PER_BATCH": "5",
    "MAX_FILE_SIZE_MB": "20",
    "MAX_TOKENS_AVALIACAO": "8192",
    "MAX_TOKENS_RESUMO": "1024",
    "TENANTID": "test-tenant",
    "GATEWAYAPPID": "test-client-id",
    "GATEWAYCLIENTSECRET": "test-secret",
    "GATEWAYAPPIDURI": "api://test-app",
}

for k, v in DEFAULT_ENV.items():
    if os.getenv(k) is None:
        os.environ[k] = v

# ==============================================================================
# 3) Fixtures básicas
# ==============================================================================
@pytest.fixture(scope="session")
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture(scope="function")
def client():
    from fastapi.testclient import TestClient
    from src.main import app
    return TestClient(app)


os.environ.setdefault("PYTHONUTF8", "1")
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
