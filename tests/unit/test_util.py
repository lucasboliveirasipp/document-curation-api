# tests/unit/test_util.py
"""
Testa as funções auxiliares de util.py, incluindo casos de borda como
entrada vazia, None e valores muito longos.
"""
import pytest

from src import util


# ─── sanitize_header_value ──────────────────────────────────────────────────

def test_sanitize_none_retorna_vazio():
    assert util.sanitize_header_value(None) == ""


def test_sanitize_remove_cr_lf():
    assert util.sanitize_header_value("linha1\r\nlinha2\r\n") == "linha1linha2"


def test_sanitize_remove_caracteres_de_controle_preserva_unicode():
    bruto = "\x01\x02Olá, mundo!\x1f café ☕"
    resultado = util.sanitize_header_value(bruto)
    assert "\x01" not in resultado
    assert "\x1f" not in resultado
    assert "café" in resultado
    assert "☕" in resultado


def test_sanitize_faz_strip_nas_bordas():
    assert util.sanitize_header_value("   Bearer token   ") == "Bearer token"


def test_sanitize_trunca_em_8192():
    valor_longo = "a" * 9000
    resultado = util.sanitize_header_value(valor_longo)
    assert len(resultado) == util.MAX_HEADER_LEN


def test_sanitize_converte_nao_string_para_string():
    assert util.sanitize_header_value(12345) == "12345"


# ─── get_file_extension ─────────────────────────────────────────────────────

@pytest.mark.parametrize(
    "filename,esperado",
    [
        ("documento.pdf", "pdf"),
        ("RELATORIO.PDF", "pdf"),
        ("arquivo.tar.gz", "gz"),
        ("sem_extensao", None),
        (".gitignore", None),
        ("", None),
        (None, None),
        (123, None),
    ],
)
def test_get_file_extension_casos(filename, esperado):
    assert util.get_file_extension(filename) == esperado
