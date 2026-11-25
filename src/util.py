"""
util.py

Caixa de ferramentas com funções pequenas e genéricas que mais de um
arquivo do projeto precisa usar — por isso não ficam dentro de um
service específico. Não guardam estado nem fazem chamada de rede, só
recebem um valor e devolvem outro.
"""
from typing import Any, Optional
import os

MAX_HEADER_LEN = 8192


def sanitize_header_value(value: Optional[Any]) -> str:
    """
    Sanitiza valor destinado a header HTTP (ex: token de autenticação
    antes de virar "Authorization: Bearer <token>"). Existe para
    prevenir CRLF/header injection caso o valor recebido da Gateway
    venha corrompido ou manipulado — remove quebras de linha e
    caracteres de controle, preserva acentuação/Unicode (útil em
    mensagens de erro) e trunca em MAX_HEADER_LEN para não estourar
    limite de tamanho de header em servidores HTTP.
    """
    if value is None:
        return ""

    s = str(value)
    s = s.replace("\r", "").replace("\n", "")
    s = "".join(ch for ch in s if (0x20 <= ord(ch) <= 0x7E) or (ord(ch) >= 0x80))
    s = s.strip()

    if len(s) > MAX_HEADER_LEN:
        s = s[:MAX_HEADER_LEN]

    return s


def get_file_extension(filename):
    """Extrai a extensão de um nome de arquivo em minúsculas, sem o
    ponto (ex: "Relatorio.PDF" -> "pdf"). Usada pelo controller para
    validar contra ALLOWED_EXTENSIONS antes de processar o upload.
    Devolve None para entrada inválida (None, vazio, não-string) ou
    sem extensão (ex: ".gitignore", que o Path trata como nome sem
    extensão, não como extensão "gitignore")."""
    if not filename or not isinstance(filename, str):
        return None

    _, ext = os.path.splitext(filename)
    return ext[1:].lower() if ext else None
