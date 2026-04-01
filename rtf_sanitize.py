r"""
Sanitização de RTF com lixo de DDE/bookmark ({\*\bkmkstart __DdeLink__...).
Uso em leitura, gravação ou jobs em lote sobre o campo conteudo.
"""

from __future__ import annotations

# Primeira ocorrência deste trecho marca o início do lixo repetitivo (DDE links).
MARKER_DDE_BOOKMARK = r"{\*\bkmkstart __DdeLink__"


def limpar_arquivo_rtf(conteudo_bruto: str) -> str:
    """
    Remove tudo após a primeira ocorrência de MARKER_DDE_BOOKMARK e garante
    que o RTF termine com '}'.
    """
    if not conteudo_bruto:
        return conteudo_bruto

    idx = conteudo_bruto.find(MARKER_DDE_BOOKMARK)
    if idx == -1:
        return conteudo_bruto

    conteudo_limpo = conteudo_bruto[:idx].rstrip()

    if not conteudo_limpo.endswith("}"):
        conteudo_limpo += "}"

    return conteudo_limpo


def precisa_limpeza(conteudo: str, tamanho_minimo_alerta: int = 1_000_000) -> bool:
    """True se há marcador de lixo ou tamanho acima do limiar (vale rodar a limpeza)."""
    if not conteudo:
        return False
    if MARKER_DDE_BOOKMARK in conteudo:
        return True
    return len(conteudo) >= tamanho_minimo_alerta


def parece_rtf(conteudo: str) -> bool:
    r"""Detecta RTF pelos prefixos usuais ({\rtf ou {\urtf)."""
    s = conteudo.lstrip("\ufeff \t\r\n")
    low = s[:24].lower()
    return low.startswith("{\\rtf") or low.startswith("{\\urtf")


__all__ = [
    "MARKER_DDE_BOOKMARK",
    "limpar_arquivo_rtf",
    "precisa_limpeza",
    "parece_rtf",
]
