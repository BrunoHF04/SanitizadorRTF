r"""
Sanitização de RTF com lixo de DDE/bookmark ({\*\bkmkstart __DdeLink__...).
Uso em leitura, gravação ou jobs em lote sobre o campo conteudo.
"""

from __future__ import annotations

import json
from pathlib import Path

# Primeira ocorrência deste trecho marca o início do lixo repetitivo (DDE links).
MARKER_DDE_BOOKMARK = r"{\*\bkmkstart __DdeLink__"
DEFAULT_MARKERS = [MARKER_DDE_BOOKMARK]


def _calcular_grupos_abertos(rtf: str) -> int:
    """
    Conta o saldo de grupos RTF abertos por chaves não escapadas.

    Em RTF, somente "{" e "}" não precedidos por "\" delimitam grupos.
    """
    saldo = 0
    i = 0
    n = len(rtf)
    while i < n:
        ch = rtf[i]
        if ch == "\\":
            # Ignora o caractere seguinte (inclui \{, \}, \\ e escapes hex).
            i += 2
            continue
        if ch == "{":
            saldo += 1
        elif ch == "}":
            if saldo > 0:
                saldo -= 1
        i += 1
    return saldo


def _encontrar_primeiro_marcador(conteudo_bruto: str, markers: list[str] | None = None) -> tuple[int, str | None]:
    alvos = [m for m in (markers or DEFAULT_MARKERS) if m]
    melhor_idx = -1
    melhor_marker: str | None = None
    for marker in alvos:
        idx = conteudo_bruto.find(marker)
        if idx == -1:
            continue
        if melhor_idx == -1 or idx < melhor_idx:
            melhor_idx = idx
            melhor_marker = marker
    return melhor_idx, melhor_marker


def limpar_arquivo_rtf(conteudo_bruto: str, markers: list[str] | None = None) -> str:
    """
    Remove tudo após a primeira ocorrência de MARKER_DDE_BOOKMARK e garante
    que os grupos RTF pendentes sejam fechados.
    """
    if not conteudo_bruto:
        return conteudo_bruto

    idx, _ = _encontrar_primeiro_marcador(conteudo_bruto, markers)
    if idx == -1:
        return conteudo_bruto

    conteudo_limpo = conteudo_bruto[:idx].rstrip()
    grupos_abertos = _calcular_grupos_abertos(conteudo_limpo)
    if grupos_abertos > 0:
        conteudo_limpo += "}" * grupos_abertos

    return conteudo_limpo


def validar_estrutura_rtf(conteudo: str) -> bool:
    """Valida estrutura básica de grupos RTF por chaves não escapadas."""
    if not conteudo:
        return False
    if not parece_rtf(conteudo):
        return False
    return _calcular_grupos_abertos(conteudo) == 0


def analisar_limpeza(conteudo_bruto: str, markers: list[str] | None = None) -> dict[str, object]:
    idx, marker = _encontrar_primeiro_marcador(conteudo_bruto, markers)
    limpo = limpar_arquivo_rtf(conteudo_bruto, markers=markers)
    removido = ""
    if idx != -1:
        removido = conteudo_bruto[idx:]
    return {
        "marker_found": idx != -1,
        "marker_used": marker,
        "marker_index": idx,
        "before_len": len(conteudo_bruto or ""),
        "after_len": len(limpo or ""),
        "removed_len": len(removido),
        "removed_preview_start": removido[:220],
        "removed_preview_end": removido[-220:] if removido else "",
        "was_rtf_before": parece_rtf(conteudo_bruto or ""),
        "is_structurally_valid_before": validar_estrutura_rtf(conteudo_bruto or ""),
        "is_structurally_valid_after": validar_estrutura_rtf(limpo or ""),
    }


def carregar_marcadores_de_json(caminho_json: str | None) -> list[str]:
    """
    Carrega marcadores de limpeza a partir de ficheiro JSON.
    Formato esperado: {"markers": ["...","..."]} ou ["...","..."].
    """
    if not caminho_json:
        return list(DEFAULT_MARKERS)
    p = Path(caminho_json)
    if not p.exists():
        return list(DEFAULT_MARKERS)
    obj = json.loads(p.read_text(encoding="utf-8"))
    if isinstance(obj, dict):
        raw = obj.get("markers", [])
    else:
        raw = obj
    if not isinstance(raw, list):
        return list(DEFAULT_MARKERS)
    markers = [str(x).strip() for x in raw if str(x).strip()]
    return markers or list(DEFAULT_MARKERS)


def precisa_limpeza(conteudo: str, tamanho_minimo_alerta: int = 1_000_000) -> bool:
    """True se há marcador de lixo ou tamanho acima do limiar (vale rodar a limpeza)."""
    if not conteudo:
        return False
    for marker in DEFAULT_MARKERS:
        if marker in conteudo:
            return True
    return len(conteudo) >= tamanho_minimo_alerta


def parece_rtf(conteudo: str) -> bool:
    r"""Detecta RTF pelos prefixos usuais ({\rtf ou {\urtf)."""
    s = conteudo.lstrip("\ufeff \t\r\n")
    low = s[:24].lower()
    return low.startswith("{\\rtf") or low.startswith("{\\urtf")


__all__ = [
    "MARKER_DDE_BOOKMARK",
    "DEFAULT_MARKERS",
    "analisar_limpeza",
    "carregar_marcadores_de_json",
    "limpar_arquivo_rtf",
    "precisa_limpeza",
    "parece_rtf",
    "validar_estrutura_rtf",
]
