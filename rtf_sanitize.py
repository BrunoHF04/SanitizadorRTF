r"""
Sanitização de RTF com lixo de DDE/bookmark ({\*\bkmkstart __DdeLink__...).
Uso em leitura, gravação ou jobs em lote sobre o campo conteudo.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

# Primeira ocorrência deste trecho marca o início do lixo repetitivo (DDE links).
MARKER_DDE_BOOKMARK = r"{\*\bkmkstart __DdeLink__"
# Marcadores usados em precisa_limpeza / SELECT no banco e no fallback de corte tardio.
DEFAULT_MARKERS = [
    MARKER_DDE_BOOKMARK,
    r"{\*\shppict",  # Formas + imagem (Word/LibreOffice)
    r"{\nonshppict",  # Fallback WMF/PNG após shppict (ficheiros muito grandes)
    r"{\pict{",  # Grupo \pict sem \* (filho de shppict/nonshppict ou solto)
    r"{\*\objdata",  # Objetos OLE (planilhas coladas etc.) que incham
    r"{\*\pict",  # Destino \*\pict
]
# Remoção de grupos pict/obj/shppict pode apagar imagens — só automática acima deste tamanho.
MIN_BYTES_HEAVY_RTF_CLEANUP = 5 * 1024 * 1024
# Bloco contíguo só com hex (sem estrutura RTF no meio) típico de lixo/corrupção.
HEX_ORPHAN_MIN_RUN = 500_000

# Ordem: envolver “shape” antes do pict interno.
_HEAVY_GROUP_PREFIXES = (
    r"{\*\shppict",
    r"{\nonshppict",
    r"{\*\objdata",
    r"{\*\pict",
    r"{\pict{",
)
_RE_DDE_BKMK = re.compile(r"\{\\\*\\bkmk(?:start|end)\s+__DdeLink__[^{}]*\}")
SAFE_LEVEL = "seguro"
INTERMEDIATE_LEVEL = "intermediario"
AGGRESSIVE_LEVEL = "agressivo"

_INTERMEDIATE_GROUP_PREFIXES = (
    r"{\*\generator",
    r"{\*\userprops",
    r"{\*\xmlnstbl",
    r"{\*\rsidtbl",
    r"{\*\themedata",
    r"{\*\colorschememapping",
)


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


def _find_group_end(text: str, start_idx: int) -> int:
    depth = 0
    i = start_idx
    n = len(text)
    while i < n:
        ch = text[i]
        if ch == "\\":
            i += 2
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return -1


def _remove_groups_by_prefixes(text: str, prefixes: tuple[str, ...]) -> str:
    out = text
    changed = True
    while changed:
        changed = False
        for pfx in prefixes:
            idx = out.find(pfx)
            if idx == -1:
                continue
            end = _find_group_end(out, idx)
            if end == -1:
                continue
            out = out[:idx] + out[end + 1 :]
            changed = True
    return out


def _find_hex_orphan_run_start(text: str, min_run: int = HEX_ORPHAN_MIN_RUN) -> int | None:
    """
    Localiza o início do primeiro trecho com >= min_run dígitos hex [0-9A-Fa-f].
    Espaço, tab e fins de linha entre dígitos são ignorados (pngblip com quebras de linha).
    """
    if min_run <= 0 or len(text) < min_run:
        return None
    i = 0
    n = len(text)
    run_start = 0
    run_len = 0
    hex_set = frozenset("0123456789abcdefABCDEF")
    ws = frozenset(" \t\r\n\f\v")
    while i < n:
        ch = text[i]
        if ch in hex_set:
            if run_len == 0:
                run_start = i
            run_len += 1
            if run_len >= min_run:
                return run_start
            i += 1
        elif ch in ws:
            i += 1
        else:
            run_len = 0
            i += 1
    return None


def _truncar_balanceando_grupos(texto: str, corte: int) -> str:
    truncado = texto[:corte].rstrip()
    grupos_abertos = _calcular_grupos_abertos(truncado)
    if grupos_abertos > 0:
        truncado += "}" * grupos_abertos
    return truncado


def limpar_arquivo_rtf(
    conteudo_bruto: str,
    markers: list[str] | None = None,
    cleaning_level: str = SAFE_LEVEL,
) -> str:
    """
    Remove blocos DDE/bookmark; em níveis superiores, metadados RTF auxiliares.

    Em modo agressivo, com conteúdo acima de ~5 MB, remove também grupos de
    imagem/OLE: \\*\\shppict, nonshppict, \\*\\pict, \\pict{ e \\*\\objdata.
    Nesse nível, detecta “massas” hexadecimais órfãs (>=500k dígitos hex, com
    espaços/newlines opcionais entre dígitos)
    e trunca antes delas. O fallback de corte tardio usa o primeiro marcador de
    `markers` (ou DEFAULT_MARKERS) que apareça nos últimos 10%% do ficheiro.

    Rollback: em base de dados, `db_sanitize` grava o conteúdo anterior em
    `rtf_sanitize_audit` (old_content), permitindo reverter lotes agressivos.
    """
    if not conteudo_bruto:
        return conteudo_bruto

    alvo_marcadores = [m for m in (markers or DEFAULT_MARKERS) if m]
    conteudo_limpo = conteudo_bruto

    # Regra principal: remove apenas os grupos de bookmark DDE.
    conteudo_limpo = _RE_DDE_BKMK.sub("", conteudo_limpo)

    if cleaning_level in (INTERMEDIATE_LEVEL, AGGRESSIVE_LEVEL):
        conteudo_limpo = _remove_groups_by_prefixes(conteudo_limpo, _INTERMEDIATE_GROUP_PREFIXES)

    if cleaning_level == AGGRESSIVE_LEVEL and len(conteudo_limpo) > MIN_BYTES_HEAVY_RTF_CLEANUP:
        conteudo_limpo = _remove_groups_by_prefixes(conteudo_limpo, _HEAVY_GROUP_PREFIXES)

    if cleaning_level == AGGRESSIVE_LEVEL:
        hex_cut = _find_hex_orphan_run_start(conteudo_limpo, HEX_ORPHAN_MIN_RUN)
        if hex_cut is not None and hex_cut > 0:
            conteudo_limpo = _truncar_balanceando_grupos(conteudo_limpo, hex_cut)

    # Fallback conservador: corte se algum marcador aparecer muito perto do fim.
    idx, _ = _encontrar_primeiro_marcador(conteudo_limpo, alvo_marcadores)
    if idx != -1:
        ratio = idx / max(len(conteudo_limpo), 1)
        if ratio >= 0.90:
            conteudo_limpo = _truncar_balanceando_grupos(conteudo_limpo, idx)

    return conteudo_limpo


def validar_estrutura_rtf(conteudo: str) -> bool:
    """Valida estrutura básica de grupos RTF por chaves não escapadas."""
    if not conteudo:
        return False
    if not parece_rtf(conteudo):
        return False
    return _calcular_grupos_abertos(conteudo) == 0


def analisar_limpeza(
    conteudo_bruto: str,
    markers: list[str] | None = None,
    cleaning_level: str = SAFE_LEVEL,
) -> dict[str, object]:
    idx, marker = _encontrar_primeiro_marcador(conteudo_bruto, markers)
    limpo = limpar_arquivo_rtf(conteudo_bruto, markers=markers, cleaning_level=cleaning_level)
    # Prévia por diferença de comprimento (evita assumir truncamento total).
    removed_len = max(len(conteudo_bruto) - len(limpo), 0)
    preview_start = ""
    preview_end = ""
    if removed_len > 0 and idx != -1:
        preview_start = conteudo_bruto[idx : idx + 220]
        preview_end = conteudo_bruto[max(len(conteudo_bruto) - 220, 0) :]
    return {
        "marker_found": idx != -1,
        "marker_used": marker,
        "marker_index": idx,
        "before_len": len(conteudo_bruto or ""),
        "after_len": len(limpo or ""),
        "removed_len": removed_len,
        "removed_preview_start": preview_start,
        "removed_preview_end": preview_end,
        "was_rtf_before": parece_rtf(conteudo_bruto or ""),
        "is_structurally_valid_before": validar_estrutura_rtf(conteudo_bruto or ""),
        "is_structurally_valid_after": validar_estrutura_rtf(limpo or ""),
        "cleaning_level": cleaning_level,
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
    "MIN_BYTES_HEAVY_RTF_CLEANUP",
    "HEX_ORPHAN_MIN_RUN",
    "SAFE_LEVEL",
    "INTERMEDIATE_LEVEL",
    "AGGRESSIVE_LEVEL",
    "analisar_limpeza",
    "carregar_marcadores_de_json",
    "limpar_arquivo_rtf",
    "precisa_limpeza",
    "parece_rtf",
    "validar_estrutura_rtf",
]
