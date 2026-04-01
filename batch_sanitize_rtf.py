#!/usr/bin/env python3
"""
Manutenção em lote: documento_mesclado.conteudo (PostgreSQL).

Critério de seleção padrão: LENGTH(conteudo) > --min-length OU conteúdo contém
o marcador DDE (para não perder registros médios já corrompidos).

Variáveis de ambiente (exemplo):
  PGHOST, PGPORT, PGDATABASE, PGUSER, PGPASSWORD
  ou DATABASE_URL=postgresql://user:pass@host:5432/dbname

Uso:
  python batch_sanitize_rtf.py
  python batch_sanitize_rtf.py --execute --min-length 1000000

O SELECT usa cursor com itersize=1 para não acumular todos os registros na memória.
Se conteudo for bytea, ajuste o SQL (decode) conforme o encoding usado.
"""

from __future__ import annotations

import argparse
import os
import sys
from typing import Any

import re as _re

from rtf_sanitize import MARKER_DDE_BOOKMARK, limpar_arquivo_rtf, parece_rtf


def _valid_sql_identifier(name: str) -> bool:
    return bool(_re.fullmatch(r"[a-zA-Z_][a-zA-Z0-9_]*", name))


def get_connection():
    try:
        import psycopg2
    except ImportError as e:
        raise SystemExit(
            "Instale dependências: pip install -r requirements.txt"
        ) from e

    url = os.environ.get("DATABASE_URL")
    if url:
        return psycopg2.connect(url)
    return psycopg2.connect(
        host=os.environ.get("PGHOST", "localhost"),
        port=os.environ.get("PGPORT", "5432"),
        dbname=os.environ.get("PGDATABASE", ""),
        user=os.environ.get("PGUSER", ""),
        password=os.environ.get("PGPASSWORD", ""),
    )


def iter_candidates(
    conn: Any,
    id_column: str,
    min_length: int,
    only_rtf: bool,
    limit: int | None,
) -> Any:
    """
    Cursor do servidor (streaming): tuplas (pk, conteudo, len).
    id_column já validado como identificador SQL.
    """
    cond_length = "LENGTH(conteudo::text) > %s"
    cond_marker = "position(%s in conteudo::text) > 0"
    where = f"(({cond_length}) OR ({cond_marker}))"
    params: list[Any] = [min_length, MARKER_DDE_BOOKMARK]

    sql = f"""
        SELECT {id_column}, conteudo::text, LENGTH(conteudo::text) AS len
        FROM documento_mesclado
        WHERE conteudo IS NOT NULL
          AND ({where})
        ORDER BY len DESC
    """
    if limit is not None:
        sql += " LIMIT %s"
        params.append(limit)

    cur = conn.cursor(name="sanitize_rtf_stream")
    cur.itersize = 1
    cur.execute(sql, params)
    try:
        for row in cur:
            if only_rtf and not parece_rtf((row[1] or "")):
                continue
            yield row
    finally:
        cur.close()


def main() -> int:
    p = argparse.ArgumentParser(description="Limpa lixo DDE em RTF (documento_mesclado).")
    p.add_argument("--execute", action="store_true", help="Aplicar UPDATE (sem isto: dry-run)")
    p.add_argument("--min-length", type=int, default=1_000_000, metavar="N")
    p.add_argument("--only-rtf", action="store_true", help="Ignorar linhas que não parecem RTF")
    p.add_argument("--limit", type=int, default=None, help="Máximo de registros a processar")
    p.add_argument(
        "--id-column",
        default="id",
        help="Nome da coluna PK (default: id)",
    )
    args = p.parse_args()
    if not _valid_sql_identifier(args.id_column):
        print("--id-column deve ser um identificador SQL simples (ex.: id).", file=sys.stderr)
        return 2

    conn = get_connection()
    try:
        conn.autocommit = False
        upd_cur = conn.cursor()
        try:
            updated = 0
            skipped = 0
            for row_id, conteudo, length in iter_candidates(
                conn,
                args.id_column,
                args.min_length,
                args.only_rtf,
                args.limit,
            ):
                if not conteudo:
                    skipped += 1
                    continue
                limpo = limpar_arquivo_rtf(conteudo)
                if limpo == conteudo:
                    skipped += 1
                    continue
                novo_len = len(limpo)
                print(
                    f"{args.id_column}={row_id} len {length} -> {novo_len} "
                    f"({'UPDATE' if args.execute else 'WOULD UPDATE'})"
                )
                if args.execute:
                    col = args.id_column
                    upd_cur.execute(
                        f"UPDATE documento_mesclado SET conteudo = %s WHERE {col} = %s",
                        (limpo, row_id),
                    )
                    updated += upd_cur.rowcount
                else:
                    updated += 1

            if args.execute:
                conn.commit()
            else:
                conn.rollback()
        finally:
            upd_cur.close()

        print(f"Feito. Alterados (contagem): {updated}. Ignorados sem mudança: {skipped}.")
    finally:
        conn.close()

    return 0


if __name__ == "__main__":
    sys.exit(main())
