from __future__ import annotations

import json
import re as _re
from typing import Callable
from uuid import uuid4

from rtf_sanitize import MARKER_DDE_BOOKMARK, limpar_arquivo_rtf, parece_rtf


def _valid_sql_identifier(name: str) -> bool:
    return bool(_re.fullmatch(r"[a-zA-Z_][a-zA-Z0-9_]*", name))


def _valid_sql_table_name(name: str) -> bool:
    parts = name.split(".")
    if not (1 <= len(parts) <= 2):
        return False
    return all(_valid_sql_identifier(p) for p in parts)


def sanitize_documento_mesclado(
    database_url: str,
    *,
    execute: bool,
    min_length: int = 1_000_000,
    min_megabytes: float | None = None,
    full_scan: bool = False,
    only_rtf: bool = False,
    limit: int | None = None,
    table_name: str = "documento_mesclado",
    content_column: str = "conteudo",
    id_column: str | None = None,
    report_columns: list[str] | None = None,
    log: Callable[[str], None] | None = None,
) -> tuple[int, int, str | None]:
    """
    Higieniza documento_mesclado.conteudo no PostgreSQL.
    Retorna (alterados_ou_seriam_alterados, ignorados_sem_mudanca).
    """
    if not database_url.strip():
        raise ValueError("DATABASE_URL é obrigatório.")
    if not _valid_sql_table_name(table_name):
        raise ValueError("table_name inválido. Use 'tabela' ou 'schema.tabela'.")
    if not _valid_sql_identifier(content_column):
        raise ValueError("content_column inválido. Use apenas identificador simples.")
    if id_column is not None and not _valid_sql_identifier(id_column):
        raise ValueError("id_column inválido. Use apenas identificador simples (ex.: id).")
    report_columns = report_columns or []
    if min_megabytes is not None and min_megabytes < 0:
        raise ValueError("min_megabytes deve ser >= 0.")
    for col in report_columns:
        if "." in col:
            parts = col.split(".")
            if len(parts) < 2:
                raise ValueError(f"report_columns contém formato inválido: {col}")
            rel_table = ".".join(parts[:-1])
            rel_col = parts[-1]
            if not _valid_sql_table_name(rel_table) or not _valid_sql_identifier(rel_col):
                raise ValueError(f"report_columns contém coluna inválida: {col}")
        elif not _valid_sql_identifier(col):
            raise ValueError(f"report_columns contém coluna inválida: {col}")

    try:
        import psycopg2
    except ImportError as e:
        raise RuntimeError("Instale dependências: pip install -r requirements.txt") from e

    def _log(msg: str) -> None:
        if log:
            log(msg)

    conn = psycopg2.connect(database_url.strip())
    try:
        conn.autocommit = False
        if id_column is None:
            pks = list_postgres_primary_keys(database_url, table_name)
            if not pks:
                raise ValueError(
                    f"A tabela {table_name} não tem PK. Defina uma coluna de chave para permitir update/rollback."
                )
            id_column = pks[0]

        if execute:
            _ensure_audit_table(conn)
            batch_id = str(uuid4())
        else:
            batch_id = None

        report_selects: list[str] = []
        for idx, rc in enumerate(report_columns):
            alias = f"rep_{idx}"
            if "." in rc:
                parts = rc.split(".")
                rel_table = ".".join(parts[:-1])
                rel_col = parts[-1]
                # Busca valor relacionado pela mesma chave (id_column) quando disponível.
                report_selects.append(
                    f"(SELECT rel.{rel_col}::text FROM {rel_table} rel "
                    f"WHERE rel.{id_column}::text = src.{id_column}::text "
                    f"ORDER BY rel.{rel_col} NULLS LAST LIMIT 1) AS {alias}"
                )
            else:
                report_selects.append(f"src.{rc}::text AS {alias}")
        select_report = (", " + ", ".join(report_selects)) if report_selects else ""
        params: list[object] = []
        if full_scan:
            where_sql = "1=1"
        else:
            where_parts = ["position(%s in src.{content_column}::text) > 0".format(content_column=content_column)]
            params.append(MARKER_DDE_BOOKMARK)
            if min_length > 0:
                where_parts.append("LENGTH(src.{content_column}::text) > %s".format(content_column=content_column))
                params.append(min_length)
            if min_megabytes is not None and min_megabytes > 0:
                where_parts.append("OCTET_LENGTH(src.{content_column}::text) > %s".format(content_column=content_column))
                params.append(int(min_megabytes * 1024 * 1024))
            where_sql = " OR ".join(where_parts)

        sql = f"""
            SELECT {id_column}::text AS row_ref, {content_column}::text, LENGTH({content_column}::text) AS len
            {select_report}
            FROM {table_name} src
            WHERE {content_column} IS NOT NULL
              AND ({where_sql})
            ORDER BY len DESC
        """
        if limit is not None:
            sql += " LIMIT %s"
            params.append(limit)

        read_cur = conn.cursor(name="sanitize_rtf_stream")
        read_cur.itersize = 1
        upd_cur = conn.cursor()
        updated = 0
        skipped = 0
        try:
            read_cur.execute(sql, params)
            for row in read_cur:
                row_id, conteudo, length = row[0], row[1], row[2]
                report_values = row[3:]
                if only_rtf and not parece_rtf(conteudo or ""):
                    skipped += 1
                    continue
                if not conteudo:
                    skipped += 1
                    continue
                limpo = limpar_arquivo_rtf(conteudo)
                if limpo == conteudo:
                    skipped += 1
                    continue

                _log(
                    f"ref={row_id} len {length} -> {len(limpo)} "
                    f"({'UPDATE' if execute else 'WOULD UPDATE'})"
                )
                if execute:
                    report_data = {}
                    for i, col in enumerate(report_columns):
                        report_data[col] = report_values[i] if i < len(report_values) else None
                    upd_cur.execute(
                        """
                        INSERT INTO rtf_sanitize_audit
                        (batch_id, table_name, key_column, key_value, content_column, report_data, old_content, new_content, old_len, new_len)
                        VALUES (%s, %s, %s, %s, %s, %s::jsonb, %s, %s, %s, %s)
                        """,
                        (
                            batch_id,
                            table_name,
                            id_column,
                            row_id,
                            content_column,
                            json.dumps(report_data, ensure_ascii=False),
                            conteudo,
                            limpo,
                            len(conteudo),
                            len(limpo),
                        ),
                    )
                    upd_cur.execute(
                        f"UPDATE {table_name} SET {content_column} = %s WHERE {id_column}::text = %s",
                        (limpo, row_id),
                    )
                    updated += upd_cur.rowcount
                else:
                    updated += 1
        finally:
            read_cur.close()
            upd_cur.close()

        if execute:
            conn.commit()
        else:
            conn.rollback()
        return updated, skipped, batch_id
    finally:
        conn.close()


def _ensure_audit_table(conn: object) -> None:
    with conn.cursor() as cur:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS rtf_sanitize_audit (
                id BIGSERIAL PRIMARY KEY,
                batch_id TEXT NOT NULL,
                table_name TEXT NOT NULL,
                key_column TEXT NOT NULL,
                key_value TEXT NOT NULL,
                content_column TEXT NOT NULL,
                report_data JSONB,
                old_content TEXT NOT NULL,
                new_content TEXT NOT NULL,
                old_len INTEGER NOT NULL,
                new_len INTEGER NOT NULL,
                changed_at TIMESTAMPTZ NOT NULL DEFAULT now()
            )
            """
        )
        cur.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_rtf_sanitize_audit_batch
            ON rtf_sanitize_audit(batch_id, id)
            """
        )


def test_postgres_connection(database_url: str) -> str:
    """Testa conexão e retorna resumo amigável."""
    if not database_url.strip():
        raise ValueError("DATABASE_URL é obrigatório.")
    try:
        import psycopg2
    except ImportError as e:
        raise RuntimeError("Instale dependências: pip install -r requirements.txt") from e

    conn = psycopg2.connect(database_url.strip())
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT current_database(), current_user")
            dbname, dbuser = cur.fetchone()
        return f"Conexão OK. Banco: {dbname} | Usuário: {dbuser}"
    finally:
        conn.close()


def list_postgres_tables(database_url: str) -> list[str]:
    """Lista tabelas base (schema.tabela), excluindo schemas de sistema."""
    if not database_url.strip():
        raise ValueError("DATABASE_URL é obrigatório.")
    try:
        import psycopg2
    except ImportError as e:
        raise RuntimeError("Instale dependências: pip install -r requirements.txt") from e

    conn = psycopg2.connect(database_url.strip())
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT table_schema, table_name
                FROM information_schema.tables
                WHERE table_type = 'BASE TABLE'
                  AND table_schema NOT IN ('pg_catalog', 'information_schema')
                ORDER BY table_schema, table_name
                """
            )
            return [f"{schema}.{table}" for schema, table in cur.fetchall()]
    finally:
        conn.close()


def list_postgres_columns(database_url: str, table_name: str) -> list[str]:
    """Lista colunas da tabela informada (tabela ou schema.tabela)."""
    if not database_url.strip():
        raise ValueError("DATABASE_URL é obrigatório.")
    if not _valid_sql_table_name(table_name):
        raise ValueError("table_name inválido. Use 'tabela' ou 'schema.tabela'.")

    parts = table_name.split(".")
    if len(parts) == 2:
        schema_name, plain_table = parts
    else:
        schema_name, plain_table = "public", parts[0]

    try:
        import psycopg2
    except ImportError as e:
        raise RuntimeError("Instale dependências: pip install -r requirements.txt") from e

    conn = psycopg2.connect(database_url.strip())
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT column_name
                FROM information_schema.columns
                WHERE table_schema = %s
                  AND table_name = %s
                ORDER BY ordinal_position
                """,
                (schema_name, plain_table),
            )
            return [row[0] for row in cur.fetchall()]
    finally:
        conn.close()


def list_postgres_primary_keys(database_url: str, table_name: str) -> list[str]:
    """Lista colunas da PK da tabela (ordem da constraint)."""
    if not database_url.strip():
        raise ValueError("DATABASE_URL é obrigatório.")
    if not _valid_sql_table_name(table_name):
        raise ValueError("table_name inválido. Use 'tabela' ou 'schema.tabela'.")

    parts = table_name.split(".")
    if len(parts) == 2:
        schema_name, plain_table = parts
    else:
        schema_name, plain_table = "public", parts[0]

    try:
        import psycopg2
    except ImportError as e:
        raise RuntimeError("Instale dependências: pip install -r requirements.txt") from e

    conn = psycopg2.connect(database_url.strip())
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT kcu.column_name
                FROM information_schema.table_constraints tc
                JOIN information_schema.key_column_usage kcu
                  ON tc.constraint_name = kcu.constraint_name
                 AND tc.table_schema = kcu.table_schema
                 AND tc.table_name = kcu.table_name
                WHERE tc.constraint_type = 'PRIMARY KEY'
                  AND tc.table_schema = %s
                  AND tc.table_name = %s
                ORDER BY kcu.ordinal_position
                """,
                (schema_name, plain_table),
            )
            return [row[0] for row in cur.fetchall()]
    finally:
        conn.close()


def get_batch_report(database_url: str, batch_id: str, limit: int = 500) -> list[dict]:
    if not database_url.strip():
        raise ValueError("DATABASE_URL é obrigatório.")
    if not batch_id.strip():
        raise ValueError("batch_id é obrigatório.")
    try:
        import psycopg2
    except ImportError as e:
        raise RuntimeError("Instale dependências: pip install -r requirements.txt") from e

    conn = psycopg2.connect(database_url.strip())
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT table_name, key_column, key_value, content_column, old_len, new_len, report_data, changed_at
                FROM rtf_sanitize_audit
                WHERE batch_id = %s
                ORDER BY id
                LIMIT %s
                """,
                (batch_id, limit),
            )
            rows = cur.fetchall()
            out = []
            for r in rows:
                out.append(
                    {
                        "table_name": r[0],
                        "key_column": r[1],
                        "key_value": r[2],
                        "content_column": r[3],
                        "old_len": r[4],
                        "new_len": r[5],
                        "report_data": r[6] or {},
                        "changed_at": str(r[7]),
                    }
                )
            return out
    finally:
        conn.close()


def rollback_batch(database_url: str, batch_id: str) -> int:
    if not database_url.strip():
        raise ValueError("DATABASE_URL é obrigatório.")
    if not batch_id.strip():
        raise ValueError("batch_id é obrigatório.")
    try:
        import psycopg2
    except ImportError as e:
        raise RuntimeError("Instale dependências: pip install -r requirements.txt") from e

    conn = psycopg2.connect(database_url.strip())
    try:
        conn.autocommit = False
        updated = 0
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT table_name, key_column, key_value, content_column, old_content
                FROM rtf_sanitize_audit
                WHERE batch_id = %s
                ORDER BY id DESC
                """,
                (batch_id,),
            )
            rows = cur.fetchall()
            if not rows:
                conn.rollback()
                return 0

            for table_name, key_column, key_value, content_column, old_content in rows:
                if not _valid_sql_table_name(table_name):
                    raise ValueError(f"Tabela inválida no audit: {table_name}")
                if not _valid_sql_identifier(key_column) or not _valid_sql_identifier(content_column):
                    raise ValueError("Coluna inválida no audit.")
                cur.execute(
                    f"UPDATE {table_name} SET {content_column} = %s WHERE {key_column}::text = %s",
                    (old_content, key_value),
                )
                updated += cur.rowcount

        conn.commit()
        return updated
    finally:
        conn.close()
