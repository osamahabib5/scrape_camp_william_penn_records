#!/usr/bin/env python3
"""Audit Trail & Multi-User Authentication — Migration Script

Run this script once to set up:

1. ``users_information`` table  — admin panel login credentials
2. ``audit`` schema            — database-level change tracking
3. Trigger function            — captures INSERT / UPDATE / DELETE
4. Triggers on every table     — automatically log all modifications

Usage::

    python implement_audit_trail.py

The script reads ``DB_CONNECTION_STRING`` from the ``.env`` file in the
``backend/`` directory (same file used by the FastAPI application).
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

# ── Ensure backend/ is on the path so we can import `config` ──────────
BACKEND_DIR = Path(__file__).resolve().parent / "backend"
sys.path.insert(0, str(BACKEND_DIR))

from dotenv import load_dotenv  # noqa: E402

load_dotenv(BACKEND_DIR / ".env")

import psycopg2  # noqa: E402
from psycopg2 import sql  # noqa: E402

# ── Configuration ──────────────────────────────────────────────────────
DB_DSN = os.getenv("DB_CONNECTION_STRING", "")
if not DB_DSN:
    print("ERROR: DB_CONNECTION_STRING not set in backend/.env")
    sys.exit(1)

DEFAULT_ADMIN_USER = os.getenv("ADMIN_USERNAME", "admin")
DEFAULT_ADMIN_PASS = os.getenv("ADMIN_PASSWORD", "change-me")

TARGET_TABLES_PK: dict[str, str] = {
    "family_members":   "member_id",
    "locations":        "locations_id",
    "book_of_negroes":  "bon_id",
    "usct_connecticut": "id",
    "rev_war_details":  "rev_war_detail_id",
    "sofafea_members":  "sofafea_member_id",
}

# ── SQL Templates ──────────────────────────────────────────────────────

CREATE_USERS_TABLE = """
CREATE TABLE IF NOT EXISTS public.users_information (
    id              BIGSERIAL PRIMARY KEY,
    username        TEXT        NOT NULL UNIQUE,
    password        TEXT        NOT NULL,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
    role            VARCHAR(255) NOT NULL DEFAULT 'viewer'
);
"""

CREATE_AUDIT_SCHEMA = """
CREATE SCHEMA IF NOT EXISTS audit;
"""

CREATE_AUDIT_TABLE = """
CREATE TABLE IF NOT EXISTS audit.logged_actions (
    id              BIGSERIAL PRIMARY KEY,
    schema_name     TEXT        NOT NULL DEFAULT 'public',
    table_name      TEXT        NOT NULL,
    action          TEXT        NOT NULL CHECK (action IN ('INSERT', 'UPDATE', 'DELETE')),
    row_id          TEXT        NOT NULL,
    changed_by      TEXT,
    changed_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
    old_values      JSONB,
    new_values      JSONB,
    changed_columns TEXT[]
);

CREATE INDEX IF NOT EXISTS idx_audit_table_time
    ON audit.logged_actions (table_name, changed_at DESC);

CREATE INDEX IF NOT EXISTS idx_audit_changed_by
    ON audit.logged_actions (changed_by);
"""

# The trigger function must be created with CREATE OR REPLACE so it is
# safe to run this script multiple times.
CREATE_TRIGGER_FUNCTION = """
CREATE OR REPLACE FUNCTION audit.log_table_changes()
RETURNS TRIGGER
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = audit, public
AS $$
DECLARE
    v_row_id        TEXT;
    v_changed_cols  TEXT[];
    v_old_json      JSONB;
    v_new_json      JSONB;
    v_action        TEXT := TG_OP;
    v_changed_by    TEXT;
    v_pk_col        TEXT := TG_ARGV[0];
BEGIN
    -- Read the username from the session-level variable set by the app
    v_changed_by := NULLIF(current_setting('audit.username', true), '');

    -- Determine row identifier from the PK column
    IF TG_OP IN ('UPDATE', 'DELETE') THEN
        v_row_id := row_to_json(OLD)->>v_pk_col;
    ELSE
        v_row_id := row_to_json(NEW)->>v_pk_col;
    END IF;

    IF TG_OP = 'INSERT' THEN
        v_new_json     := row_to_json(NEW);
        v_changed_cols := ARRAY(SELECT jsonb_object_keys(v_new_json));

    ELSIF TG_OP = 'UPDATE' THEN
        v_old_json := row_to_json(OLD);
        v_new_json := row_to_json(NEW);

        SELECT ARRAY_AGG(key)
        INTO v_changed_cols
        FROM jsonb_object_keys(v_new_json) AS t(key)
        WHERE v_old_json->>key IS DISTINCT FROM v_new_json->>key;

        -- Skip audit entry if nothing actually changed
        IF v_changed_cols IS NULL OR array_length(v_changed_cols, 1) = 0 THEN
            RETURN NEW;
        END IF;

    ELSIF TG_OP = 'DELETE' THEN
        v_old_json     := row_to_json(OLD);
        v_changed_cols := ARRAY(SELECT jsonb_object_keys(v_old_json));
    END IF;

    INSERT INTO audit.logged_actions (
        schema_name, table_name, action, row_id,
        changed_by, old_values, new_values, changed_columns
    ) VALUES (
        TG_TABLE_SCHEMA, TG_TABLE_NAME, v_action, v_row_id,
        v_changed_by, v_old_json, v_new_json, v_changed_cols
    );

    RETURN COALESCE(NEW, OLD);
END;
$$;
"""

CREATE_TRIGGER_TEMPLATE = """
DROP TRIGGER IF EXISTS {trigger_name} ON public.{table_name};

CREATE TRIGGER {trigger_name}
    AFTER INSERT OR UPDATE OR DELETE
    ON public.{table_name}
    FOR EACH ROW
    EXECUTE FUNCTION audit.log_table_changes({pk_column});
"""

INSERT_ADMIN_USER = """
INSERT INTO public.users_information (username, password)
VALUES (%s, %s)
ON CONFLICT (username) DO NOTHING;
"""


# ── Main ───────────────────────────────────────────────────────────────

def run() -> None:
    print("Connecting to PostgreSQL …")
    connection = psycopg2.connect(DB_DSN)
    connection.autocommit = True  # DDL needs autocommit

    try:
        with connection.cursor() as cursor:

            # 1. Users table
            print("  Creating users_information table …")
            cursor.execute(CREATE_USERS_TABLE)

            # 2. Audit schema & table
            print("  Creating audit schema …")
            cursor.execute(CREATE_AUDIT_SCHEMA)
            cursor.execute(CREATE_AUDIT_TABLE)

            # 3. Trigger function
            print("  Creating trigger function audit.log_table_changes() …")
            cursor.execute(CREATE_TRIGGER_FUNCTION)

            # 4. Triggers on every target table
            for table_name, pk_column in TARGET_TABLES_PK.items():
                trigger_name = f"trg_audit_{table_name}"
                print(f"  Attaching trigger {trigger_name} → {table_name}")
                cursor.execute(
                    sql.SQL(CREATE_TRIGGER_TEMPLATE).format(
                        trigger_name=sql.Identifier(trigger_name),
                        table_name=sql.Identifier(table_name),
                        pk_column=sql.Literal(pk_column),
                    )
                )

            # 5. Seed default admin user
            print(f"  Seeding admin user '{DEFAULT_ADMIN_USER}' …")
            cursor.execute(
                INSERT_ADMIN_USER,
                (DEFAULT_ADMIN_USER, DEFAULT_ADMIN_PASS),
            )

        print("\n✅ Migration complete.")
        print(f"   Admin user   : {DEFAULT_ADMIN_USER}")
        print(f"   Audit tables : audit.logged_actions")
        print(f"   Triggers     : {len(TARGET_TABLES_PK)} tables")

    except Exception as exc:
        print(f"\n❌ Migration failed: {exc}")
        sys.exit(1)
    finally:
        connection.close()


if __name__ == "__main__":
    run()
