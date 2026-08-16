from __future__ import annotations

import json
import os
from typing import Any

import psycopg


def get_database_url() -> str:
    database_url = os.getenv("DATABASE_URL", "").strip()

    if not database_url:
        raise RuntimeError(
            "DATABASE_URL environment variable is not configured"
        )

    return database_url


def get_connection():
    return psycopg.connect(
        get_database_url(),
        sslmode="require",
    )


def initialize_database() -> None:
    """
    Create the history table when it does not exist yet.
    """

    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS history_state (
                    id INTEGER PRIMARY KEY,
                    data JSONB NOT NULL,
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                )
                """
            )

        connection.commit()


def read_history_db(
    base_headers: list[str],
) -> tuple[list[str], list[dict[str, str]]]:
    """
    Load the complete history from PostgreSQL.

    If no history exists yet, return an empty history.
    """

    initialize_database()

    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT data
                FROM history_state
                WHERE id = 1
                """
            )

            result = cursor.fetchone()

    if result is None:
        return list(base_headers), []

    data: dict[str, Any] = result[0]

    headers = data.get("headers", [])
    rows = data.get("rows", [])

    if not isinstance(headers, list):
        raise ValueError(
            "Database history contains invalid headers"
        )

    if not isinstance(rows, list):
        raise ValueError(
            "Database history contains invalid rows"
        )

    return headers, rows


def save_history_db(
    headers: list[str],
    rows: list[dict[str, str]],
) -> None:
    """
    Persist the complete history to PostgreSQL.
    """

    initialize_database()

    payload = {
        "headers": headers,
        "rows": rows,
    }

    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO history_state (
                    id,
                    data,
                    updated_at
                )
                VALUES (
                    1,
                    %s::jsonb,
                    NOW()
                )
                ON CONFLICT (id)
                DO UPDATE SET
                    data = EXCLUDED.data,
                    updated_at = NOW()
                """,
                (
                    json.dumps(
                        payload,
                        ensure_ascii=False,
                    ),
                ),
            )

        connection.commit()


def clear_history_db() -> None:
    """
    Delete the stored history.
    """

    initialize_database()

    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                DELETE FROM history_state
                WHERE id = 1
                """
            )

        connection.commit()