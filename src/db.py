"""Read-only SQLite access for the synthetic retail analysis dataset."""

from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path
import sqlite3
from typing import Iterator


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DB_PATH = PROJECT_ROOT / "data" / "app.db"


def open_readonly_connection(path: str | Path = DEFAULT_DB_PATH) -> sqlite3.Connection:
    """Open an existing database in read-only mode; never create an empty DB."""
    db_path = Path(path).expanduser().resolve()
    if not db_path.is_file():
        raise FileNotFoundError(f"Synthetic analysis database not found: {db_path}")
    uri = f"file:{db_path.as_posix()}?mode=ro"
    connection = sqlite3.connect(uri, uri=True, timeout=5.0)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    connection.execute("PRAGMA query_only = ON")
    return connection


@contextmanager
def read_connection(path: str | Path = DEFAULT_DB_PATH) -> Iterator[sqlite3.Connection]:
    """Yield a short-lived read-only connection and always close it."""
    connection = open_readonly_connection(path)
    try:
        yield connection
    finally:
        connection.close()
