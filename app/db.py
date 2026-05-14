"""SQLite persistence (optional).

Two tables:
  papers(id, name, source_filename, source_blob, created_at)
  sets(id, paper_id, set_number, shuffle_questions, shuffle_options, created_at)

A "paper" is the originally-uploaded file (kept in source_blob so we can
re-parse it exactly). A "set" is just the (paper_id, set_number, mode) tuple
— since shuffles are deterministically seeded, we can regenerate the bytes
on demand without storing them.

If persistence is OFF in the UI, the server skips this module entirely and
keeps the in-memory upload only for the duration of a single request.
"""

from __future__ import annotations
import os
import sqlite3
import time
import uuid
from typing import List, Optional, Tuple


_SCHEMA = """
CREATE TABLE IF NOT EXISTS papers (
    id              TEXT PRIMARY KEY,
    name            TEXT NOT NULL,
    source_filename TEXT NOT NULL,
    source_blob     BLOB NOT NULL,
    created_at      INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS sets (
    id                  TEXT PRIMARY KEY,
    paper_id            TEXT NOT NULL,
    set_number          INTEGER NOT NULL,
    shuffle_questions   INTEGER NOT NULL,
    shuffle_options     INTEGER NOT NULL,
    created_at          INTEGER NOT NULL,
    FOREIGN KEY (paper_id) REFERENCES papers(id) ON DELETE CASCADE,
    UNIQUE (paper_id, set_number, shuffle_questions, shuffle_options)
);
"""


class Store:
    def __init__(self, path: str):
        self.path = path
        os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
        with self._conn() as c:
            c.executescript(_SCHEMA)

    def _conn(self):
        c = sqlite3.connect(self.path)
        c.row_factory = sqlite3.Row
        c.execute("PRAGMA foreign_keys = ON")
        return c

    # --- papers --------------------------------------------------------------

    def add_paper(self, name: str, filename: str, blob: bytes) -> str:
        pid = uuid.uuid4().hex
        with self._conn() as c:
            c.execute(
                "INSERT INTO papers (id, name, source_filename, source_blob, created_at) "
                "VALUES (?, ?, ?, ?, ?)",
                (pid, name, filename, blob, int(time.time())),
            )
        return pid

    def get_paper(self, paper_id: str) -> Optional[Tuple[str, str, bytes]]:
        with self._conn() as c:
            row = c.execute(
                "SELECT name, source_filename, source_blob FROM papers WHERE id = ?",
                (paper_id,),
            ).fetchone()
        if not row:
            return None
        return row["name"], row["source_filename"], bytes(row["source_blob"])

    def list_papers(self) -> List[dict]:
        with self._conn() as c:
            rows = c.execute(
                "SELECT id, name, source_filename, created_at "
                "FROM papers ORDER BY created_at DESC"
            ).fetchall()
        return [dict(r) for r in rows]

    def delete_paper(self, paper_id: str) -> bool:
        with self._conn() as c:
            cur = c.execute("DELETE FROM papers WHERE id = ?", (paper_id,))
        return cur.rowcount > 0

    # --- sets ----------------------------------------------------------------

    def record_set(self, paper_id: str, set_number: int,
                   shuffle_questions: bool, shuffle_options: bool) -> str:
        """Record that a set was generated. Idempotent on the UNIQUE key."""
        sid = uuid.uuid4().hex
        with self._conn() as c:
            try:
                c.execute(
                    "INSERT INTO sets "
                    "(id, paper_id, set_number, shuffle_questions, shuffle_options, created_at) "
                    "VALUES (?, ?, ?, ?, ?, ?)",
                    (sid, paper_id, int(set_number),
                     int(bool(shuffle_questions)), int(bool(shuffle_options)),
                     int(time.time())),
                )
            except sqlite3.IntegrityError:
                # Already recorded — return the existing row's id.
                row = c.execute(
                    "SELECT id FROM sets WHERE paper_id = ? AND set_number = ? "
                    "AND shuffle_questions = ? AND shuffle_options = ?",
                    (paper_id, int(set_number),
                     int(bool(shuffle_questions)), int(bool(shuffle_options))),
                ).fetchone()
                return row["id"] if row else sid
        return sid

    def list_sets(self, paper_id: str) -> List[dict]:
        with self._conn() as c:
            rows = c.execute(
                "SELECT id, set_number, shuffle_questions, shuffle_options, created_at "
                "FROM sets WHERE paper_id = ? ORDER BY set_number ASC",
                (paper_id,),
            ).fetchall()
        return [dict(r) for r in rows]
