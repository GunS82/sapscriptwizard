import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Optional, Iterable, Tuple, Any


class RunHistory:
    """Simple wrapper around SQLite database for storing run history."""

    def __init__(self, db_path: Path | str = "run_history.db") -> None:
        self.db_path = Path(db_path)
        self.conn = sqlite3.connect(self.db_path)
        self._init_db()

    def _init_db(self) -> None:
        self.conn.execute(
            """
            CREATE TABLE IF NOT EXISTS runs(
                id INTEGER PRIMARY KEY,
                start_time TEXT,
                end_time TEXT,
                script TEXT,
                status TEXT,
                error_message TEXT,
                log_file TEXT,
                screenshot_path TEXT
            )
            """
        )
        self.conn.commit()

    def start_run(self, script: str) -> int:
        start_time = datetime.now().isoformat()
        cur = self.conn.cursor()
        cur.execute(
            """INSERT INTO runs(start_time, end_time, script, status, error_message, log_file, screenshot_path)
                VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (start_time, "", script, "running", "", "", ""),
        )
        self.conn.commit()
        return cur.lastrowid

    def finish_run(
        self,
        run_id: int,
        status: str,
        error_message: str = "",
        screenshot_path: str = "",
        log_file: str = "",
    ) -> None:
        end_time = datetime.now().isoformat()
        self.conn.execute(
            """UPDATE runs SET end_time=?, status=?, error_message=?, log_file=?, screenshot_path=? WHERE id=?""",
            (end_time, status, error_message, log_file, screenshot_path, run_id),
        )
        self.conn.commit()

    def fetch_runs(self, limit: Optional[int] = None) -> Iterable[Tuple[Any, ...]]:
        cur = self.conn.cursor()
        query = "SELECT * FROM runs ORDER BY id DESC"
        if limit:
            query += " LIMIT ?"
            cur.execute(query, (limit,))
        else:
            cur.execute(query)
        return cur.fetchall()

    def __del__(self) -> None:
        try:
            self.conn.close()
        except Exception:
            pass
