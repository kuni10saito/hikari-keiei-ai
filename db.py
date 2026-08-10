"""使用量ログ（SQLite）。

会話履歴はプロセス内メモリに持ち、ここには課金・監査に必要な情報だけを残す。
サーバを再起動すると会話はリセットされるが、使用量ログは残る。

保存先は DATA_DIR（未設定ならこのフォルダ）。Render では永続ディスクの
マウント先を指す。ここが揮発すると利用上限が毎回リセットされ、
API課金の歯止めが効かなくなる。
"""
from __future__ import annotations

import datetime
import os
import pathlib
import sqlite3

DATA_DIR = pathlib.Path(os.environ.get("DATA_DIR", str(pathlib.Path(__file__).parent)))
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = DATA_DIR / "usage.sqlite3"


def _conn() -> sqlite3.Connection:
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c


def init() -> None:
    with _conn() as c:
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS usage (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id     TEXT    NOT NULL,
                day            TEXT    NOT NULL,
                ts             TEXT    NOT NULL,
                input_tokens   INTEGER NOT NULL DEFAULT 0,
                output_tokens  INTEGER NOT NULL DEFAULT 0,
                cache_write    INTEGER NOT NULL DEFAULT 0,
                cache_read     INTEGER NOT NULL DEFAULT 0,
                usd            REAL    NOT NULL DEFAULT 0,
                prompt         TEXT
            )
            """
        )
        c.execute("CREATE INDEX IF NOT EXISTS idx_usage_day ON usage (student_id, day)")


def today() -> str:
    return datetime.date.today().isoformat()


def record(
    student_id: str,
    *,
    input_tokens: int,
    output_tokens: int,
    cache_write: int,
    cache_read: int,
    usd: float,
    prompt: str,
) -> None:
    with _conn() as c:
        c.execute(
            "INSERT INTO usage (student_id, day, ts, input_tokens, output_tokens,"
            " cache_write, cache_read, usd, prompt) VALUES (?,?,?,?,?,?,?,?,?)",
            (
                student_id,
                today(),
                datetime.datetime.now().isoformat(timespec="seconds"),
                input_tokens,
                output_tokens,
                cache_write,
                cache_read,
                usd,
                prompt[:500],
            ),
        )


def spent_today_usd(student_id: str) -> float:
    with _conn() as c:
        row = c.execute(
            "SELECT COALESCE(SUM(usd), 0) AS s FROM usage WHERE student_id = ? AND day = ?",
            (student_id, today()),
        ).fetchone()
    return float(row["s"])


def spent_today_total_usd() -> float:
    """全学生の合計。サーキットブレーカー用。"""
    with _conn() as c:
        row = c.execute(
            "SELECT COALESCE(SUM(usd), 0) AS s FROM usage WHERE day = ?", (today(),)
        ).fetchone()
    return float(row["s"])


def summary() -> list[dict]:
    """管理用：学生ごとの累計。"""
    with _conn() as c:
        rows = c.execute(
            "SELECT student_id,"
            "       COUNT(*)                AS turns,"
            "       SUM(input_tokens)       AS input_tokens,"
            "       SUM(output_tokens)      AS output_tokens,"
            "       SUM(cache_read)         AS cache_read,"
            "       SUM(usd)                AS usd"
            " FROM usage GROUP BY student_id ORDER BY usd DESC"
        ).fetchall()
    return [dict(r) for r in rows]
