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
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS artifacts (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id TEXT NOT NULL,
                ts         TEXT NOT NULL,
                filename   TEXT NOT NULL,
                ext        TEXT NOT NULL
            )
            """
        )
        c.execute("CREATE INDEX IF NOT EXISTS idx_art ON artifacts (student_id, ext)")
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS history (
                student_id TEXT PRIMARY KEY,
                updated_at TEXT NOT NULL,
                payload    TEXT NOT NULL
            )
            """
        )


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


def spent_period_usd(student_id: str) -> float:
    """課題期間の通算（このDBが作られてから全部）。"""
    with _conn() as c:
        row = c.execute(
            "SELECT COALESCE(SUM(usd), 0) AS s FROM usage WHERE student_id = ?",
            (student_id,),
        ).fetchone()
    return float(row["s"])


def spent_period_total_usd() -> float:
    """全学生の通算。日次上限が毎日リセットされても、ここで最終的に止まる。"""
    with _conn() as c:
        row = c.execute("SELECT COALESCE(SUM(usd), 0) AS s FROM usage").fetchone()
    return float(row["s"])


# --------------------------------------------------------------------------
# 生成物の記録
#
# 種類ごとの作成回数を数えるために使う。スライド(pptx)は1回あたり約156円と
# 突出して高いため、回数制限をかける根拠になる。
# --------------------------------------------------------------------------

def record_artifact(student_id: str, filename: str) -> None:
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    with _conn() as c:
        c.execute(
            "INSERT INTO artifacts (student_id, ts, filename, ext) VALUES (?,?,?,?)",
            (student_id, datetime.datetime.now().isoformat(timespec="seconds"), filename, ext),
        )


def count_artifacts(student_id: str, ext: str) -> int:
    with _conn() as c:
        row = c.execute(
            "SELECT COUNT(*) n FROM artifacts WHERE student_id = ? AND ext = ?",
            (student_id, ext),
        ).fetchone()
    return int(row["n"])


# --------------------------------------------------------------------------
# 会話履歴の永続化
#
# 課題として数日〜数週間使われるため、プロセス内メモリだけだと
# 再デプロイや再起動で学生の作業が消える。永続ディスク上のDBに退避する。
# --------------------------------------------------------------------------

def save_history(student_id: str, payload: str) -> None:
    with _conn() as c:
        c.execute(
            "INSERT INTO history (student_id, updated_at, payload) VALUES (?,?,?)"
            " ON CONFLICT(student_id) DO UPDATE SET updated_at=excluded.updated_at,"
            " payload=excluded.payload",
            (student_id, datetime.datetime.now().isoformat(timespec="seconds"), payload),
        )


def load_history(student_id: str) -> str | None:
    with _conn() as c:
        row = c.execute(
            "SELECT payload FROM history WHERE student_id = ?", (student_id,)
        ).fetchone()
    return row["payload"] if row else None


def clear_history(student_id: str) -> None:
    with _conn() as c:
        c.execute("DELETE FROM history WHERE student_id = ?", (student_id,))


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
