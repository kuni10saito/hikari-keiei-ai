"""株式会社ヒカリ 経営演習AI — Web アプリ本体

company.md を system プロンプトとして読み込み、Agent Skills（xlsx/docx/pptx/pdf）と
code execution を有効にした Claude を、学生がブラウザから使えるようにする。

  学生 → ブラウザ → FastAPI → Claude API（あなたの APIキー1本）
                        ↓
                  outputs/<学籍番号>/*.xlsx などを保存してダウンロード提供

起動:
    set ANTHROPIC_API_KEY=sk-ant-...
    uvicorn app:app --host 0.0.0.0 --port 8000
"""
from __future__ import annotations

import datetime
import json
import os
import pathlib
import re
import secrets
import time
import unicodedata
from typing import Any

import anthropic
from fastapi import Body, Cookie, FastAPI, HTTPException, Request, Response
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

import db

# --------------------------------------------------------------------------
# 設定
# --------------------------------------------------------------------------

BASE = pathlib.Path(__file__).parent
COMPANY_FILE = os.environ.get("COMPANY_FILE", "company.md")
COMPANY_MD = (BASE / COMPANY_FILE).read_text(encoding="utf-8")

# ダッシュボード表示用の数値。company.md の【演習用・想定】と一致させること。
COMPANY_DATA: dict[str, Any] = json.loads(
    (BASE / "company_data.json").read_text(encoding="utf-8")
)

# 書き込むデータの置き場所。ローカルではこのフォルダ、Render では
# 永続ディスクのマウント先（DATA_DIR=/var/data）を指す。
# ここを間違えると再デプロイのたびに usage.sqlite3 が消え、利用上限が
# リセットされて課金の歯止めが外れるので、最も重要な設定。
DATA_DIR = pathlib.Path(os.environ.get("DATA_DIR", str(BASE)))
DATA_DIR.mkdir(parents=True, exist_ok=True)

OUTPUT_ROOT = DATA_DIR / "outputs"
OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

MODEL = "claude-sonnet-5"
SKILL_IDS = ["xlsx", "docx", "pptx", "pdf"]
BETAS = ["code-execution-2025-08-25", "skills-2025-10-02", "files-api-2025-04-14"]

CLASS_PASSWORD = os.environ.get("CLASS_PASSWORD", "hikari")
DAILY_YEN_CAP = float(os.environ.get("DAILY_YEN_CAP", "500"))        # 学生1人あたり
DAILY_TOTAL_YEN_CAP = float(os.environ.get("DAILY_TOTAL_YEN_CAP", "8000"))  # 全体
USD_JPY = float(os.environ.get("USD_JPY", "155"))
MAX_CARRY_FILES = 3          # 直前ターンの成果物を何件まで次のターンに持ち越すか
MAX_RESUME = 5               # pause_turn の再開上限

# ログイン総当たり対策。インターネットに公開するとクラス共通パスワードが
# 総当たりの標的になるため、IP単位で失敗回数を数えて締め出す。
LOGIN_MAX_FAILS = 5
LOGIN_WINDOW_SEC = 600       # この秒数のあいだの失敗を数える
LOGIN_LOCK_SEC = 900         # 上限に達したら締め出す秒数

# claude-sonnet-5 の料金（$/1Mトークン）。
# 2026-08-31 までは導入価格 $2/$10、それ以降は通常価格 $3/$15。
# 利用上限の判定に使う値なので、期限が切れたら自動で通常価格に戻るようにしてある
# （安く見積もったままだと上限が実質的に緩んでしまう）。
SONNET5_INTRO_UNTIL = datetime.date(2026, 8, 31)
PRICE_INTRO = (2.00, 10.00)
PRICE_STANDARD = (3.00, 15.00)


def _prices() -> tuple[float, float]:
    """($/token) の (入力, 出力) を返す。呼び出しのたびに日付を見る。"""
    inp, out = PRICE_INTRO if datetime.date.today() <= SONNET5_INTRO_UNTIL else PRICE_STANDARD
    return inp / 1_000_000, out / 1_000_000

# code execution はモデルの生成時間に加えてコンテナ実行時間がかかる。
# SDK 既定の 10 分では足りないことがあるので明示的に伸ばす（単位は秒）。
client = anthropic.Anthropic().with_options(timeout=900.0)

def _norm(value: str) -> str:
    """全角→半角に正規化し、前後の空白を落として小文字化する。

    日本語IMEで学籍番号を打つと全角（２８ｂ０１１３）になりやすく、
    見た目がほぼ同じなのに文字列としては一致しない。ここで吸収する。
    """
    return unicodedata.normalize("NFKC", value or "").strip().casefold()


# パスワードは大文字小文字を区別する（casefold しない）が、全角と前後空白は吸収する。
_CLASS_PASSWORD_NORM = unicodedata.normalize("NFKC", CLASS_PASSWORD).strip()

def _load_roster() -> set[str]:
    """受講者名簿を読む。

    学籍番号と氏名は個人情報なので、リポジトリには入れない。
    優先順:
      1. 環境変数 ROSTER（カンマ/空白/改行区切り）— Render ではこれを使う
      2. roster.txt — ローカル実行用。.gitignore 済み
    """
    env = os.environ.get("ROSTER", "").strip()
    if env:
        return {_norm(x) for x in re.split(r"[,\s]+", env) if _norm(x)}

    path = BASE / "roster.txt"
    if not path.exists():
        return set()
    return {
        _norm(line.split("#", 1)[0])
        for line in path.read_text(encoding="utf-8").splitlines()
        if line.split("#", 1)[0].strip()
    }


ROSTER: set[str] = _load_roster()
if not ROSTER:
    raise SystemExit(
        "受講者名簿が空です。環境変数 ROSTER に学籍番号を設定するか、"
        "roster.txt を用意してください（例: ROSTER=\"28b0113,29a0001\"）。"
    )

# プロセス内状態。学生12名程度の演習を想定した割り切り。
SESSIONS: dict[str, str] = {}            # cookie トークン -> 学籍番号
HISTORY: dict[str, list[Any]] = {}       # 学籍番号 -> messages
CARRY: dict[str, list[str]] = {}         # 学籍番号 -> 直近に生成された file_id
LOGIN_FAILS: dict[str, list[float]] = {} # IP -> 失敗時刻
LOGIN_LOCK: dict[str, float] = {}        # IP -> ロック解除時刻

app = FastAPI(title="ヒカリ経営演習AI")
db.init()


@app.on_event("startup")
def _banner() -> None:
    """起動時に設定状況を出す。ログインで詰まったときの一次切り分け用。

    print はパイプ越しだとバッファされて表示されないことがあるので flush する。
    """
    from_env = "CLASS_PASSWORD" in os.environ
    lines = [
        "=" * 58,
        "  ヒカリ経営演習AI  起動しました",
        f"  会社設定    : {COMPANY_FILE}（{len(COMPANY_MD):,} 文字）",
        f"  名簿        : {len(ROSTER)}名 -> {sorted(ROSTER)}",
        f"  パスワード  : {'環境変数 CLASS_PASSWORD を使用' if from_env else '未設定のため既定値 hikari'}"
        f"（{len(_CLASS_PASSWORD_NORM)}文字）",
        f"  モデル      : {MODEL}"
        f"（入力 ${_prices()[0] * 1_000_000:.2f} / 出力 ${_prices()[1] * 1_000_000:.2f} per 1M）",
        f"  1日上限     : {DAILY_YEN_CAP:.0f} 円/人 ／ クラス全体 {DAILY_TOTAL_YEN_CAP:.0f} 円",
        f"  ログイン防御: {LOGIN_MAX_FAILS}回失敗で{LOGIN_LOCK_SEC // 60}分ロック",
        f"  APIキー     : {'設定済み' if os.environ.get('ANTHROPIC_API_KEY') else '★未設定★'}",
        f"  データ保存先: {DATA_DIR}"
        f"{'（DATA_DIR 指定）' if os.environ.get('DATA_DIR') else '（既定：アプリと同じ場所）'}",
        f"  使用量DB    : {db.DB_PATH.name} "
        f"{'既存' if db.DB_PATH.exists() else '新規作成'}",
        "=" * 58,
    ]
    print("\n".join(lines), flush=True)


# --------------------------------------------------------------------------
# ヘルパー
# --------------------------------------------------------------------------

def _cost_usd(*, input_tokens: int, output_tokens: int,
              cache_write: int, cache_read: int) -> float:
    """1ターンの概算コスト。キャッシュ書込は 1.25 倍、読出は 0.1 倍で課金される。"""
    price_in, price_out = _prices()
    return (
        input_tokens * price_in
        + cache_write * price_in * 1.25
        + cache_read * price_in * 0.10
        + output_tokens * price_out
    )


def _safe_name(filename: str) -> str | None:
    """パストラバーサル対策。ディレクトリ成分を落として妥当性を確認する。"""
    name = os.path.basename(filename or "")
    if not name or name in (".", "..") or "/" in name or "\\" in name:
        return None
    return name


def _require_student(session: str | None) -> str:
    student_id = SESSIONS.get(session or "")
    if not student_id:
        raise HTTPException(status_code=401, detail="ログインしてください")
    return student_id


def _client_ip(request: Request) -> str:
    """Cloudflare Tunnel 経由でも本来の接続元が分かるようにする。"""
    for header in ("cf-connecting-ip", "x-forwarded-for"):
        value = request.headers.get(header)
        if value:
            return value.split(",")[0].strip()
    return request.client.host if request.client else "unknown"


def _is_https(request: Request) -> bool:
    proto = request.headers.get("x-forwarded-proto", "")
    return proto.lower() == "https" or request.url.scheme == "https"


def _login_guard(ip: str) -> None:
    """ロック中なら弾く。"""
    until = LOGIN_LOCK.get(ip, 0.0)
    if until > time.time():
        raise HTTPException(
            status_code=429,
            detail=f"試行回数が多すぎます。{int(until - time.time()) // 60 + 1}分後にやり直してください",
        )


def _login_failed(ip: str) -> None:
    now = time.time()
    fails = [t for t in LOGIN_FAILS.get(ip, []) if now - t < LOGIN_WINDOW_SEC]
    fails.append(now)
    LOGIN_FAILS[ip] = fails
    if len(fails) >= LOGIN_MAX_FAILS:
        LOGIN_LOCK[ip] = now + LOGIN_LOCK_SEC
        LOGIN_FAILS[ip] = []
        print(f"[SECURITY] {ip} をログイン失敗{LOGIN_MAX_FAILS}回で"
              f"{LOGIN_LOCK_SEC // 60}分ロックしました", flush=True)


def _request_kwargs(messages: list[Any]) -> dict[str, Any]:
    return dict(
        model=MODEL,
        max_tokens=16000,
        betas=BETAS,
        container={
            "skills": [
                {"type": "anthropic", "skill_id": sid, "version": "latest"}
                for sid in SKILL_IDS
            ]
        },
        tools=[{"type": "code_execution_20260521", "name": "code_execution"}],
        # cache_control は system の最後のブロックに置く。tools + system がまとめて
        # キャッシュされ、全学生が同じ company.md を共有するので読出コストは約1/10。
        # ここに日時や学生名を差し込むとキャッシュが毎回無効化されるので絶対に入れない。
        system=[
            {
                "type": "text",
                "text": COMPANY_MD,
                "cache_control": {"type": "ephemeral", "ttl": "1h"},
            }
        ],
        messages=messages,
    )


def _harvest_files(response: Any, student_id: str) -> list[dict[str, str]]:
    """code execution が生成したファイルを outputs/<学籍番号>/ に保存する。"""
    out_dir = OUTPUT_ROOT / student_id
    out_dir.mkdir(parents=True, exist_ok=True)

    saved: list[dict[str, str]] = []
    for block in response.content:
        if getattr(block, "type", None) != "bash_code_execution_tool_result":
            continue
        result = getattr(block, "content", None)
        if getattr(result, "type", None) != "bash_code_execution_result":
            continue
        for ref in getattr(result, "content", None) or []:
            if getattr(ref, "type", None) != "bash_code_execution_output":
                continue
            file_id = ref.file_id
            meta = client.beta.files.retrieve_metadata(file_id)
            name = _safe_name(meta.filename)
            if name is None:
                continue
            client.beta.files.download(file_id).write_to_file(out_dir / name)
            saved.append({"name": name, "file_id": file_id})
    return saved


def _run_turn(messages: list[Any]) -> tuple[Any, dict[str, int]]:
    """1ターン実行。pause_turn は上限つきで自動再開する。

    code execution はサーバ側ツールなので、実行が長引くと stop_reason=pause_turn で
    いったん返ってくる。追加の user メッセージは足さず、そのまま再送する。
    """
    totals = {"input": 0, "output": 0, "cache_write": 0, "cache_read": 0}
    response = None

    for _ in range(MAX_RESUME):
        response = client.beta.messages.create(**_request_kwargs(messages))
        u = response.usage
        totals["input"] += u.input_tokens or 0
        totals["output"] += u.output_tokens or 0
        totals["cache_write"] += getattr(u, "cache_creation_input_tokens", 0) or 0
        totals["cache_read"] += getattr(u, "cache_read_input_tokens", 0) or 0

        messages.append({"role": "assistant", "content": response.content})
        if response.stop_reason != "pause_turn":
            break

    return response, totals


# --------------------------------------------------------------------------
# API
# --------------------------------------------------------------------------

@app.post("/api/login")
def login(request: Request, response: Response, body: dict = Body(...)):
    ip = _client_ip(request)
    _login_guard(ip)

    student_id = _norm(str(body.get("student_id", "")))
    password = unicodedata.normalize("NFKC", str(body.get("password", ""))).strip()

    # 教室内ツールなので、原因が分かるようにエラーを分けている。
    if not student_id:
        raise HTTPException(status_code=401, detail="学籍番号が入力されていません")
    if student_id not in ROSTER:
        _login_failed(ip)
        raise HTTPException(
            status_code=401,
            detail=f"学籍番号「{student_id}」は名簿(roster.txt)にありません",
        )
    if not secrets.compare_digest(password, _CLASS_PASSWORD_NORM):
        _login_failed(ip)
        raise HTTPException(status_code=401, detail="パスワードが違います")

    LOGIN_FAILS.pop(ip, None)
    token = secrets.token_urlsafe(32)
    SESSIONS[token] = student_id
    HISTORY.setdefault(student_id, [])
    response.set_cookie(
        "session", token, httponly=True, samesite="lax",
        secure=_is_https(request),   # HTTPS 経由なら平文接続に載せない
        max_age=60 * 60 * 8,
    )
    return {"student_id": student_id, "spent_yen": db.spent_today_usd(student_id) * USD_JPY,
            "cap_yen": DAILY_YEN_CAP}


@app.post("/api/logout")
def logout(response: Response, session: str | None = Cookie(default=None)):
    SESSIONS.pop(session or "", None)
    response.delete_cookie("session")
    return {"ok": True}


@app.post("/api/reset")
def reset(session: str | None = Cookie(default=None)):
    """会話をリセット（使用量ログは残る）。"""
    student_id = _require_student(session)
    HISTORY[student_id] = []
    CARRY[student_id] = []
    return {"ok": True}


@app.post("/api/chat")
def chat(body: dict = Body(...), session: str | None = Cookie(default=None)):
    student_id = _require_student(session)
    text = str(body.get("message", "")).strip()
    if not text:
        raise HTTPException(status_code=400, detail="メッセージが空です")

    spent_yen = db.spent_today_usd(student_id) * USD_JPY
    if spent_yen >= DAILY_YEN_CAP:
        raise HTTPException(
            status_code=429,
            detail=f"本日の利用上限（{DAILY_YEN_CAP:.0f}円）に達しました。明日また使ってください。",
        )

    # サーキットブレーカー。想定外の使われ方をしたときに請求が青天井にならないよう、
    # クラス全体の合計にも上限を置く。
    total_yen = db.spent_today_total_usd() * USD_JPY
    if total_yen >= DAILY_TOTAL_YEN_CAP:
        raise HTTPException(
            status_code=429,
            detail=f"クラス全体の本日の上限（{DAILY_TOTAL_YEN_CAP:.0f}円）に達しました。教員に連絡してください。",
        )

    messages = HISTORY.setdefault(student_id, [])

    # 直前に作った成果物をコンテナに持ち込む。これがないと
    # 「さっきの Excel にグラフを足して」が通らない（コンテナは毎回新規のため）。
    blocks: list[dict[str, Any]] = [{"type": "text", "text": text}]
    for file_id in CARRY.get(student_id, [])[-MAX_CARRY_FILES:]:
        blocks.append({"type": "container_upload", "file_id": file_id})
    messages.append({"role": "user", "content": blocks})

    try:
        response, totals = _run_turn(messages)
    except anthropic.APIStatusError as exc:
        messages.pop()  # 失敗したターンは履歴から戻す
        raise HTTPException(status_code=502, detail=f"Claude API エラー: {exc.message}") from exc
    except anthropic.APIConnectionError as exc:
        messages.pop()
        raise HTTPException(status_code=502, detail="Claude API に接続できませんでした") from exc

    usd = _cost_usd(
        input_tokens=totals["input"],
        output_tokens=totals["output"],
        cache_write=totals["cache_write"],
        cache_read=totals["cache_read"],
    )
    db.record(
        student_id,
        input_tokens=totals["input"],
        output_tokens=totals["output"],
        cache_write=totals["cache_write"],
        cache_read=totals["cache_read"],
        usd=usd,
        prompt=text,
    )

    # refusal は HTTP 200 で返る。content を読む前に必ず確認する。
    if response.stop_reason == "refusal":
        return {
            "reply": "この依頼には応答できませんでした。表現を変えて試してください。",
            "files": [],
            "spent_yen": db.spent_today_usd(student_id) * USD_JPY,
            "cap_yen": DAILY_YEN_CAP,
        }

    reply = "".join(
        b.text for b in response.content if getattr(b, "type", None) == "text"
    ).strip()

    files = _harvest_files(response, student_id)
    if files:
        CARRY.setdefault(student_id, []).extend(f["file_id"] for f in files)

    if response.stop_reason == "max_tokens":
        reply += "\n\n（出力が上限に達しました。「続き」と入力すると再開します）"

    return {
        "reply": reply or "（テキスト応答なし）",
        "files": [{"name": f["name"], "url": f"/api/files/{f['name']}"} for f in files],
        "spent_yen": db.spent_today_usd(student_id) * USD_JPY,
        "cap_yen": DAILY_YEN_CAP,
    }


@app.get("/api/files/{filename}")
def download(filename: str, session: str | None = Cookie(default=None)):
    student_id = _require_student(session)
    name = _safe_name(filename)
    if name is None:
        raise HTTPException(status_code=400, detail="不正なファイル名です")

    path = (OUTPUT_ROOT / student_id / name).resolve()
    root = (OUTPUT_ROOT / student_id).resolve()
    if not path.is_file() or root not in path.parents:
        raise HTTPException(status_code=404, detail="ファイルが見つかりません")

    return FileResponse(path, filename=name)


@app.get("/api/company")
def company(session: str | None = Cookie(default=None)):
    """ダッシュボードが表示する会社データ。"""
    _require_student(session)
    return COMPANY_DATA


@app.get("/api/usage")
def usage(session: str | None = Cookie(default=None)):
    student_id = _require_student(session)
    return {
        "student_id": student_id,
        "spent_yen": db.spent_today_usd(student_id) * USD_JPY,
        "cap_yen": DAILY_YEN_CAP,
    }


@app.get("/api/admin/usage")
def admin_usage(key: str = ""):
    """教員用。ADMIN_KEY を環境変数で設定し、?key=... を付けて開く。"""
    admin_key = os.environ.get("ADMIN_KEY", "")
    if not admin_key or not secrets.compare_digest(key, admin_key):
        raise HTTPException(status_code=403, detail="forbidden")
    rows = db.summary()
    for r in rows:
        r["yen"] = round((r["usd"] or 0) * USD_JPY, 1)
    return rows


app.mount("/", StaticFiles(directory=BASE / "static", html=True), name="static")
