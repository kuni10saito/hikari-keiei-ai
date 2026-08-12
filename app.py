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

# ダッシュボード表示用の数値。上の .md の【演習用・想定】と一致させること。
# 会社を差し替えるときは COMPANY_FILE と対で切り替わる必要がある。
# 既定値を .md のファイル名から導出するので、
#   COMPANY_FILE=company_b.md → company_b_data.json
# が自動で選ばれ、環境変数を2つ設定する手間が要らない。
_default_data = COMPANY_FILE.rsplit(".", 1)[0] + "_data.json"
COMPANY_DATA_FILE = os.environ.get("COMPANY_DATA_FILE", _default_data)
_data_path = BASE / COMPANY_DATA_FILE
if not _data_path.exists():          # 既存の company_data.json への後方互換
    _data_path = BASE / "company_data.json"
COMPANY_DATA_FILE = _data_path.name
COMPANY_DATA: dict[str, Any] = json.loads(_data_path.read_text(encoding="utf-8"))

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
BETAS = [
    "code-execution-2025-08-25",
    "skills-2025-10-02",
    "files-api-2025-04-14",
    "context-management-2025-06-27",   # 古いツール結果を履歴から落とす
    "task-budgets-2026-03-13",         # 1ターンの規模に上限を置く
]

CLASS_PASSWORD = os.environ.get("CLASS_PASSWORD", "hikari")

# 利用上限は日次と通算の二段。日次だけだと課題期間の日数ぶん上限が
# リセットされ、通算では歯止めにならない。最終的に止めるのは通算の方。
DAILY_YEN_CAP = float(os.environ.get("DAILY_YEN_CAP", "300"))               # 1人/日
DAILY_TOTAL_YEN_CAP = float(os.environ.get("DAILY_TOTAL_YEN_CAP", "600"))   # 全体/日
PERIOD_YEN_CAP = float(os.environ.get("PERIOD_YEN_CAP", "500"))             # 1人/通算
PERIOD_TOTAL_YEN_CAP = float(os.environ.get("PERIOD_TOTAL_YEN_CAP", "1200"))# 全体/通算

# スライド(pptx)は1回あたり約156円で、Excel(18円)の約9倍。
# 費用の9割は Anthropic 側サーバ内部の code execution ループによる
# キャッシュ書込で、こちらからは制御できない。回数で制限する。
PPTX_LIMIT = int(os.environ.get("PPTX_LIMIT", "1"))

USD_JPY = float(os.environ.get("USD_JPY", "155"))
MAX_CARRY_FILES = 3          # 直前ターンの成果物を何件まで次のターンに持ち越すか
MAX_RESUME = 5               # pause_turn の再開上限

# 思考の深さ。思考トークンも出力として課金されるため、費用に直結する。
# 実測では1ターンの費用の56%が出力だった。Sonnet 5 の既定は high。
# medium は Sonnet 4.6 の high 相当で、経営分析の演習には十分。
# 成果物の質が落ちるようなら EFFORT=high に戻す。
EFFORT = os.environ.get("EFFORT", "medium")

# 1ターンでモデルが使える token 数の目安。モデルは残量を見ながらペース配分し、
# 尽きる前に切り上げる（max_tokens と違い、モデル自身が認識する予算）。
# 実測: 未設定だとスライド3枚の生成が cache_write 61万トークン・262円に達した。
# 最小値は 20,000。
TASK_BUDGET_TOKENS = int(os.environ.get("TASK_BUDGET_TOKENS", "50000"))

# ターン開始前に必要な余力（円）。上限判定はターン開始前にしか行えないため、
# 残額ぎりぎりで重い依頼を始めると大幅に超過する。あらかじめ余力を要求する。
TURN_RESERVE_YEN = float(os.environ.get("TURN_RESERVE_YEN", "170"))   # pptx がまだ使えるとき
LIGHT_RESERVE_YEN = float(os.environ.get("LIGHT_RESERVE_YEN", "30"))  # pptx 使い切り後

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
        f"  会社設定    : {COMPANY_FILE}（{len(COMPANY_MD):,} 文字）"
        f" ／ {COMPANY_DATA_FILE}",
        f"  名簿        : {len(ROSTER)}名 -> {sorted(ROSTER)}",
        f"  パスワード  : {'環境変数 CLASS_PASSWORD を使用' if from_env else '未設定のため既定値 hikari'}"
        f"（{len(_CLASS_PASSWORD_NORM)}文字）",
        f"  モデル      : {MODEL} / effort={EFFORT} / task_budget={TASK_BUDGET_TOKENS:,}"
        f"（入力 ${_prices()[0] * 1_000_000:.2f} / 出力 ${_prices()[1] * 1_000_000:.2f} per 1M）",
        f"  日次上限    : {DAILY_YEN_CAP:.0f} 円/人 ／ 全体 {DAILY_TOTAL_YEN_CAP:.0f} 円",
        f"  通算上限    : {PERIOD_YEN_CAP:.0f} 円/人 ／ 全体 {PERIOD_TOTAL_YEN_CAP:.0f} 円"
        f"（現在 {db.spent_period_total_usd() * USD_JPY:.1f} 円）",
        f"  スライド    : 学生1人あたり {PPTX_LIMIT} 回まで（1回 約156円）",
        f"  必要余力    : {TURN_RESERVE_YEN:.0f} 円（pptx可）／ {LIGHT_RESERVE_YEN:.0f} 円（pptx使用済）",
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


def _usage_payload(student_id: str) -> dict[str, Any]:
    """画面に返す利用状況。日次と通算の両方を出す。"""
    return {
        "student_id": student_id,
        "spent_yen": db.spent_today_usd(student_id) * USD_JPY,
        "cap_yen": DAILY_YEN_CAP,
        "period_yen": db.spent_period_usd(student_id) * USD_JPY,
        "period_cap_yen": PERIOD_YEN_CAP,
        "pptx_used": db.count_artifacts(student_id, "pptx"),
        "pptx_limit": PPTX_LIMIT,
    }


def _check_caps(student_id: str, heavy: bool = True) -> None:
    """上限を確認する。

    上限判定はターン開始前にしか行えず、1ターンの途中では止められない。
    そのため残額ぎりぎりで重い依頼を始めると超過する。

    必要な余力は「そのターンで起きうる最大費用」で決める。
      heavy=True  … pptx がまだ使える。実測 約156円のターンがありうる
      heavy=False … pptx は使い切り。最大でも Excel 生成の 約18円
    使い切った学生に156円の余力を要求すると、使えるはずの予算を無駄にする。
    """
    reserve = TURN_RESERVE_YEN if heavy else LIGHT_RESERVE_YEN
    checks = (
        (db.spent_today_usd(student_id) * USD_JPY, DAILY_YEN_CAP, "本日のあなた"),
        (db.spent_today_total_usd() * USD_JPY, DAILY_TOTAL_YEN_CAP, "本日のクラス全体"),
        (db.spent_period_usd(student_id) * USD_JPY, PERIOD_YEN_CAP, "課題期間のあなた"),
        (db.spent_period_total_usd() * USD_JPY, PERIOD_TOTAL_YEN_CAP, "課題期間のクラス全体"),
    )
    for spent, cap, label in checks:
        if spent >= cap:
            raise HTTPException(
                status_code=429,
                detail=f"{label}の利用上限（{cap:.0f}円）に達しました。教員に連絡してください。",
            )
        if spent + reserve > cap:
            raise HTTPException(
                status_code=429,
                detail=f"{label}の残額が {cap - spent:.0f} 円です。"
                       f"1回の依頼に {reserve:.0f} 円の余力が必要なため、ここで停止しました。",
            )


# --------------------------------------------------------------------------
# 会話履歴の永続化
# --------------------------------------------------------------------------

def _block_type(block: Any) -> str:
    return block.get("type", "") if isinstance(block, dict) else getattr(block, "type", "")


def _prune(messages: list[Any]) -> list[Any]:
    """履歴からテキスト以外のブロックを落とす。

    code execution の実行ログとスクリプトは1回のExcel/スライド生成で
    数十万文字になり、それが以後のターンで毎回再送されて費用が跳ね上がる。
    （実測: 1回のスライド生成で履歴が 59,439 → 536,373 文字、そのターンが262円）

    context editing はサーバ側でモデルに見せる文脈を削るだけで、こちらが
    送信する履歴は減らない。送信量そのものを抑えるにはここで落とす必要がある。

    残すのはテキストのみ。会話の流れは保たれ、生成済みファイルは
    CARRY の file_id で次ターンに引き継がれる。
    """
    pruned: list[Any] = []
    for m in messages:
        content = m["content"]
        if isinstance(content, str):
            pruned.append({"role": m["role"], "content": content})
            continue
        blocks = [b for b in content if _block_type(b) == "text"]
        if blocks:                      # 中身が全部ツールだった回は丸ごと落とす
            pruned.append({"role": m["role"], "content": blocks})
    return pruned


def _jsonable(messages: list[Any]) -> str:
    """SDK の応答ブロック（Pydantic）を含む messages を JSON 文字列にする。"""
    out: list[dict[str, Any]] = []
    for m in messages:
        content = m["content"]
        if isinstance(content, str):
            out.append({"role": m["role"], "content": content})
            continue
        blocks: list[Any] = []
        for b in content:
            if isinstance(b, dict):
                blocks.append(b)
            elif hasattr(b, "model_dump"):
                blocks.append(b.model_dump(mode="json", exclude_none=True))
            else:
                blocks.append(b)
        out.append({"role": m["role"], "content": blocks})
    return json.dumps(out, ensure_ascii=False)


def _persist_history(student_id: str) -> None:
    """ターン成功後に履歴を刈り込み、メモリとディスクの両方を更新する。

    メモリ側も置き換えるのが要点。ここを保存用のコピーだけに留めると、
    次のターンで送信されるのは肥大化したままの履歴になる。
    """
    try:
        pruned = _prune(HISTORY.get(student_id, []))
        HISTORY[student_id] = pruned
        db.save_history(student_id, _jsonable(pruned))
    except Exception as exc:  # 保存に失敗しても会話自体は続行させる
        print(f"[WARN] 履歴の保存に失敗 ({student_id}): {exc}", flush=True)


def _restore_history(student_id: str) -> list[Any]:
    """ディスクから履歴を復元する。壊れていれば黙って捨てて新規開始。"""
    if student_id in HISTORY:
        return HISTORY[student_id]
    messages: list[Any] = []
    try:
        payload = db.load_history(student_id)
        if payload:
            loaded = json.loads(payload)
            if isinstance(loaded, list):
                messages = loaded
    except Exception as exc:
        print(f"[WARN] 履歴の復元に失敗 ({student_id})、新規開始します: {exc}", flush=True)
        db.clear_history(student_id)
    HISTORY[student_id] = messages
    return messages


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


def _skills_for(student_id: str) -> list[str]:
    """この学生に渡すスキル。上限に達した pptx は外す。

    スキルを渡さなければモデルはスライドを作れず「作成できない」と答える。
    依頼を受け付けてから拒否するのではなく、費用が発生する前に封じる。
    """
    if PPTX_LIMIT <= 0 or db.count_artifacts(student_id, "pptx") >= PPTX_LIMIT:
        return [s for s in SKILL_IDS if s != "pptx"]
    return SKILL_IDS


def _request_kwargs(messages: list[Any], skills: list[str] | None = None) -> dict[str, Any]:
    skills = SKILL_IDS if skills is None else skills
    return dict(
        model=MODEL,
        max_tokens=16000,
        betas=BETAS,
        output_config={
            "effort": EFFORT,
            "task_budget": {"type": "tokens", "total": TASK_BUDGET_TOKENS},
        },
        container={
            "skills": [
                {"type": "anthropic", "skill_id": sid, "version": "latest"}
                for sid in skills
            ]
        },
        tools=[{"type": "code_execution_20260521", "name": "code_execution"}],
        # Excel等を1回作るたびに、code execution の巨大な実行ログとスクリプトが
        # 履歴に積まれ、以後のターンで毎回再送されて費用が雪だるま式に増える。
        # 古いツール結果と入力を履歴から自動で落とす。
        # （実測: 未対策だと会話が30万トークンに達し、短い一言に265円かかった）
        context_management={
            "edits": [{"type": "clear_tool_uses_20250919", "clear_tool_inputs": True}]
        },
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
            db.record_artifact(student_id, name)
            saved.append({"name": name, "file_id": file_id})
    return saved


def _run_turn(messages: list[Any], skills: list[str]) -> tuple[Any, dict[str, int]]:
    """1ターン実行。pause_turn は上限つきで自動再開する。

    code execution はサーバ側ツールなので、実行が長引くと stop_reason=pause_turn で
    いったん返ってくる。追加の user メッセージは足さず、そのまま再送する。
    """
    totals = {"input": 0, "output": 0, "cache_write": 0, "cache_read": 0}
    response = None

    for attempt in range(1, MAX_RESUME + 1):
        response = client.beta.messages.create(**_request_kwargs(messages, skills))
        u = response.usage
        cw = getattr(u, "cache_creation_input_tokens", 0) or 0
        cr = getattr(u, "cache_read_input_tokens", 0) or 0
        totals["input"] += u.input_tokens or 0
        totals["output"] += u.output_tokens or 0
        totals["cache_write"] += cw
        totals["cache_read"] += cr

        # ターン内部の各回を可視化する。pause_turn の再送が繰り返されると
        # そのたびに肥大化した文脈を送り直してキャッシュを書き、費用が跳ねる。
        yen = _cost_usd(input_tokens=u.input_tokens or 0, output_tokens=u.output_tokens or 0,
                        cache_write=cw, cache_read=cr) * USD_JPY
        print(f"  [turn] {attempt}/{MAX_RESUME} stop={response.stop_reason} "
              f"in={u.input_tokens or 0:,} out={u.output_tokens or 0:,} "
              f"cache_w={cw:,} cache_r={cr:,} → {yen:.1f}円", flush=True)

        messages.append({"role": "assistant", "content": response.content})
        if response.stop_reason != "pause_turn":
            break
    else:
        print(f"  [turn] pause_turn が {MAX_RESUME} 回続いたため打ち切りました。"
              f"依頼が重すぎる可能性があります。", flush=True)

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
    _restore_history(student_id)   # 前回の続きから再開できるようにする
    response.set_cookie(
        "session", token, httponly=True, samesite="lax",
        secure=_is_https(request),   # HTTPS 経由なら平文接続に載せない
        max_age=60 * 60 * 8,
    )
    return _usage_payload(student_id)


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
    db.clear_history(student_id)
    return {"ok": True}


@app.post("/api/chat")
def chat(body: dict = Body(...), session: str | None = Cookie(default=None)):
    student_id = _require_student(session)
    text = str(body.get("message", "")).strip()
    if not text:
        raise HTTPException(status_code=400, detail="メッセージが空です")

    # pptx を使い切っていれば、そのスキル自体を渡さない。
    # 残り回数によって、そのターンで起きうる最大費用が変わる。
    skills = _skills_for(student_id)
    _check_caps(student_id, heavy="pptx" in skills)

    messages = _restore_history(student_id)

    # 直前に作った成果物をコンテナに持ち込む。これがないと
    # 「さっきの Excel にグラフを足して」が通らない（コンテナは毎回新規のため）。
    blocks: list[dict[str, Any]] = [{"type": "text", "text": text}]
    for file_id in CARRY.get(student_id, [])[-MAX_CARRY_FILES:]:
        blocks.append({"type": "container_upload", "file_id": file_id})
    messages.append({"role": "user", "content": blocks})

    try:
        response, totals = _run_turn(messages, skills)
    except anthropic.BadRequestError as exc:
        # ディスクから復元した履歴が API に受け付けられないケース。
        # 学生を詰まらせないよう、履歴を捨てて今回の発言だけで一度やり直す。
        print(f"[WARN] 復元履歴が拒否されたため会話をリセット ({student_id}): {exc}", flush=True)
        db.clear_history(student_id)
        CARRY[student_id] = []
        messages = HISTORY[student_id] = [{"role": "user", "content": [{"type": "text", "text": text}]}]
        try:
            response, totals = _run_turn(messages, skills)
        except anthropic.APIStatusError as exc2:
            HISTORY[student_id] = []
            raise HTTPException(
                status_code=502, detail=f"Claude API エラー: {exc2.message}"
            ) from exc2
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

    # 会話が膨らむと1ターンの費用が跳ね上がる。異常に気づけるようログに出す。
    context_tokens = totals["input"] + totals["cache_write"] + totals["cache_read"]
    if context_tokens > 200_000:
        print(f"[WARN] {student_id} の会話が肥大化しています "
              f"({context_tokens:,} トークン / このターン {usd * USD_JPY:.1f} 円)。"
              f" 「会話をリセット」を促してください。", flush=True)

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

    # ターンが成功したここで初めて永続化する。失敗したターンは残さない。
    _persist_history(student_id)

    return {
        "reply": reply or "（テキスト応答なし）",
        "files": [{"name": f["name"], "url": f"/api/files/{f['name']}"} for f in files],
        "spent_yen": db.spent_today_usd(student_id) * USD_JPY,
        "cap_yen": DAILY_YEN_CAP,
        "period_yen": db.spent_period_usd(student_id) * USD_JPY,
        "period_cap_yen": PERIOD_YEN_CAP,
        "pptx_used": db.count_artifacts(student_id, "pptx"),
        "pptx_limit": PPTX_LIMIT,
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
    return _usage_payload(_require_student(session))


def _require_admin(key: str) -> None:
    admin_key = os.environ.get("ADMIN_KEY", "")
    if not admin_key or not secrets.compare_digest(key, admin_key):
        raise HTTPException(status_code=403, detail="forbidden")


@app.get("/api/admin/usage")
def admin_usage(key: str = ""):
    """教員用。ADMIN_KEY を環境変数で設定し、?key=... を付けて開く。"""
    _require_admin(key)
    pi, po = _prices()
    rows = db.summary()
    for r in rows:
        r["yen"] = round((r["usd"] or 0) * USD_JPY, 1)
        # 費用の内訳。cache_write が9割を超えることがあり、そこが見えないと
        # 高額になった原因を管理画面から追えない。
        r["yen_breakdown"] = {
            "入力": round((r["input_tokens"] or 0) * pi * USD_JPY, 1),
            "キャッシュ書込": round((r["cache_write"] or 0) * pi * 1.25 * USD_JPY, 1),
            "キャッシュ読出": round((r["cache_read"] or 0) * pi * 0.10 * USD_JPY, 1),
            "出力": round((r["output_tokens"] or 0) * po * USD_JPY, 1),
        }
    return rows


@app.get("/api/admin/turns")
def admin_turns(key: str = "", limit: int = 50):
    """1ターンずつの明細。どの依頼が高かったのかを特定する。

    合計だけでは「スライドが高い」までしか分からず、
    実際にどの発言で何円かかったのかを追えない。
    """
    _require_admin(key)
    pi, po = _prices()
    out = []
    for r in db.recent_turns(min(max(limit, 1), 200)):
        out.append({
            "ts": r["ts"],
            "student_id": r["student_id"],
            "yen": round((r["usd"] or 0) * USD_JPY, 1),
            "prompt": r["prompt"],
            "tokens": {
                "input": r["input_tokens"], "output": r["output_tokens"],
                "cache_write": r["cache_write"], "cache_read": r["cache_read"],
            },
            "yen_breakdown": {
                "入力": round((r["input_tokens"] or 0) * pi * USD_JPY, 1),
                "キャッシュ書込": round((r["cache_write"] or 0) * pi * 1.25 * USD_JPY, 1),
                "キャッシュ読出": round((r["cache_read"] or 0) * pi * 0.10 * USD_JPY, 1),
                "出力": round((r["output_tokens"] or 0) * po * USD_JPY, 1),
            },
        })
    return out


@app.get("/api/admin/status")
def admin_status(key: str = ""):
    """設定と現状を1画面で確認する。

    最重要は disk_ok。False なら永続ディスクが効いておらず、
    再デプロイのたびに使用量がリセットされて上限が無意味になる。
    """
    _require_admin(key)
    disk_ok = os.environ.get("DATA_DIR") is not None and str(DATA_DIR) != str(BASE)
    return {
        # どの会社が読み込まれているか。COMPANY_FILE を変えたあとの確認用。
        "company": {
            "name": COMPANY_DATA.get("company", {}).get("name", ""),
            "md_file": COMPANY_FILE,
            "data_file": COMPANY_DATA_FILE,
            "periods": COMPANY_DATA.get("periods", []),
        },
        "disk_ok": disk_ok,
        "data_dir": str(DATA_DIR),
        "db_exists": db.DB_PATH.exists(),
        "warning": None if disk_ok else
                   "永続ディスクが効いていません。再デプロイのたびに使用量が"
                   "リセットされ、利用上限が機能しません。Render の "
                   "Settings → Disks で /var/data を追加してください。",
        "model": MODEL,
        "effort": EFFORT,
        "roster": sorted(ROSTER),
        "pptx_limit": PPTX_LIMIT,
        "caps_yen": {
            "daily_per_student": DAILY_YEN_CAP,
            "daily_total": DAILY_TOTAL_YEN_CAP,
            "period_per_student": PERIOD_YEN_CAP,
            "period_total": PERIOD_TOTAL_YEN_CAP,
        },
        "spent_yen": {
            "today_total": round(db.spent_today_total_usd() * USD_JPY, 1),
            "period_total": round(db.spent_period_total_usd() * USD_JPY, 1),
            "per_student": {
                sid: {
                    "today": round(db.spent_today_usd(sid) * USD_JPY, 1),
                    "period": round(db.spent_period_usd(sid) * USD_JPY, 1),
                    "pptx_used": db.count_artifacts(sid, "pptx"),
                }
                for sid in sorted(ROSTER)
            },
        },
    }


@app.post("/api/admin/reset")
@app.get("/api/admin/reset")
def admin_reset(key: str = "", student_id: str = ""):
    """指定した学生の使用量・会話履歴・生成物記録を消す。

    Render 上では DB を直接触れないため、ここから消せるようにしてある。
    誤操作を避けるため対象の学籍番号を必須にし、全消しは用意しない。
    """
    _require_admin(key)
    target = _norm(student_id)
    if not target:
        raise HTTPException(
            status_code=400,
            detail="student_id を指定してください"
                   "（例: /api/admin/reset?key=...&student_id=28b0113）",
        )
    if target not in ROSTER:
        raise HTTPException(status_code=404, detail=f"「{target}」は名簿にありません")

    deleted = db.reset_student(target)
    HISTORY.pop(target, None)
    CARRY.pop(target, None)
    print(f"[ADMIN] {target} をリセット: {deleted}", flush=True)
    return {
        "student_id": target,
        "deleted": deleted,
        "spent_yen_now": round(db.spent_period_usd(target) * USD_JPY, 1),
    }


app.mount("/", StaticFiles(directory=BASE / "static", html=True), name="static")
