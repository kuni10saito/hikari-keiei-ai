# 株式会社ヒカリ 経営演習AI

名古屋学院大学 経営学部の演習用。学生がブラウザから Claude を使い、
株式会社ヒカリ（実在企業・公開情報ベース＋演習用想定財務）の経営分析を行い、
Excel・PowerPoint・Word・PDF の成果物を生成させる。

**学生に Claude のアカウント契約は不要。** API 課金はすべて教員のキー1本に集約される。

---

## 構成

```
company-ai/
├── company.md        ← CLAUDE.md 相当。会社設定（system プロンプト）
├── app.py            ← FastAPI サーバ
├── db.py             ← 使用量ログ（SQLite）
├── roster.txt        ← 受講者の学籍番号
├── static/index.html ← 画面
├── outputs/          ← 生成物（学籍番号ごと・自動作成）
└── usage.sqlite3     ← 使用量DB（自動作成）
```

## セットアップ

`anthropic` / `fastapi` / `uvicorn` はこの環境に導入済み。追加が必要な場合のみ:

```powershell
pip install -r requirements.txt
```

## 起動

```powershell
cd C:\Users\saito\Downloads\company-ai

$env:ANTHROPIC_API_KEY = "sk-ant-..."     # 必須
$env:CLASS_PASSWORD    = "任意のパスワード"   # 学生に配る（既定: hikari）
$env:DAILY_YEN_CAP     = "500"             # 学生1人あたりの1日上限（円）
$env:ADMIN_KEY         = "任意の管理キー"     # 使用量閲覧用

python -m uvicorn app:app --host 0.0.0.0 --port 8000
```

> **`uvicorn` 単体では動きません。** 必ず `python -m uvicorn` と書くこと。
> このPCではパッケージが `--user` 側（`AppData\Roaming\Python\Python314\Scripts\`）に
> 入っており、そのフォルダが PATH に無いため `uvicorn.exe` が見つからない。
> PATH 上の Python は `C:\Python314\Scripts` で、そちらには exe が置かれていない。

ブラウザで `http://localhost:8000/` を開く。
学生の端末から使わせるなら、同一 LAN で `http://<このPCのIP>:8000/`。

学外からアクセスさせる場合は、このまま公開せず Tailscale か
リバースプロキシ（HTTPS 終端）を挟むこと。

## インターネット公開（Cloudflare Tunnel）

学外からアクセスさせる場合。ルータの設定やグローバルIPは不要で、HTTPS になる。

### 初回だけ

```powershell
winget install --id Cloudflare.cloudflared
```

インストール後、**PowerShell を開き直す**（PATH を読み直すため）。

### 毎回

```powershell
cd C:\Users\saito\Downloads\company-ai
$env:ANTHROPIC_API_KEY  = "sk-ant-..."
$env:CLASS_PASSWORD     = "十分に長い合言葉"   # 8文字以上。既定値のまま公開しないこと
$env:DAILY_TOTAL_YEN_CAP = "8000"

.\start_public.ps1
```

サーバ用のウィンドウが別に開き、このウィンドウに配布URLが表示される。

```
==============================================================
  学生に配布するURL

    https://xxxx-yyyy-zzzz.trycloudflare.com

  パスワード : ...
==============================================================
```

**終了は `Ctrl+C`。** トンネルが閉じ、外部からは即座に到達不能になる。
サーバ側のウィンドウも閉じること。

### 注意

| 項目 | 内容 |
|---|---|
| URL は毎回変わる | 無料の quick tunnel のため。授業ごとに配り直す。固定したいなら Cloudflare アカウントで named tunnel を作る |
| PC を閉じると止まる | このPCがサーバ本体。スリープさせないこと |
| 公開中は誰でもURLに到達できる | 防御はクラス共通パスワードと名簿。パスワードは口頭で伝え、SNS等に貼らせない |

### 公開時に効いている防御

- **名簿方式** — `roster.txt` にない学籍番号は入れない。被害額の上限が構造的に決まる
- **総当たり対策** — 同一IPから5回失敗で15分ロック（正しいパスワードでも入れない）
- **Secure cookie** — HTTPS 経由のときだけ自動で付与され、平文接続に載らない
- **二重の支出上限** — 学生ごと `DAILY_YEN_CAP` に加え、クラス全体 `DAILY_TOTAL_YEN_CAP`

## 教員用API

いずれも `?key=<ADMIN_KEY>` が必要。`<>` は付けずに値だけを書く。

| URL | 用途 |
|---|---|
| `/api/admin/status?key=...` | 設定と現状の一覧。**`disk_ok` が最重要** |
| `/api/admin/usage?key=...` | 学生ごとの累計と費用の内訳 |
| `/api/admin/reset?key=...&student_id=28b0113` | 指定した学生の使用量・履歴・生成物記録を消す |

### `disk_ok`

`false` なら永続ディスクが効いておらず、再デプロイのたびに使用量が
リセットされて利用上限が機能しない。Render の Settings → Disks で
`/var/data` を追加すること。

### 費用の内訳

`usage` は `yen_breakdown` を返す。**`キャッシュ書込` が9割を超えることがある**
（実測: 2ターン392円のうち362.6円）。ここが見えないと高額の原因を追えない。

### リセット

Render 上では DB を直接触れないため API から消す。誤操作防止のため
`student_id` は必須で、全消しは用意していない。
**課題開始前に、動作確認で消費した分を必ずリセットすること。**

---

## 設計上の要点

### 1. company.md が CLAUDE.md にあたる

`app.py` が `company.md` を丸ごと読み、`system` パラメータに載せる。
会社を差し替えるなら `company_b.md` を作って
`$env:COMPANY_FILE = "company_b.md"` を指定する。

### 2. プロンプトキャッシュ

`system` の最後のブロックに `cache_control` を置いてある（TTL 1時間）。
全学生が同じ `company.md` を読むので、2人目以降はこの部分が**約1/10の料金**になる。

**触ってはいけない点**：`system` に日時・学生名・可変情報を入れると
キャッシュが毎回無効化されて効果が消える。可変情報は必ず `messages` 側に置くこと。

### 3. Agent Skills

`container.skills` で `xlsx` / `docx` / `pptx` / `pdf` を有効化し、
`code_execution` ツールと組み合わせている。beta ヘッダは3つ必要：

- `code-execution-2025-08-25`
- `skills-2025-10-02`
- `files-api-2025-04-14`

### 4. 生成ファイルの持ち越し

コンテナはリクエストごとに新規のため、そのままでは
「さっきの Excel にグラフを足して」が通らない。
直近3件の `file_id` を `container_upload` ブロックで次ターンに持ち込んでいる。

### 5. pause_turn

code execution は実行が長引くと `stop_reason: "pause_turn"` で返る。
追加の user メッセージを足さずに再送する処理を `_run_turn()` に入れてある（上限5回）。

### 6. 使用量上限（日次＋通算の二段）

| 環境変数 | 既定 | 対象 |
|---|---:|---|
| `DAILY_YEN_CAP` | 300円 | 学生1人 / 1日 |
| `DAILY_TOTAL_YEN_CAP` | 600円 | クラス全体 / 1日 |
| `PERIOD_YEN_CAP` | 500円 | 学生1人 / 課題期間の通算 |
| `PERIOD_TOTAL_YEN_CAP` | 1,200円 | クラス全体 / 課題期間の通算 |
| `PPTX_LIMIT` | 1回 | 学生1人あたりのスライド作成回数 |

### 実測単価（claude-sonnet-5 / effort=medium）

| 依頼 | 実測 |
|---|---:|
| 分析・議論 | 2〜5円 |
| Excel 生成 | 約18円 |
| **スライド生成** | **約156円** |

スライドが突出して高い。費用の9割は Anthropic 側サーバ内部の code execution
ループによるキャッシュ書込で、こちらからは制御できない（`task_budget` も
`MAX_RESUME` も届かない）。そのため**回数で制限する**。

上限に達した学生には `pptx` スキル自体を渡さないので、依頼を受け付けてから
断るのではなく、費用が発生する前に封じられる。

**日次だけでは歯止めにならない。** 課題が4週間なら日次上限が28回リセットされるため、
最終的に止めるのは通算の方。通算は `usage.sqlite3` の全レコードを集計するので、
永続ディスクが効いていることが前提になる。

キャッシュ読出は 0.1 倍、書込は 1.25 倍で計算に入れてある。

---

## 既知の制約

| 項目 | 内容 |
|---|---|
| 会話履歴 | 永続ディスクの SQLite に保存。再起動・再デプロイをまたいで再開できる |
| 認証 | 学籍番号＋クラス共通パスワード。演習用の割り切りで、本番相当の認証ではない |
| 同時実行 | 単一プロセス想定。多人数が同時に重い依頼を出すと待たされる |
| 為替 | `USD_JPY` を固定値で持っている（既定 155）。円換算は概算 |
| モデル | `claude-sonnet-5`。変更するなら `app.py` の `MODEL` と、`PRICE_INTRO` / `PRICE_STANDARD` を**必ずセットで**直す（単価がずれると利用上限が意図どおり効かなくなる） |

---

## ⚠️ 財務データの扱い

株式会社ヒカリは**実在企業**で、財務諸表は公開されていない。
`company.md` の財務・市場・課題は**すべて演習用に作成した架空の数値**であり、
公開情報の節（会社概要・沿革・商品・理念）とは明確に分離してある。

AI 側にも「財務に言及するときは想定値だと明示する」「成果物に注記を入れる」を
指示済み。学生に配布する際も口頭で伝えること。
