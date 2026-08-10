# ヒカリ経営演習AI — インターネット公開用の起動スクリプト
#
#   .\start_public.ps1
#
# uvicorn を別ウィンドウで起動し、このウィンドウで Cloudflare Tunnel を張る。
# 学生に配るURLはこのウィンドウに大きく表示される。
# 終了するときは Ctrl+C（トンネルが閉じ、外部からは即座に到達不能になる）。
#
# 環境変数（ANTHROPIC_API_KEY など）は子プロセスに自動で継承されるので、
# このウィンドウで設定しておけば別ウィンドウ側にも引き継がれる。

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

# ---- 必須設定の確認 ----------------------------------------------------
if (-not $env:ANTHROPIC_API_KEY) {
    Write-Host "ANTHROPIC_API_KEY が未設定です。先に実行してください:" -ForegroundColor Red
    Write-Host '  $env:ANTHROPIC_API_KEY = "sk-ant-..."'
    exit 1
}
if (-not $env:CLASS_PASSWORD) {
    Write-Host "CLASS_PASSWORD が未設定です。公開するなら必ず設定してください:" -ForegroundColor Red
    Write-Host '  $env:CLASS_PASSWORD = "十分に長い合言葉"'
    Write-Host "  (既定値 hikari のまま公開しないこと)"
    exit 1
}
if ($env:CLASS_PASSWORD.Length -lt 8) {
    Write-Host "CLASS_PASSWORD が短すぎます (8文字以上にしてください)" -ForegroundColor Red
    exit 1
}
# cloudflared は winget で入れても PATH に載らないことがある
# (実体は C:\Program Files (x86)\cloudflared\)。既知の場所も探す。
$cloudflared = (Get-Command cloudflared -ErrorAction SilentlyContinue).Source
if (-not $cloudflared) {
    $candidates = @(
        "$env:ProgramFiles\cloudflared\cloudflared.exe",
        "${env:ProgramFiles(x86)}\cloudflared\cloudflared.exe",
        "$env:LOCALAPPDATA\Microsoft\WinGet\Links\cloudflared.exe"
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) { $cloudflared = $c; break }
    }
}
if (-not $cloudflared) {
    Write-Host "cloudflared が見つかりません。先にインストールしてください:" -ForegroundColor Red
    Write-Host "  winget install --id Cloudflare.cloudflared"
    exit 1
}
Write-Host "cloudflared: $cloudflared" -ForegroundColor DarkGray

# ---- アプリ本体を別ウィンドウで起動 ------------------------------------
# 外部公開はトンネル経由に限定するため、uvicorn は 127.0.0.1 だけで待ち受ける。
Write-Host "サーバを起動しています..." -ForegroundColor Cyan

$serverCmd = "Set-Location '$PSScriptRoot'; python -m uvicorn app:app --host 127.0.0.1 --port 8000"
$server = Start-Process powershell -PassThru -ArgumentList "-NoExit", "-Command", $serverCmd

$ready = $false
foreach ($i in 1..40) {
    Start-Sleep -Milliseconds 500
    try {
        Invoke-WebRequest "http://127.0.0.1:8000/" -UseBasicParsing -TimeoutSec 2 | Out-Null
        $ready = $true
        break
    } catch { }
}

if (-not $ready) {
    Write-Host "サーバが起動しませんでした。別ウィンドウのエラーを確認してください。" -ForegroundColor Red
    exit 1
}

Write-Host "サーバ起動 OK (PID $($server.Id))" -ForegroundColor Green
Write-Host ""
Write-Host "Cloudflare Tunnel を張っています。URL が出るまで数秒かかります..." -ForegroundColor Cyan
Write-Host ""

# cloudflared は起動ログを標準エラーに出す。
# PowerShell 5.1 で native exe に 2>&1 を使うと、正常なログ行まで
# NativeCommandError として扱われて落ちるので、ファイルに逃がして読む。
$outLog = Join-Path $env:TEMP "hikari_cloudflared_out.log"
$errLog = Join-Path $env:TEMP "hikari_cloudflared_err.log"
Remove-Item $outLog, $errLog -ErrorAction SilentlyContinue

$tunnel = Start-Process -FilePath $cloudflared -PassThru -NoNewWindow `
    -ArgumentList @("tunnel", "--url", "http://127.0.0.1:8000") `
    -RedirectStandardOutput $outLog -RedirectStandardError $errLog

$bar = "=" * 62
$url = $null

try {
    foreach ($i in 1..80) {
        Start-Sleep -Milliseconds 500
        $text = ""
        foreach ($f in @($outLog, $errLog)) {
            if (Test-Path $f) {
                $text += (Get-Content $f -Raw -ErrorAction SilentlyContinue)
            }
        }
        if ($text -match 'https://[a-z0-9-]+\.trycloudflare\.com') {
            $url = $Matches[0]
            break
        }
        if ($tunnel.HasExited) { break }
    }

    if (-not $url) {
        Write-Host "URL を取得できませんでした。cloudflared のログ:" -ForegroundColor Red
        foreach ($f in @($outLog, $errLog)) {
            if (Test-Path $f) { Get-Content $f | Select-Object -Last 25 }
        }
        return
    }

    Write-Host $bar -ForegroundColor Green
    Write-Host "  学生に配布するURL" -ForegroundColor Green
    Write-Host ""
    Write-Host "    $url" -ForegroundColor White -BackgroundColor DarkGreen
    Write-Host ""
    Write-Host "  パスワード : $env:CLASS_PASSWORD"
    Write-Host $bar -ForegroundColor Green
    Write-Host ""
    Write-Host "公開中です。終了するには Ctrl+C を押してください。" -ForegroundColor Cyan
    Write-Host "(サーバとトンネルの両方を自動で停止します)" -ForegroundColor DarkGray

    while (-not $tunnel.HasExited) { Start-Sleep -Seconds 1 }
}
finally {
    Write-Host ""
    Write-Host "停止しています..." -ForegroundColor Yellow
    foreach ($p in @($tunnel, $server)) {
        if ($p -and -not $p.HasExited) {
            Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
        }
    }
    Write-Host "トンネルとサーバを停止しました。外部からは到達できません。" -ForegroundColor Green
}
