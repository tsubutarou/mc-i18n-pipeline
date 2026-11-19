# 単体パック（互換パック等、単体で完結したリソースパック/データパック）用ランチャー。
# modpack本体用の run.ps1 / config.json とは別の config.resourcepack.json を使う。
#
# 使い方:
#   .\run_rp.ps1               scan実行(既定)。single_packs_dir をスキャンしてxlsx生成
#   .\run_rp.ps1 -Step scan    同上
#   .\run_rp.ps1 -Step apply   翻訳済みxlsxを各パックのコピーに反映
#   .\run_rp.ps1 -Config other.json    別のconfigを使う
#
# 事前準備: setup.bat を実行するか、config.resourcepack.example.json を
#           config.resourcepack.json にコピーして single_packs_dir を書き換えておく。
#           単体パック(.zipまたは展開済みフォルダ)は single_packs_dir 直下に置く。

param(
  [ValidateSet("scan", "apply")]
  [string]$Step = "scan",
  [string]$Config = "config.resourcepack.json"
)

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

if (-not (Test-Path -LiteralPath $Config)) {
  Write-Host "[ERROR] $Config が見つかりません。setup.bat を実行するか、config.resourcepack.example.json をコピーして作成してください。" -ForegroundColor Red
  exit 1
}

if (Get-Command py -ErrorAction SilentlyContinue) { $pyExe = "py"; $pyPre = @("-3") }
elseif (Get-Command python -ErrorAction SilentlyContinue) { $pyExe = "python"; $pyPre = @() }
else {
  Write-Host "[ERROR] Python が見つかりません。https://www.python.org/ からインストールしてください。" -ForegroundColor Red
  exit 1
}

$internalStep = if ($Step -eq "scan") { "sp_scan" } else { "sp_apply" }

& $pyExe @($pyPre + @("-u", "run_all.py", "--config", $Config, "--step", $internalStep))
$code = $LASTEXITCODE

if ($code -ne 0) {
  Write-Host "[ERR] 失敗しました (exit code = $code)" -ForegroundColor Red
  exit $code
}

Write-Host "[OK] 完了" -ForegroundColor Green
