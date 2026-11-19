# modpack本体（複数mod）用ランチャー
#
# 使い方:
#   .\run.ps1                 全ステップ実行 (mods_jar_dir設定時は 展開 -> xlsx生成 -> 衝突チェック -> パック生成)
#   .\run.ps1 -Step extract   jarの展開だけ(mods_jar_dirの.jarをmods_rootに展開)
#   .\run.ps1 -Step xlsx      xlsxの生成だけ
#   .\run.ps1 -Step check     翻訳済みxlsxの衝突チェックだけ
#   .\run.ps1 -Step pack      リソースパック生成だけ
#   .\run.ps1 -Step diagnose  xlsxの行ズレ・欠落modidの診断
#   .\run.ps1 -Config other_config.json    別のconfigを使う
#   .\run.ps1 -Step extract -Force         既に展開済みのjarも再展開する
#
# 事前準備: setup.bat を実行するか、config.example.json を config.json にコピーして
#           mods_jar_dir (.jarが入っているフォルダ) などを書き換えておく。

param(
  [ValidateSet("all", "extract", "xlsx", "check", "pack", "diagnose", "sp_scan", "sp_apply")]
  [string]$Step = "all",
  [string]$Config = "config.json",
  [switch]$Force
)

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

if (-not (Test-Path -LiteralPath $Config)) {
  Write-Host "[ERROR] $Config が見つかりません。setup.bat を実行するか、config.example.json をコピーして作成してください。" -ForegroundColor Red
  exit 1
}

if (Get-Command py -ErrorAction SilentlyContinue) { $pyExe = "py"; $pyPre = @("-3") }
elseif (Get-Command python -ErrorAction SilentlyContinue) { $pyExe = "python"; $pyPre = @() }
else {
  Write-Host "[ERROR] Python が見つかりません。https://www.python.org/ からインストールしてください。" -ForegroundColor Red
  exit 1
}

$runArgs = $pyPre + @("-u", "run_all.py", "--config", $Config, "--step", $Step)
if ($Force) { $runArgs += "--force" }

& $pyExe @runArgs
$code = $LASTEXITCODE

if ($code -ne 0) {
  Write-Host "[ERR] 失敗しました (exit code = $code)" -ForegroundColor Red
  exit $code
}

Write-Host "[OK] 完了" -ForegroundColor Green
