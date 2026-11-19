# 初回セットアップ: 必要なPythonパッケージを入れて、configのひな形をコピーする。
$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

if (Get-Command py -ErrorAction SilentlyContinue) { $pyExe = "py"; $pyPre = @("-3") }
elseif (Get-Command python -ErrorAction SilentlyContinue) { $pyExe = "python"; $pyPre = @() }
else {
  Write-Host "[ERROR] Python が見つかりません。https://www.python.org/ からインストールしてください。" -ForegroundColor Red
  exit 1
}

Write-Host "Pythonパッケージをインストールします..."
& $pyExe @($pyPre + @("-m", "pip", "install", "-r", "requirements.txt"))
if ($LASTEXITCODE -ne 0) {
  Write-Host "[ERR] pip install に失敗しました" -ForegroundColor Red
  exit $LASTEXITCODE
}

foreach ($pair in @(@("config.example.json", "config.json"), @("config.resourcepack.example.json", "config.resourcepack.json"))) {
  if (Test-Path -LiteralPath $pair[1]) {
    Write-Host "$($pair[1]) は既にあるのでそのままにします"
  } else {
    Copy-Item -LiteralPath $pair[0] -Destination $pair[1]
    Write-Host "$($pair[1]) を作成しました。中のフォルダパスを自分の環境に書き換えてください。"
  }
}

Write-Host "[OK] セットアップ完了" -ForegroundColor Green
