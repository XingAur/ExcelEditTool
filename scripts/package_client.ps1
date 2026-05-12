$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $scriptDir

Push-Location $root
try {
  python -m pip install -r requirements.txt

  Push-Location (Join-Path $root "frontend")
  try {
    npm install
    npm run build
  }
  finally {
    Pop-Location
  }

  python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --name ExcelEditTool `
    --icon "packaging/excel_editor_icon.ico" `
    --add-data "frontend/dist;frontend/dist" `
    --collect-submodules webview `
    desktop/main.py

  Copy-Item `
    -LiteralPath (Join-Path $root "packaging/client-readme.txt") `
    -Destination (Join-Path $root "dist/ExcelEditTool/README.txt") `
    -Force

  $releaseDir = Join-Path $root "release"
  New-Item -ItemType Directory -Force -Path $releaseDir | Out-Null
  Compress-Archive `
    -Path (Join-Path $root "dist/ExcelEditTool/*") `
    -DestinationPath (Join-Path $releaseDir "ExcelEditTool.zip") `
    -Force
}
finally {
  Pop-Location
}
