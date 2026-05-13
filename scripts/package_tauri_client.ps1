$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $scriptDir
$frontendDir = Join-Path $root "frontend"
$bundleDir = Join-Path $frontendDir "src-tauri\target\release\bundle\nsis"
$releaseDir = Join-Path $root "release"
$releaseSetupPath = Join-Path $releaseDir "ExcelEditTool_Tauri_Setup.exe"

function Invoke-Native {
  param(
    [Parameter(Mandatory = $true)]
    [string]$FilePath,
    [Parameter(Mandatory = $true)]
    [string[]]$ArgumentList
  )

  & $FilePath @ArgumentList
  if ($LASTEXITCODE -ne 0) {
    throw "$FilePath failed with exit code $LASTEXITCODE."
  }
}

Push-Location $frontendDir
try {
  Invoke-Native -FilePath "npm" -ArgumentList @("install")
  Invoke-Native -FilePath "npm" -ArgumentList @("run", "tauri:build")
}
finally {
  Pop-Location
}

$setup = Get-ChildItem `
  -LiteralPath $bundleDir `
  -Filter "ExcelEditTool_*_x64-setup.exe" `
  -File |
  Sort-Object LastWriteTime -Descending |
  Select-Object -First 1

if (-not $setup) {
  throw "Tauri NSIS setup package was not found in $bundleDir."
}

New-Item -ItemType Directory -Force -Path $releaseDir | Out-Null
Copy-Item -LiteralPath $setup.FullName -Destination $releaseSetupPath -Force

Get-Item -LiteralPath $releaseSetupPath
