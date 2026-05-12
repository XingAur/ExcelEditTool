$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $scriptDir
$distDir = Join-Path $root "dist\ExcelEditTool"
$releaseDir = Join-Path $root "release"
$stageDir = Join-Path $root "build\installer"
$setupPath = Join-Path $releaseDir "ExcelEditTool_Setup.exe"
$payloadPath = Join-Path $stageDir "ExcelEditToolPayload.zip"
$webView2InstallerPath = Join-Path $root "packaging\MicrosoftEdgeWebView2Setup.exe"
$webView2BootstrapperUrl = "https://go.microsoft.com/fwlink/p/?LinkId=2124703"
$iexpressPath = Join-Path $env:WINDIR "System32\iexpress.exe"

if (-not (Test-Path -LiteralPath $iexpressPath)) {
  throw "IExpress was not found. This script requires Windows' built-in iexpress.exe."
}

Push-Location $root
try {
  & (Join-Path $root "scripts\package_client.ps1")

  if (-not (Test-Path -LiteralPath $webView2InstallerPath)) {
    Invoke-WebRequest `
      -Uri $webView2BootstrapperUrl `
      -OutFile $webView2InstallerPath `
      -UseBasicParsing
  }

  if (Test-Path -LiteralPath $stageDir) {
    Remove-Item -LiteralPath $stageDir -Recurse -Force
  }
  New-Item -ItemType Directory -Force -Path $stageDir | Out-Null
  New-Item -ItemType Directory -Force -Path $releaseDir | Out-Null

  Compress-Archive `
    -Path (Join-Path $distDir "*") `
    -DestinationPath $payloadPath `
    -Force

  Copy-Item `
    -LiteralPath (Join-Path $root "packaging\install.ps1") `
    -Destination (Join-Path $stageDir "install.ps1") `
    -Force

  Copy-Item `
    -LiteralPath $webView2InstallerPath `
    -Destination (Join-Path $stageDir "MicrosoftEdgeWebView2Setup.exe") `
    -Force

  if (Test-Path -LiteralPath $setupPath) {
    Remove-Item -LiteralPath $setupPath -Force
  }

  $sourceRoot = $stageDir.TrimEnd("\") + "\"
  $sedPath = Join-Path $stageDir "ExcelEditTool.iexpress.sed"
  $sed = @"
[Version]
Class=IEXPRESS
SEDVersion=3
[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=0
HideExtractAnimation=1
UseLongFileName=1
InsideCompressed=0
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=%InstallPrompt%
DisplayLicense=%DisplayLicense%
FinishMessage=%FinishMessage%
TargetName=%TargetName%
FriendlyName=%FriendlyName%
AppLaunched=%AppLaunched%
PostInstallCmd=%PostInstallCmd%
AdminQuietInstCmd=%AdminQuietInstCmd%
UserQuietInstCmd=%UserQuietInstCmd%
SourceFiles=SourceFiles
[Strings]
InstallPrompt=
DisplayLicense=
FinishMessage=ExcelEditTool has been installed.
TargetName=$setupPath
FriendlyName=ExcelEditTool Setup
AppLaunched=powershell.exe -NoProfile -ExecutionPolicy Bypass -File install.ps1
PostInstallCmd=<None>
AdminQuietInstCmd=
UserQuietInstCmd=
FILE0=install.ps1
FILE1=ExcelEditToolPayload.zip
FILE2=MicrosoftEdgeWebView2Setup.exe
[SourceFiles]
SourceFiles0=$sourceRoot
[SourceFiles0]
%FILE0%=
%FILE1%=
%FILE2%=
"@

  Set-Content -LiteralPath $sedPath -Value $sed -Encoding ASCII
  $iexpressProcess = Start-Process `
    -FilePath $iexpressPath `
    -ArgumentList "/N", "/Q", $sedPath `
    -Wait `
    -PassThru `
    -WindowStyle Hidden
  $iexpressExitCode = $iexpressProcess.ExitCode

  if (-not (Test-Path -LiteralPath $setupPath)) {
    throw "IExpress did not create $setupPath. Exit code: $iexpressExitCode."
  }

  Get-Item -LiteralPath $setupPath
  $global:LASTEXITCODE = 0
}
finally {
  Pop-Location
}
