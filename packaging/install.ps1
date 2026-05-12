$ErrorActionPreference = "Stop"

$appName = "ExcelEditTool"
$preferredInstallDir = "D:\$appName"
$fallbackInstallDir = Join-Path $env:LOCALAPPDATA "Programs\$appName"
$installDir = if (Test-Path -LiteralPath "D:\") { $preferredInstallDir } else { $fallbackInstallDir }
$payloadPath = Join-Path $PSScriptRoot "ExcelEditToolPayload.zip"
$webView2Installer = Join-Path $PSScriptRoot "MicrosoftEdgeWebView2Setup.exe"
$webView2ProductId = "{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"

function Test-WebView2Runtime {
  $registryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\$webView2ProductId",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\$webView2ProductId",
    "HKCU:\SOFTWARE\Microsoft\EdgeUpdate\Clients\$webView2ProductId"
  )

  foreach ($path in $registryPaths) {
    if (Test-Path $path) {
      $version = (Get-ItemProperty -LiteralPath $path -Name "pv" -ErrorAction SilentlyContinue).pv
      if (-not [string]::IsNullOrWhiteSpace($version)) {
        return $true
      }
    }
  }

  return $false
}

function Install-WebView2Runtime {
  if (Test-WebView2Runtime) {
    return
  }

  if (-not (Test-Path -LiteralPath $webView2Installer)) {
    throw "MicrosoftEdgeWebView2Setup.exe was not found in the installer package."
  }

  $process = Start-Process `
    -FilePath $webView2Installer `
    -ArgumentList "/silent", "/install" `
    -Wait `
    -PassThru `
    -WindowStyle Hidden

  if ($process.ExitCode -ne 0 -and $process.ExitCode -ne 3010) {
    throw "WebView2 Runtime installer failed with exit code $($process.ExitCode)."
  }

  if (-not (Test-WebView2Runtime)) {
    throw "WebView2 Runtime was not detected after installation."
  }
}

function Install-AppFiles {
  if (-not (Test-Path -LiteralPath $payloadPath)) {
    throw "ExcelEditToolPayload.zip was not found in the installer package."
  }

  if (Test-Path -LiteralPath $installDir) {
    Remove-Item -LiteralPath $installDir -Recurse -Force
  }

  New-Item -ItemType Directory -Force -Path $installDir | Out-Null
  Expand-Archive -LiteralPath $payloadPath -DestinationPath $installDir -Force
}

function New-AppShortcut {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ShortcutPath
  )

  $exePath = Join-Path $installDir "$appName.exe"
  $shell = New-Object -ComObject WScript.Shell
  $shortcut = $shell.CreateShortcut($ShortcutPath)
  $shortcut.TargetPath = $exePath
  $shortcut.WorkingDirectory = $installDir
  $shortcut.IconLocation = "$exePath,0"
  $shortcut.Save()
}

Install-WebView2Runtime
Install-AppFiles

$startMenuShortcut = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\$appName.lnk"
$desktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "$appName.lnk"
New-AppShortcut -ShortcutPath $startMenuShortcut
New-AppShortcut -ShortcutPath $desktopShortcut

Start-Process -FilePath (Join-Path $installDir "$appName.exe") -WorkingDirectory $installDir
