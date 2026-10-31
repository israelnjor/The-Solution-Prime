<#
.SYNOPSIS
  Builds the Prime Word installer -> apps\word\src-tauri\target\release\bundle\nsis\Prime Word_<version>_x64-setup.exe

  Installs for the current user only (no administrator prompt): Start menu entry under "Prime Productivity",
  "Open with Prime Word" for .docx files, and an uninstaller in Settings > Apps.
  The first build downloads the NSIS installer tools from Tauri's official releases (hash-checked by Tauri).
#>
$ErrorActionPreference = "Continue"
$root = Split-Path -Parent $PSScriptRoot
$env:Path = "$env:USERPROFILE\.cargo\bin;$env:Path"

Push-Location $root
try {
  npm run build | Out-Null
  if ($LASTEXITCODE -ne 0) { throw "engine packages failed to build" }
  Push-Location (Join-Path $root "apps\word")
  try { npx tauri build --bundles nsis } finally { Pop-Location }
  if ($LASTEXITCODE -ne 0) { throw "installer build failed (exit $LASTEXITCODE)" }
} finally { Pop-Location }
Get-ChildItem (Join-Path $root "apps\word\src-tauri\target\release\bundle\nsis") -Filter "*-setup.exe" | ForEach-Object { Write-Host "Built: $($_.FullName)" }
