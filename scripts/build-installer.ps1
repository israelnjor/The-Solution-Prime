<#
.SYNOPSIS
  Builds the Prime Word installer -> apps\word\src-tauri\target\release\bundle\nsis\Prime Word_<version>_x64-setup.exe

  Installs for the current user only (no administrator prompt): Start menu entry under "Prime Productivity",
  "Open with Prime Word" for .docx files, and an uninstaller in Settings > Apps.
  The first build downloads the NSIS installer tools from Tauri's official releases (hash-checked by Tauri).
#>
$ErrorActionPreference = "Continue"
$root = Split-Path -Parent $PSScriptRoot
$release = Join-Path $root "apps\word\src-tauri\target\release"
$exe = Join-Path $release "prime-word.exe"
$env:Path = "$env:USERPROFILE\.cargo\bin;$env:Path"

# A running Prime Word locks its .exe against deletion but not renaming (see build-desktop.ps1).
Get-ChildItem $release -Filter "prime-word.old-*.exe" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
if (Test-Path $exe) {
  try { Rename-Item $exe "prime-word.old-$(Get-Date -Format yyyyMMddHHmmss).exe" } catch { Write-Host "Could not move the old build aside: $($_.Exception.Message)" }
}

Push-Location $root
try {
  npm run build | Out-Null
  if ($LASTEXITCODE -ne 0) { throw "engine packages failed to build" }
  Push-Location (Join-Path $root "apps\word")
  try { npx tauri build --bundles nsis } finally { Pop-Location }
  if ($LASTEXITCODE -ne 0) { throw "installer build failed (exit $LASTEXITCODE)" }
} finally { Pop-Location }
Get-ChildItem (Join-Path $root "apps\word\src-tauri\target\release\bundle\nsis") -Filter "*-setup.exe" | ForEach-Object { Write-Host "Built: $($_.FullName)" }
