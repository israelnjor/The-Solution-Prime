<#
.SYNOPSIS
  Builds the Prime Word desktop app (release) -> apps\word\src-tauri\target\release\prime-word.exe

  If Prime Word is running, its .exe is locked against deletion but not against renaming, so the old
  file is moved aside first; the running window keeps working, and the next start uses the new build.
#>
# Not "Stop": Windows PowerShell treats tools' informational stderr output as errors. Exit codes are checked instead.
$ErrorActionPreference = "Continue"
$root = Split-Path -Parent $PSScriptRoot
$release = Join-Path $root "apps\word\src-tauri\target\release"
$exe = Join-Path $release "prime-word.exe"
$env:Path = "$env:USERPROFILE\.cargo\bin;$env:Path"

# Clean up builds moved aside earlier (succeeds only once those copies have exited).
Get-ChildItem $release -Filter "prime-word.old-*.exe" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
if (Test-Path $exe) {
  try { Rename-Item $exe "prime-word.old-$(Get-Date -Format yyyyMMddHHmmss).exe" } catch { Write-Host "Could not move the old build aside: $($_.Exception.Message)" }
}

Push-Location $root
try {
  npm run build | Out-Null
  if ($LASTEXITCODE -ne 0) { throw "engine packages failed to build" }
  Push-Location (Join-Path $root "apps\word")
  try { npx tauri build --no-bundle } finally { Pop-Location }
  if ($LASTEXITCODE -ne 0) { throw "desktop build failed (exit $LASTEXITCODE)" }
} finally { Pop-Location }
Write-Host "Built: $exe"
