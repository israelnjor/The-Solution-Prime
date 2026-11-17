<#
.SYNOPSIS
  Builds the Prime Word installer -> apps\word\src-tauri\target\release\bundle\nsis\Prime Word_<version>_x64-setup.exe

  Installs for the current user only (no administrator prompt): Start menu entry under "Prime Productivity",
  "Open with Prime Word" for .docx files, and an uninstaller in Settings > Apps.
  The first build downloads the NSIS installer tools from Tauri's official releases (hash-checked by Tauri).

  Updates: with the signing key in %USERPROFILE%\.prime-word\updater.key, the installer is signed for
  Prime Word's updater, and release\v<version>\ gets what to publish: the installer, its signature and
  latest.json. Publish them as a GitHub release named v<version> (see docs\releasing.md). Without the key,
  the installer is built as before, without update files.

.PARAMETER Version
  The new version (e.g. 0.1.1), set before building. Every release needs a higher version than the last
  one, or installed copies will not see it as an update.

.PARAMETER Notes
  What is new in this version, shown in Prime Word when it offers the update.
#>
param([string]$Version = "", [string]$Notes = "")
$ErrorActionPreference = "Continue"
$root = Split-Path -Parent $PSScriptRoot
$release = Join-Path $root "apps\word\src-tauri\target\release"
$exe = Join-Path $release "prime-word.exe"
$conf = Join-Path $root "apps\word\src-tauri\tauri.conf.json"
$env:Path = "$env:USERPROFILE\.cargo\bin;$env:Path"

# The version: set when given (tauri.conf.json is where the app reads it from).
if ($Version) {
  if ($Version -notmatch '^\d+\.\d+\.\d+$') { throw "Version must look like 0.1.1" }
  $text = Get-Content $conf -Raw
  $text = [regex]::Replace($text, '"version":\s*"[^"]+"', "`"version`": `"$Version`"", 1)
  [IO.File]::WriteAllText($conf, $text)
}
$version = (Get-Content $conf -Raw | ConvertFrom-Json).version

# Signing for updates: the private key stays on this computer (never in the repository).
$key = Join-Path $env:USERPROFILE ".prime-word\updater.key"
$extra = @()
if (Test-Path $key) {
  $env:TAURI_SIGNING_PRIVATE_KEY = (Get-Content $key -Raw).Trim()
  # The key has no password: an empty one, set through Windows itself (PowerShell's $env:X = "" deletes
  # the variable instead, and Tauri would then stop and ask for a password).
  if (-not ("Win32.Env" -as [type])) {
    Add-Type -Namespace Win32 -Name Env -MemberDefinition '[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] public static extern bool SetEnvironmentVariable(string name, string value);'
  }
  [Win32.Env]::SetEnvironmentVariable("TAURI_SIGNING_PRIVATE_KEY_PASSWORD", "") | Out-Null
} else {
  Write-Host "No update signing key at $key : building without update files."
  $extra = @("--config", '{"bundle":{"createUpdaterArtifacts":false}}')
}

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
  try { npx tauri build --bundles nsis @extra } finally { Pop-Location }
  if ($LASTEXITCODE -ne 0) { throw "installer build failed (exit $LASTEXITCODE)" }
} finally { Pop-Location }

$nsis = Join-Path $release "bundle\nsis"
$installer = Get-ChildItem $nsis -Filter "*_${version}_x64-setup.exe" | Select-Object -First 1
Get-ChildItem $nsis -Filter "*-setup.exe" | ForEach-Object { Write-Host "Built: $($_.FullName)" }

# What to publish for the updater: the installer (a name without spaces), its signature, and latest.json.
$sig = if ($installer) { "$($installer.FullName).sig" } else { "" }
if ($installer -and (Test-Path $sig)) {
  $out = Join-Path $root "release\v$version"
  New-Item -ItemType Directory -Force $out | Out-Null
  $name = "PrimeWord_${version}_x64-setup.exe"
  Copy-Item $installer.FullName (Join-Path $out $name) -Force
  Copy-Item $sig (Join-Path $out "$name.sig") -Force
  $latest = [ordered]@{
    version   = $version
    notes     = $Notes
    pub_date  = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    platforms = [ordered]@{
      "windows-x86_64" = [ordered]@{
        signature = (Get-Content $sig -Raw).Trim()
        url       = "https://github.com/israelnjor/The-Solution-Prime/releases/download/v$version/$name"
      }
    }
  }
  [IO.File]::WriteAllText((Join-Path $out "latest.json"), ($latest | ConvertTo-Json -Depth 5))
  Write-Host "Release files: $out (publish as GitHub release v${version}: $name, $name.sig, latest.json)"
}
