<#
.SYNOPSIS
  Round-trip your own Word documents through Prime Word's DOCX engine and get a scored report.

.EXAMPLE
  .\scripts\try-docx.ps1                                           # opens a file picker
  .\scripts\try-docx.ps1 "C:\Users\me\Documents\report.docx"
  .\scripts\try-docx.ps1 "C:\Users\me\Documents\Contracts"       # every .docx in the folder
  .\scripts\try-docx.ps1 report.docx -OpenInWord                   # also open original + result in Word
  .\scripts\try-docx.ps1 -Pdf                                      # pick files; also export PDFs with our layout engine

  Inputs are copied into a private, git-ignored folder; your originals are never modified.
#>
param(
  [string]$Path,
  [switch]$OpenInWord,
  [switch]$Pdf,     # also export each document to PDF with Prime Word's layout engine
  [switch]$NoOpen   # do not open the report automatically
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$exp = Join-Path $root "experiments\e1-docx-roundtrip"
$python = Join-Path $root ".venv\Scripts\python.exe"

if (-not (Test-Path (Join-Path $root "packages\docx\dist\cli.js"))) {
  Write-Host "Building the DOCX engine..."
  Push-Location $root; npm install --no-fund --no-audit | Out-Null; npm run build | Out-Null; Pop-Location
}

if (-not $Path) {
  Add-Type -AssemblyName System.Windows.Forms
  $dialog = New-Object System.Windows.Forms.OpenFileDialog
  $dialog.Title = "Choose Word documents to test with Prime Word"
  $dialog.Filter = "Word documents (*.docx)|*.docx"
  $dialog.Multiselect = $true
  $dialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
  if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { Write-Host "No file chosen."; exit 0 }
  $files = $dialog.FileNames | ForEach-Object { Get-Item $_ }
} elseif (-not (Test-Path $Path)) {
  Write-Host "Can't find '$Path'." -ForegroundColor Yellow
  Write-Host "Tip: run the script with no path to pick a file, or drag a .docx onto this window to paste its path."
  exit 1
} else {
  $files = if (Test-Path $Path -PathType Container) { Get-ChildItem $Path -Filter *.docx -File } else { Get-Item $Path }
}
$files = $files | Where-Object { -not $_.Name.StartsWith('~$') }
if (-not $files) { Write-Host "No .docx files found at $Path" -ForegroundColor Yellow; exit 1 }

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$inbox = Join-Path $exp "corpus-private\try-$stamp"
New-Item -ItemType Directory -Force $inbox | Out-Null
foreach ($f in $files) { Copy-Item $f.FullName (Join-Path $inbox $f.Name) }
Write-Host "Testing $(@($files).Count) document(s)..."

Push-Location $exp
try { & $python run.py --candidates prime-docx --fixtures-dir $inbox } finally { Pop-Location }

$run = Get-ChildItem (Join-Path $exp "results") -Directory | Sort-Object Name | Select-Object -Last 1
Write-Host "`nReport:  $($run.FullName)\report.md"
Write-Host "Outputs: $($run.FullName)\prime-docx\  (open these in Word to judge for yourself)"
if (-not $NoOpen) { Start-Process (Join-Path $run.FullName "report.md") }
if ($Pdf) {
  $pdfDir = Join-Path $run.FullName "prime-pdf"
  New-Item -ItemType Directory -Force $pdfDir | Out-Null
  foreach ($f in $files) {
    $out = Join-Path $pdfDir ([IO.Path]::ChangeExtension($f.Name, ".pdf"))
    node (Join-Path $root "packages\pdf\dist\cli.js") (Join-Path $inbox $f.Name) $out
    if (-not $NoOpen) { Start-Process $out }
  }
  Write-Host "PDFs:    $pdfDir  (compare with File > Save As > PDF in Word)"
}
if ($OpenInWord) {
  foreach ($f in $files) {
    Start-Process (Join-Path $inbox $f.Name)
    Start-Process (Join-Path $run.FullName "prime-docx\$($f.Name)")
  }
}
