$ErrorActionPreference = "Stop"

# Run from Windows PowerShell:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\pipeline\run_win32com_export_step4_01_05.ps1

$repoRoot = Split-Path -Parent -LiteralPath $PSScriptRoot
$scriptPath = Join-Path $PSScriptRoot "export_step4_slides_01_05_win32com.py"
$outPath = Join-Path $repoRoot "Step4\deck_summer_product_rnd_pages01-05_editable_com_v1.pptx"

Write-Output "[INFO] repoRoot: $repoRoot"
Write-Output "[INFO] script  : $scriptPath"
Write-Output "[INFO] output  : $outPath"
Write-Output "[STEP] Validate Python script path"

if (-not (Test-Path -LiteralPath $scriptPath)) {
  throw "Python script not found: $scriptPath"
}

try {
  Write-Output "[STEP] Check pywin32 availability"
  py -3 -c "import win32com.client; print('WIN32COM_OK')" | Out-Host
} catch {
  Write-Output "[WARN] pywin32 not found, installing..."
  py -3 -m pip install pywin32
}

Write-Output "[STEP] Run COM export"
$env:STEP4_ROOT = $repoRoot
$env:OUTPUT_PPT = $outPath
py -3 $scriptPath

if (Test-Path -LiteralPath $outPath) {
  Write-Output "[OK] Generated: $outPath"
} else {
  throw "Export failed, output not found: $outPath"
}
