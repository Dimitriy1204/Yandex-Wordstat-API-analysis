# Build yandex_wordstat_agent.exe (stops running instances first).
$ErrorActionPreference = "Stop"
$ProjectDir = $PSScriptRoot
Set-Location $ProjectDir

Write-Host "=== Stopping agent processes ===" -ForegroundColor Cyan
& "$ProjectDir\scripts\stop_agent.ps1"

$distExe = Join-Path $ProjectDir "dist\yandex_wordstat_agent.exe"
if (Test-Path $distExe) {
    Write-Host "Waiting for file unlock: $distExe" -ForegroundColor Yellow
    $deadline = (Get-Date).AddSeconds(15)
    while ((Get-Date) -lt $deadline) {
        try {
            $fs = [System.IO.File]::Open($distExe, "Open", "ReadWrite", "None")
            $fs.Close()
            break
        } catch {
            Start-Sleep -Milliseconds 500
        }
    }
}

Write-Host "=== PyInstaller ===" -ForegroundColor Cyan
python -m PyInstaller yandex_wordstat_agent.spec --noconfirm
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$tplOk = (Test-Path (Join-Path $ProjectDir "yandex_analysis_exemple.xlsx")) -or `
         (Test-Path (Join-Path $ProjectDir "yandex_analysis_example.xlsx"))
if (-not $tplOk) {
    Write-Warning "yandex_analysis_exemple.xlsx not found - template will not be embedded in exe."
}

Write-Host "=== Done: dist\yandex_wordstat_agent.exe ===" -ForegroundColor Green
