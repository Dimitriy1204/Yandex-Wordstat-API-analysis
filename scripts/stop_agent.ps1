# Stop all Yandex Wordstat Agent processes (exe and python on port 8000).
$ErrorActionPreference = "SilentlyContinue"

foreach ($proc in Get-Process -Name "yandex_wordstat_agent" -ErrorAction SilentlyContinue) {
    Write-Host "Stopping $($proc.ProcessName) PID $($proc.Id)"
    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
}

$port = if ($env:PORT) { [int]$env:PORT } else { 8000 }
netstat -ano | Select-String ":$port\s+.*LISTENING" | ForEach-Object {
    $pid = ($_.Line -split '\s+')[-1]
    if ($pid -match '^\d+$') {
        $p = Get-Process -Id $pid -ErrorAction SilentlyContinue
        if ($p -and $p.ProcessName -match 'python|yandex_wordstat') {
            Write-Host "Stopping port $port listener PID $pid ($($p.ProcessName))"
            Stop-Process -Id $pid -Force -ErrorAction SilentlyContinue
        }
    }
}

Start-Sleep -Seconds 1
Write-Host "Stop complete."
