# ============================================================
# Cleanup Script for Old Logs
# ============================================================

$scriptPath = $PSScriptRoot
$logPath = Join-Path $scriptPath "logs"
$days = 30  # delete logs older than 30 days

Get-ChildItem -Path $logPath -File |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$days) } |
    Remove-Item -Force

Add-Content -Path "$logPath\Cleanup.log" -Value "$(Get-Date) - Cleanup completed"