# ============================================================
# Performance Mode Toggle
# ============================================================

$scriptPath = $PSScriptRoot
$log = Join-Path $scriptPath "logs\SystemPowerOptimize.log"
Add-Content -Path $log -Value "$(Get-Date) - Running PerformanceMode.ps1"

# Switch to High Performance
powercfg /setactive SCHEME_MIN

# CPU min state 100%
powercfg -setacvalueindex SCHEME_MIN SUB_PROCESSOR PROCTHROTTLEMIN 100

# Wireless max performance
powercfg -setacvalueindex SCHEME_MIN SUB_WIFI POWER_SAVEMODE 0

# PCIe ASPM off
powercfg -setacvalueindex SCHEME_MIN SUB_PCIEXPRESS ASPM 0

Add-Content -Path $log -Value "$(Get-Date) - Completed PerformanceMode.ps1"