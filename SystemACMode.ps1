# ============================================================
# System AC Mode Script (with logging)
# ============================================================

$scriptPath = $PSScriptRoot
$log = Join-Path $scriptPath "logs\SystemPowerOptimize.log"
Add-Content -Path $log -Value "$(Get-Date) - Running SystemACMode.ps1"

# Switch to Balanced (or Performance if you prefer)
powercfg /setactive SCHEME_BALANCED

# CPU min state higher on AC
powercfg -setacvalueindex SCHEME_BALANCED SUB_PROCESSOR PROCTHROTTLEMIN 20

# Wireless adapter medium power saving
powercfg -setacvalueindex SCHEME_BALANCED SUB_WIFI POWER_SAVEMODE 1

# PCIe ASPM moderate
powercfg -setacvalueindex SCHEME_BALANCED SUB_PCIEXPRESS ASPM 1

# Ensure Turbo Boost is enabled (Max 100%)
powercfg -setacvalueindex SCHEME_BALANCED SUB_PROCESSOR PROCTHROTTLEMAX 100

# Apply changes
powercfg /setactive SCHEME_BALANCED

Add-Content -Path $log -Value "$(Get-Date) - Completed SystemACMode.ps1"