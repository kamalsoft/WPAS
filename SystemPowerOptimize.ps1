# ============================================================
# System Power Optimization Script (with logging)
# ============================================================

$scriptPath = $PSScriptRoot
$log = Join-Path $scriptPath "logs\SystemPowerOptimize.log"
Add-Content -Path $log -Value "$(Get-Date) - Running SystemPowerOptimize.ps1"

# 1. Switch to Balanced plan
powercfg /setactive SCHEME_BALANCED

# 2. Intelligent Cooling (If supported)
$lenovoKey = "HKLM:\SOFTWARE\Lenovo\PWRMGRV\Data"
if (Test-Path $lenovoKey) {
    Set-ItemProperty -Path $lenovoKey -Name "ThermalMode" -Value 0 -Force
}

# 3. Disable hidden performance override
$perfKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Power\PowerThrottling"
if (Test-Path $perfKey) {
    Set-ItemProperty -Path $perfKey -Name "PowerThrottlingOff" -Value 0 -Force
}

# 4. Modern Standby tuning
powercfg /setdcvalueindex SCHEME_BALANCED SUB_SLEEP STANDBYIDLE 1
powercfg /setacvalueindex SCHEME_BALANCED SUB_SLEEP STANDBYIDLE 1

# 5. Thunderbolt power save
$tbKey = "HKLM:\SYSTEM\CurrentControlSet\Services\ThunderboltService"
if (Test-Path $tbKey) {
    Set-ItemProperty -Path $tbKey -Name "PowerSaveMode" -Value 1 -Force
}

# 6. Wireless power saving
powercfg -setdcvalueindex SCHEME_BALANCED SUB_WIFI POWER_SAVEMODE 3
powercfg -setacvalueindex SCHEME_BALANCED SUB_WIFI POWER_SAVEMODE 2

# 7. PCIe ASPM
powercfg -setdcvalueindex SCHEME_BALANCED SUB_PCIEXPRESS ASPM 2
powercfg -setacvalueindex SCHEME_BALANCED SUB_PCIEXPRESS ASPM 1

# 8. CPU min state
powercfg -setdcvalueindex SCHEME_BALANCED SUB_PROCESSOR PROCTHROTTLEMIN 5
powercfg -setacvalueindex SCHEME_BALANCED SUB_PROCESSOR PROCTHROTTLEMIN 5

# 9. USB selective suspend
powercfg -setdcvalueindex SCHEME_BALANCED SUB_USB USBSELECTIVE 1
powercfg -setacvalueindex SCHEME_BALANCED SUB_USB USBSELECTIVE 1

# 10. Fix specific USB device (VID_1EA7&PID_0064) Selective Suspend
$usbDevs = Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Enum\USB\VID_1EA7&PID_0064" -ErrorAction SilentlyContinue
foreach ($dev in $usbDevs) {
    $params = Join-Path $dev.PSPath "Device Parameters"
    if (Test-Path $params) {
        # Enable Enhanced Power Management
        Set-ItemProperty -Path $params -Name "EnhancedPowerManagementEnabled" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
        # Allow Idle IRP
        Set-ItemProperty -Path $params -Name "AllowIdleIrpInD3" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
    }
}

# 11. Additional CPU Power Settings
# System Cooling Policy: Active (1) on AC, Passive (0) on DC
powercfg -setacvalueindex SCHEME_BALANCED SUB_PROCESSOR SYSCOOLPOL 1
powercfg -setdcvalueindex SCHEME_BALANCED SUB_PROCESSOR SYSCOOLPOL 0

# Max Processor State (99% on DC disables Turbo Boost to save power)
powercfg -setacvalueindex SCHEME_BALANCED SUB_PROCESSOR PROCTHROTTLEMAX 100
powercfg -setdcvalueindex SCHEME_BALANCED SUB_PROCESSOR PROCTHROTTLEMAX 99

# Apply changes
powercfg /setactive SCHEME_BALANCED

Add-Content -Path $log -Value "$(Get-Date) - Completed SystemPowerOptimize.ps1"