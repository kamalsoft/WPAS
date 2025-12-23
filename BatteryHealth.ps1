# ============================================================
# Battery Health Monitoring Script
# ============================================================

$scriptPath = $PSScriptRoot
$log = Join-Path $scriptPath "logs\BatteryHealth.log"
Add-Content -Path $log -Value "`n$(Get-Date) - Battery Health Check"

$battery = Get-CimInstance -Class Win32_Battery

if ($battery -and $battery.DesignCapacity -gt 0) {
    $design = $battery.DesignCapacity
    $full = $battery.FullChargeCapacity
    $wear = [math]::Round((1 - ($full / $design)) * 100, 2)

    Add-Content -Path $log -Value "Design Capacity: $design mWh"
    Add-Content -Path $log -Value "Full Charge Capacity: $full mWh"
    Add-Content -Path $log -Value "Wear Level: $wear %"
} else {
    Add-Content -Path $log -Value "Battery detected but capacity data is unavailable via WMI."
}