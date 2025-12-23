# ============================================================
# WPAS Dashboard (Error-Proof Version)
# ============================================================

Write-Host "================ Windows Power Automation Suite ================" -ForegroundColor Cyan

# -------------------------------
# Power Source
# -------------------------------
Add-Type -AssemblyName System.Windows.Forms
$powerStatus = [System.Windows.Forms.SystemInformation]::PowerStatus.PowerLineStatus
Write-Host "`nPower Source:" -ForegroundColor Yellow
if ($powerStatus -eq "Online") {
    Write-Host "  AC Power (Plugged In)" -ForegroundColor Green
} else {
    Write-Host "  Battery Power" -ForegroundColor Red
}

# -------------------------------
# Battery Health (Safe Handling)
# -------------------------------
$battery = Get-CimInstance -Class Win32_Battery

Write-Host "`nBattery Health:" -ForegroundColor Yellow

if ($battery -and $battery.DesignCapacity -and $battery.FullChargeCapacity) {
    $design = $battery.DesignCapacity
    $full = $battery.FullChargeCapacity

    if ($design -gt 0 -and $full -gt 0) {
        $wear = [math]::Round((1 - ($full / $design)) * 100, 2)
        Write-Host "  Design Capacity: $design mWh"
        Write-Host "  Full Charge Capacity: $full mWh"
        Write-Host "  Wear Level: $wear %"
    } else {
        Write-Host "  Battery does not report capacity values."
    }
}
else {
    Write-Host "  Battery health data unavailable on this system."
}

# -------------------------------
# Active Power Plan
# -------------------------------
$currentPlan = powercfg /getactivescheme
Write-Host "`nActive Power Plan:" -ForegroundColor Yellow
Write-Host "  $currentPlan"

# -------------------------------
# CPU Power Settings
# -------------------------------
Write-Host "`nCPU Power Settings:" -ForegroundColor Yellow
powercfg /query SCHEME_CURRENT SUB_PROCESSOR | Select-String "Minimum" -Context 0,1
powercfg /query SCHEME_CURRENT SUB_PROCESSOR | Select-String "Maximum" -Context 0,1

# -------------------------------
# Wi-Fi Power Settings (Safe)
# -------------------------------
Write-Host "`nWi-Fi Power Settings:" -ForegroundColor Yellow
try {
    powercfg /query SCHEME_CURRENT SUB_WIFI | Select-String "Power Saving Mode" -Context 0,1
}
catch {
    Write-Host "  Wi-Fi power settings unavailable on this system."
}

# -------------------------------
# PCIe ASPM
# -------------------------------
Write-Host "`nPCIe ASPM:" -ForegroundColor Yellow
powercfg /query SCHEME_CURRENT SUB_PCIEXPRESS | Select-String "ASPM" -Context 0,1

# -------------------------------
# USB Selective Suspend (Safe)
# -------------------------------
Write-Host "`nUSB Selective Suspend:" -ForegroundColor Yellow
try {
    powercfg /query SCHEME_CURRENT SUB_USB | Select-String "USB Selective Suspend" -Context 0,1
}
catch {
    Write-Host "  USB selective suspend settings unavailable."
}

# -------------------------------
# Last Optimization Runs
# -------------------------------
Write-Host "`nLast Optimization Runs:" -ForegroundColor Yellow
Get-Content "C:\Tools\pcfix\logs\LenovoPowerOptimize.log" -Tail 5

Write-Host "================================================="