# ============================================================
# WPAS Battery Health Logger
# ============================================================

$scriptPath = $PSScriptRoot
$logDir = Join-Path $scriptPath "logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

$csvFile = Join-Path $logDir "BatteryHistory.csv"

# Initialize CSV if it doesn't exist
if (-not (Test-Path $csvFile)) {
    "Timestamp,DesignCapacity,FullChargeCapacity,WearLevel,ChargePercent" | Set-Content $csvFile -Encoding UTF8
}

try {
    $battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction Stop | Select-Object -First 1
    
    if ($battery) {
        $designCap = $battery.DesignCapacity
        $fullCap = $battery.FullChargeCapacity
        $charge = $battery.EstimatedChargeRemaining
        
        # Sanity check for desktop/VM or bad data
        if (-not $designCap -or $designCap -eq 0) { $designCap = 1 }
        if (-not $fullCap) { $fullCap = $designCap }

        $wearLevel = [math]::Round((1 - ($fullCap / $designCap)) * 100, 2)
        
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        
        # Append to CSV
        "$timestamp,$designCap,$fullCap,$wearLevel,$charge" | Add-Content $csvFile -Encoding UTF8
    } else {
        # Log no battery detected (optional, usually skip to avoid spamming logs on desktops)
    }
} catch {
    # Log error to debug log
    $debugLog = Join-Path $logDir "WAPS_Debug.log"
    $err = "BatteryHealth Error: $($_.Exception.Message)"
    Add-Content -Path $debugLog -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $err"
}