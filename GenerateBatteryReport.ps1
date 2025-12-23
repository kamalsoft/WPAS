# ============================================================
# WPAS Visual Battery Health Report Generator
# ============================================================

$scriptPath = $PSScriptRoot
$reportDir = Join-Path $scriptPath "logs"

# Ensure log directory exists
if (-not (Test-Path $reportDir)) { New-Item -ItemType Directory -Path $reportDir -Force | Out-Null }
$reportFile = Join-Path $reportDir "BatteryHealthReport.html"

Write-Host "Gathering battery information..." -ForegroundColor Cyan

# Get System Info
$computer = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -First 1
$os = Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -First 1

# Get Battery Info
$battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue | Select-Object -First 1

if (-not $battery) {
    Write-Warning "No battery detected on this system."
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head><title>Battery Report</title><style>body{font-family:sans-serif;padding:50px;text-align:center;color:#555;}</style></head>
<body><h1>No Battery Detected</h1><p>Windows could not retrieve battery information. This report is intended for laptops and tablets.</p></body>
</html>
"@
    $html | Set-Content -Path $reportFile -Encoding UTF8
    Start-Process $reportFile
    exit
}

# Calculations
$designCap = $battery.DesignCapacity
$fullCap = $battery.FullChargeCapacity
$currentCharge = $battery.EstimatedChargeRemaining

# Handle potential missing data (0 or null)
if (-not $designCap -or $designCap -eq 0) { $designCap = 1 } 
if (-not $fullCap) { $fullCap = $designCap }

$wearLevel = 0
if ($designCap -gt 0) {
    $wearLevel = [math]::Round((1 - ($fullCap / $designCap)) * 100, 2)
}
$healthPercent = 100 - $wearLevel

# Status Mapping
$statusMap = @{
    1 = "Discharging"; 2 = "On AC (Not Charging)"; 3 = "Fully Charged"; 4 = "Low"; 
    5 = "Critical"; 6 = "Charging"; 7 = "Charging (High)"; 8 = "Charging (Low)"; 9 = "Charging (Critical)"
}
$statusStr = if ($statusMap.ContainsKey($battery.BatteryStatus)) { $statusMap[$battery.BatteryStatus] } else { "Unknown ($($battery.BatteryStatus))" }

# Runtime formatting
$runtime = $battery.EstimatedRunTime
$runtimeStr = if ($runtime -gt 10000) { "Calculating / On AC" } else { "$runtime minutes" }

# CSS Colors based on health/charge
$healthColor = if ($healthPercent -ge 80) { "#28a745" } elseif ($healthPercent -ge 50) { "#ffc107" } else { "#dc3545" }
$chargeColor = if ($currentCharge -ge 20) { "#007bff" } else { "#dc3545" }

# HTML Generation
$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>WPAS Battery Report</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f0f2f5; color: #333; margin: 0; padding: 20px; }
        .container { max-width: 700px; margin: 0 auto; background: #fff; padding: 40px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); }
        h1 { text-align: center; color: #2c3e50; margin-bottom: 5px; }
        .subtitle { text-align: center; color: #7f8c8d; font-size: 0.9em; margin-bottom: 30px; }
        .card { background: #fff; border: 1px solid #e1e4e8; padding: 25px; border-radius: 8px; margin-bottom: 25px; }
        .metric-row { display: flex; justify-content: space-between; margin-bottom: 12px; border-bottom: 1px solid #f0f0f0; padding-bottom: 8px; }
        .metric-row:last-child { border-bottom: none; }
        .label { font-weight: 600; color: #555; }
        .value { color: #000; }
        
        .progress-section { margin-bottom: 20px; }
        .progress-header { display: flex; justify-content: space-between; margin-bottom: 8px; font-weight: bold; font-size: 0.95em; }
        .progress-bg { background: #e9ecef; border-radius: 6px; height: 24px; width: 100%; overflow: hidden; }
        .progress-fill { height: 100%; text-align: center; color: white; line-height: 24px; font-size: 13px; font-weight: 600; transition: width 0.6s ease; }
        
        .footer { text-align: center; margin-top: 40px; font-size: 12px; color: #aaa; border-top: 1px solid #eee; padding-top: 20px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Battery Health Report</h1>
        <div class="subtitle">Generated for <strong>$($computer.Name)</strong> on $(Get-Date -Format "yyyy-MM-dd HH:mm")</div>

        <div class="card">
            <div class="progress-section">
                <div class="progress-header">
                    <span>Current Charge</span>
                    <span>$currentCharge%</span>
                </div>
                <div class="progress-bg">
                    <div class="progress-fill" style="width: $currentCharge%; background-color: $chargeColor;">$currentCharge%</div>
                </div>
            </div>

            <div class="progress-section">
                <div class="progress-header">
                    <span>Battery Health (Capacity)</span>
                    <span>$healthPercent%</span>
                </div>
                <div class="progress-bg">
                    <div class="progress-fill" style="width: $healthPercent%; background-color: $healthColor;">$healthPercent%</div>
                </div>
            </div>
        </div>

        <div class="card">
            <div class="metric-row"><span class="label">Status</span><span class="value">$statusStr</span></div>
            <div class="metric-row"><span class="label">Design Capacity</span><span class="value">$designCap mWh</span></div>
            <div class="metric-row"><span class="label">Full Charge Capacity</span><span class="value">$fullCap mWh</span></div>
            <div class="metric-row"><span class="label">Wear Level</span><span class="value">$wearLevel%</span></div>
            <div class="metric-row"><span class="label">Voltage</span><span class="value">$($battery.DesignVoltage) mV</span></div>
            <div class="metric-row"><span class="label">Estimated Runtime</span><span class="value">$runtimeStr</span></div>
        </div>

        <div class="footer">Generated by Windows Power Automation Suite (WPAS)</div>
    </div>
</body>
</html>
"@

$html | Set-Content -Path $reportFile -Encoding UTF8
Write-Host "Report generated: $reportFile" -ForegroundColor Green
Start-Process $reportFile