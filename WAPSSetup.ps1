# ============================================================
# Setup Script for Windows Power Automation Suite
# ============================================================

# Check for PowerShell 5.1 or later
if ($PSVersionTable.PSVersion -lt [version]"5.1") {
    Write-Host "Error: Windows Power Automation Suite requires PowerShell 5.1 or later." -ForegroundColor Red
    exit
}

# Check for Administrator privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Error: Setup requires Administrator privileges. Please run as Administrator." -ForegroundColor Red
    exit
}

# Get current script directory (Source)
$sourceDir = $PSScriptRoot

# Define Installation Directory
$installPath = "C:\WPAS"

# Prompt user for installation path (Optional, using Read-Host for simplicity)
Write-Host "Enter installation path (Press Enter for default: $installPath): " -NoNewline -ForegroundColor Cyan
$userInput = Read-Host
if (-not [string]::IsNullOrWhiteSpace($userInput)) {
    $installPath = $userInput
}

# Create Install Directory
if (-not (Test-Path $installPath)) {
    New-Item -ItemType Directory -Path $installPath -Force | Out-Null
}

# Copy Files
Write-Host "Installing files to $installPath..." -ForegroundColor Yellow
Copy-Item -Path "$sourceDir\*" -Destination $installPath -Recurse -Force

# Update Config with Install Path
$configFile = Join-Path $installPath "config.json"
$config = @{}
if (Test-Path $configFile) {
    try {
        $jsonObj = Get-Content $configFile | ConvertFrom-Json
        if ($jsonObj) {
            $jsonObj.PSObject.Properties | ForEach-Object {
                $config[$_.Name] = $_.Value
            }
        }
    } catch {}
}
$config["InstallPath"] = $installPath
$config | ConvertTo-Json | Set-Content $configFile

$targetScript = Join-Path $installPath "WAPS ControlCenter.ps1"
$shortcutPath = "$Home\Desktop\Windows Power Automation Suite.lnk"

# Icon Logic
$iconPath = Join-Path $installPath "app.ico"
if (-not (Test-Path $iconPath)) { $iconPath = "$PSHOME\powershell.exe" }

# Create Shortcut
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = "powershell.exe"
$Shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$targetScript`""
$Shortcut.IconLocation = $iconPath
$Shortcut.Description = "Launch the Windows Power Automation Suite (WPAS)"
$Shortcut.WorkingDirectory = $installPath
$Shortcut.Save()

# Create Scheduled Tasks
Write-Host "Creating Scheduled Tasks..." -ForegroundColor Yellow

function Register-WPASTask {
    param($TaskName, $ScriptFile, $Trigger)
    $scriptPath = Join-Path $installPath $ScriptFile
    if (Test-Path $scriptPath) {
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $Trigger -User "System" -RunLevel Highest -Force | Out-Null
        Write-Host "  Task '$TaskName' created." -ForegroundColor Green
    }
}

# 1. System Optimization (Daily at 9 AM)
Register-WPASTask -TaskName "WPAS System Optimization" -ScriptFile "SystemPowerOptimize.ps1" -Trigger (New-ScheduledTaskTrigger -Daily -At 9:00am)

# 2. Log Cleanup (Weekly on Sunday at 12 PM)
Register-WPASTask -TaskName "WPAS Log Cleanup" -ScriptFile "CleanupLogs.ps1" -Trigger (New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 12:00pm)

# 3. Battery Health Check (Daily at 10 AM)
Register-WPASTask -TaskName "WPAS Battery Health" -ScriptFile "BatteryHealth.ps1" -Trigger (New-ScheduledTaskTrigger -Daily -At 10:00am)

# 4. Clear Standby Memory (Daily at 2 PM)
Register-WPASTask -TaskName "WPAS Clear Standby Memory" -ScriptFile "ClearStandbyMemory.ps1" -Trigger (New-ScheduledTaskTrigger -Daily -At 2:00pm)

Write-Host "Shortcut created on Desktop." -ForegroundColor Green
Write-Host "You can now launch the 'Windows Power Automation Suite' from your desktop." -ForegroundColor Cyan