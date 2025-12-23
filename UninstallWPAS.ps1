# ============================================================
# Uninstall Script for Windows Power Automation Suite
# ============================================================

$appName = "Windows Power Automation Suite"
$installPath = $PSScriptRoot 
$desktopShortcut = "$Home\Desktop\$appName.lnk"
$startupShortcut = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\$appName.lnk"

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "      WPAS Uninstaller" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

# Confirmation
Write-Host "This will remove the application and all files in: $installPath" -ForegroundColor Yellow
$confirm = Read-Host "Are you sure you want to continue? (Y/N)"
if ($confirm -ne 'Y' -and $confirm -ne 'y') {
    Write-Host "Uninstallation cancelled." -ForegroundColor Gray
    exit
}

# Check for Administrator privileges
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Warning "Not running as Administrator. Some files or shortcuts might not be removed if they require elevated permissions."
}

# Close the application if running
Write-Host "Checking for running application..." -ForegroundColor Cyan
try {
    $wpasProcesses = Get-CimInstance Win32_Process -Filter "Name like 'powershell%'" | Where-Object { $_.CommandLine -like "*WAPS ControlCenter.ps1*" }
    foreach ($p in $wpasProcesses) {
        Write-Host "Stopping WPAS process (PID: $($p.ProcessId))..." -ForegroundColor Yellow
        Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 1
} catch {
    Write-Warning "Could not automatically close the application. Please ensure WPAS is closed."
}

# Remove Shortcuts
if (Test-Path $desktopShortcut) {
    try {
        Remove-Item $desktopShortcut -Force -ErrorAction Stop
        Write-Host "Removed Desktop shortcut." -ForegroundColor Green
    } catch {
        Write-Error "Failed to remove Desktop shortcut: $_"
    }
}

if (Test-Path $startupShortcut) {
    try {
        Remove-Item $startupShortcut -Force -ErrorAction Stop
        Write-Host "Removed Startup shortcut." -ForegroundColor Green
    } catch {
        Write-Error "Failed to remove Startup shortcut: $_"
    }
}

# Remove Scheduled Tasks
Write-Host "Removing Scheduled Tasks..." -ForegroundColor Yellow
$tasks = @("WPAS System Optimization", "WPAS Log Cleanup", "WPAS Battery Health", "WPAS Clear Standby Memory")
foreach ($task in $tasks) {
    if (Get-ScheduledTask -TaskName $task -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $task -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "  Removed task: $task" -ForegroundColor Green
    }
}

# Remove Installation Directory
Write-Host "Scheduling removal of installation directory..." -ForegroundColor Yellow

# Create a self-deleting batch file to remove the directory
$batchPath = "$env:TEMP\wpas_uninstall.bat"
$batchContent = @"
@echo off
timeout /t 3 /nobreak > NUL
rmdir /s /q "$installPath"
del "%~f0"
"@
Set-Content -Path $batchPath -Value $batchContent

Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$batchPath`"" -WindowStyle Hidden

Write-Host "Uninstallation initiated. The folder will be removed in a few seconds." -ForegroundColor Green
Write-Host "Goodbye!" -ForegroundColor Cyan
Start-Sleep -Seconds 2