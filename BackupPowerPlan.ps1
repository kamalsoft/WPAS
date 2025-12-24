# ============================================================
# WPAS Power Plan Backup
# ============================================================

$scriptPath = $PSScriptRoot
$backupDir = Join-Path $scriptPath "backups"

# Ensure backup directory exists
if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }

try {
    # Get Active Scheme GUID
    $activeScheme = powercfg /getactivescheme
    if ($activeScheme -match "GUID: (.*?) \(") {
        $guid = $matches[1]
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmm"
        $backupFile = Join-Path $backupDir "PowerPlanBackup_$timestamp.pow"
        
        # Export
        Write-Host "Exporting active power plan ($guid)..." -ForegroundColor Cyan
        powercfg /export $backupFile $guid
        
        if (Test-Path $backupFile) {
            Write-Host "Backup successful: $backupFile" -ForegroundColor Green
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show("Power Plan backed up successfully to:`n$backupFile", "WPAS Backup", "OK", "Information")
        } else {
            throw "Export command failed to create file."
        }
    } else {
        throw "Could not determine active power scheme GUID."
    }
} catch {
    Write-Error "Backup failed: $($_.Exception.Message)"
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("Backup failed:`n$($_.Exception.Message)", "WPAS Error", "OK", "Error")
}