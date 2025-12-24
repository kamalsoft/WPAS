# ============================================================
# WPAS Power Plan Restore
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
$scriptPath = $PSScriptRoot
$backupDir = Join-Path $scriptPath "backups"

# Ensure backup directory exists
if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }

$ofd = New-Object System.Windows.Forms.OpenFileDialog
$ofd.Title = "Select Power Plan Backup to Restore"
$ofd.InitialDirectory = $backupDir
$ofd.Filter = "Power Plan Files (*.pow)|*.pow|All Files (*.*)|*.*"

if ($ofd.ShowDialog() -eq "OK") {
    $backupFile = $ofd.FileName
    try {
        Write-Host "Restoring power plan from: $backupFile" -ForegroundColor Cyan
        
        # Import the plan
        $output = powercfg /import $backupFile
        
        # Parse output to find the new GUID
        # Output format is usually: "Imported Scheme GUID: <GUID>"
        if ($output -match "GUID: (.*)$") {
            $newGuid = $matches[1].Trim()
            
            # Set as active
            powercfg /setactive $newGuid
            
            Write-Host "Successfully imported and activated plan: $newGuid" -ForegroundColor Green
            [System.Windows.Forms.MessageBox]::Show("Power Plan restored and activated successfully!`nGUID: $newGuid", "WPAS Success", "OK", "Information")
        } else {
            # Fallback if output format varies, just notify import success
            Write-Host "Plan imported. Output: $output" -ForegroundColor Yellow
            [System.Windows.Forms.MessageBox]::Show("Power Plan imported successfully.`nPlease check Power Options to activate it.`nOutput: $output", "WPAS Info", "OK", "Information")
        }
    } catch {
        Write-Error "Restore failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to restore power plan:`n$($_.Exception.Message)", "WPAS Error", "OK", "Error")
    }
} else {
    Write-Host "Restore cancelled." -ForegroundColor Gray
}