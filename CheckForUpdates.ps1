# ============================================================
# WPAS Update Checker
# ============================================================

$scriptPath = $PSScriptRoot
$log = Join-Path $scriptPath "logs\UpdateCheck.log"
$localVersionFile = Join-Path $scriptPath "version.json"
$remoteUrl = "https://raw.githubusercontent.com/Kamalesh/WPAS/main/version.json" # Placeholder URL

# Ensure log directory exists
$logDir = Split-Path $log -Parent
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

function Write-Log {
    param($Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $log -Value "[$timestamp] $Message"
}

Write-Log "Starting Update Check..."

try {
    if (Test-Path $localVersionFile) {
        $localData = Get-Content $localVersionFile | ConvertFrom-Json
        $localVersion = [version]$localData.version
        
        Write-Log "Local Version: $localVersion"

        # Fetch remote version
        $remoteData = Invoke-RestMethod -Uri $remoteUrl -ErrorAction Stop
        $remoteVersion = [version]$remoteData.version

        Write-Log "Remote Version: $remoteVersion"

        if ($remoteVersion -gt $localVersion) {
            Write-Log "Update Available! New version: $remoteVersion"
            # In a real scenario, this might trigger a notification or download
            # For now, we log it.
        } else {
            Write-Log "WPAS is up to date."
        }
    } else {
        Write-Log "Error: Local version.json not found."
    }
} catch {
    Write-Log "Update Check Failed: $($_.Exception.Message)"
}