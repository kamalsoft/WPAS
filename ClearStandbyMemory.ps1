# ============================================================
# Clear Standby Memory Script
# ============================================================

$scriptPath = $PSScriptRoot
$log = Join-Path $scriptPath "logs\ClearStandbyMemory.log"
$logDir = Split-Path $log -Parent
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

Add-Content -Path $log -Value "$(Get-Date) - Attempting to clear standby memory..."

try {
    $code = @"
using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

public class MemoryCleaner {
    [DllImport("ntdll.dll")]
    public static extern int NtSetSystemInformation(int SystemInformationClass, IntPtr SystemInformation, int SystemInformationLength);

    public static void ClearStandbyList() {
        // SystemMemoryListInformation = 80
        int SystemMemoryListInformation = 80;
        
        // MemoryPurgeStandbyList = 4
        int Command = 4;
        
        IntPtr ptr = Marshal.AllocHGlobal(4);
        try {
            Marshal.WriteInt32(ptr, Command);
            int result = NtSetSystemInformation(SystemMemoryListInformation, ptr, 4);
            if (result != 0) {
                throw new Win32Exception(result);
            }
        } finally {
            Marshal.FreeHGlobal(ptr);
        }
    }
}
"@

    Add-Type -TypeDefinition $code -Language CSharp
    [MemoryCleaner]::ClearStandbyList()
    
    Add-Content -Path $log -Value "$(Get-Date) - Standby memory cleared successfully."
} catch {
    $err = $_.Exception.Message
    Add-Content -Path $log -Value "$(Get-Date) - Error: $err"
    Write-Error "Failed to clear standby memory: $err"
}