# Windows Power Automation Suite (WPAS)

WPAS is a comprehensive system utility designed to optimize power management, monitor system health, and automate common maintenance tasks for Windows PCs.

## Features

### Dashboard

- **Real-time Monitoring**: View Power Source, Active Power Plan, GPU Usage, CPU Temperature, System Uptime, and RAM Usage.
- **Battery Health**: Displays Design Capacity, Full Charge Capacity, and Wear Level.
- **Quick Actions**: Restart Explorer, Disk Cleanup, Check for Updates, Open Event Viewer.
- **Energy Diagnostics**: Generate a detailed Windows Power Efficiency Diagnostics Report.
- **Auto-Switch**: Automatically switches power plans and settings based on AC/Battery status.

### Automation & Tools

- **Launcher**: Quick access to optimization scripts (System AC Mode, Power Saver, etc.).
- **Startup Apps**: Manage (Enable/Disable/Delete/Add) applications that start with Windows.
- **Scheduled Tasks**: View and trigger Windows Scheduled Tasks.
- **Services**: Start and Stop essential Windows services.
- **Processes**: Monitor top CPU-consuming processes and terminate them if necessary.
- **System Clean**: Clear Temporary files, Recycle Bin, and Browser Cache.
- **Network**: View IP address, connection status, and Flush DNS.

### Customization

- **Settings**: Configure which scripts run on AC vs. Battery.
- **Dark Mode**: Toggle a dark theme for the application.
- **Tray Icon**: The application minimizes to the system tray for background monitoring.

## Installation

1.  **Download/Prepare**: Ensure all script files are located in `C:\Tools\pcfix`.
2.  **Run Setup**:
    - Open PowerShell as Administrator.
    - Navigate to the directory: `cd C:\Tools\pcfix`
    - Run the setup script: `.\WAPSSetup.ps1`
3.  **Launch**: A shortcut named "Windows Power Automation Suite" will be created on your Desktop. Double-click it to start.

## Uninstallation

To completely remove WPAS from your system:

1.  **Close the App**: Right-click the WPAS tray icon (near the clock) and select **Exit**.
2.  **Remove Startup Entry**:
    - If you enabled "Run at Windows Startup" in Settings, delete the shortcut from:
      `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\PCFix Control Center.lnk`
3.  **Delete Files**:
    - Delete the installation folder: `C:\Tools\pcfix`
    - Delete the Desktop shortcut.

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 or later
- Administrator privileges are required for some features (e.g., Power Plans, Service Management, Energy Report).

## License

Provided as-is for personal use.
