# Windows Power Automation Suite (WPAS)

WPAS is a comprehensive system utility designed to optimize power management, monitor system health, and automate common maintenance tasks for Windows PCs.

## Features

### Dashboard

- **Real-time Monitoring**: View Power Source, Active Power Plan, GPU Usage, CPU Temperature, System Uptime, RAM Usage, Network Speed, and Disk Activity.
- **Battery Health**: Displays Design Capacity, Full Charge Capacity, Wear Level, and Charge Percentage.
- **Quick Actions**: Restart Explorer, Disk Cleanup, Check for Updates, Open Event Viewer, Create Restore Point.
- **Energy Diagnostics**: Generate a detailed Windows Power Efficiency Diagnostics Report.
- **Auto-Switch**: Automatically switches power plans and settings based on AC/Battery status.

### Automation & Tools

- **Launcher**: Quick access to optimization scripts (System AC Mode, Power Saver, etc.).
- **Startup Apps**: Manage (Enable/Disable/Delete/Add) applications that start with Windows.
- **Services**: Start and Stop essential Windows services.
- **Processes**: Monitor top CPU-consuming processes and terminate them if necessary.
- **System Clean**: Clear Temporary files, Recycle Bin, Browser Cache, and Standby Memory (RAM).
- **Network**: View IP address, connection status, and Flush DNS.
- **Hardware Specs**: Detailed view of CPU, GPU, Motherboard, and RAM information.

### Scheduled Background Tasks

WPAS installs the following automated tasks to maintain system health:

- **WPAS System Optimization**: Runs daily at 9:00 AM to enforce power plan settings.
- **WPAS Log Cleanup**: Runs weekly on Sundays at 12:00 PM to manage log file sizes.
- **WPAS Battery Health**: Runs daily at 10:00 AM to record battery statistics.
- **WPAS Clear Standby Memory**: Runs daily at 2:00 PM to free up cached RAM.
- **WPAS Update Check**: Runs weekly on Mondays at 9:00 AM to check for application updates.

### Customization

- **Settings**: Configure which scripts run on AC vs. Battery.
- **Dark Mode**: Toggle a dark theme for the application.
- **Tray Icon**: The application minimizes to the system tray for background monitoring.

## Benefits of Using WPAS

- **Extended Battery Life**: Automated power plan switching and specific optimizations for battery mode help prolong usage time.
- **Enhanced Performance**: Tools to clear Standby Memory and manage startup apps ensure your system runs smoothly.
- **Centralized Control**: Access critical system tools, logs, and settings from a single interface.
- **System Health Visibility**: Real-time metrics on temperature, wear levels, and resource usage help prevent overheating and hardware degradation.

## Technology Stack

- **Language**: PowerShell 5.1+
- **GUI Framework**: Windows Forms (System.Windows.Forms)
- **System Interface**: WMI (Windows Management Instrumentation) & CIM (Common Information Model)
- **Native Interop**: C# (P/Invoke) for low-level memory management.

## Usage

1.  **Launch**: Double-click the "Windows Power Automation Suite" shortcut on your Desktop.
2.  **Dashboard**: Monitor system stats immediately. Use "Quick Actions" for maintenance.
3.  **Settings**:
    - Enable "Auto-Switch Plan" to let WPAS manage power settings automatically when you plug/unplug your charger.
    - Enable "Run at Windows Startup" to keep WPAS running in the tray.
4.  **Tabs**: Navigate through tabs to manage Services, Processes, or clean system files.
5.  **Tray Icon**: Right-click the tray icon to Exit or bring the app to the front.

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
2.  **Run Uninstall Script**: Run `UninstallWPAS.ps1` if available, or manually delete the installation folder and shortcuts.

## Developer Space

For developers looking to extend or modify WPAS:

- **Entry Point**: `WAPS ControlCenter.ps1` is the main script initializing the GUI.
- **Module**: `WPAS.psm1` contains reusable functions for power and battery status.
- **Configuration**: Settings are stored in `config.json` in the install directory.
- **Logs**: Check the `logs/` directory for `WAPS_Debug.log` to troubleshoot issues.
- **Optimization Logic**: `SystemPowerOptimize.ps1` contains the registry tweaks and powercfg commands.

## Testing Space

To test changes:

1.  Open `WAPS ControlCenter.ps1` in PowerShell ISE or VS Code.
2.  Run the script. The console will output debug logs (also saved to file).
3.  Test the "Auto-Switch" feature by physically unplugging/plugging the power adapter (if on a laptop) or simulating the event.
4.  Verify that `ClearStandbyMemory.ps1` compiles the C# code successfully by running it independently.

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 or later
- Administrator privileges are required for some features (e.g., Power Plans, Service Management, Energy Report).

## Disclaimer

This software is provided "as is", without warranty of any kind, express or implied. The developers are not responsible for any damage to hardware or data loss resulting from the use of this tool. Use at your own risk.

## Contributor

**Kamalesh**
