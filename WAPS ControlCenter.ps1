# ============================================================
# WAPS - Windows Power Automation Suite Control Center (GUI Launcher & Dashboard)
# ============================================================

param([switch]$StartMinimized)
 
# Ensure we are running in a path where we can find the module
$scriptPath = $PSScriptRoot
if (-not $scriptPath) { $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition }

# Debug Logging Setup
$logDir = Join-Path $scriptPath "logs"
$debugLog = Join-Path $logDir "WAPS_Debug.log"

# Ensure log directory exists, fallback to TEMP if necessary
try {
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    # Verify write permissions by creating a temp file
    $testFile = Join-Path $logDir "test_write.tmp"
    "test" | Set-Content -Path $testFile -Force -ErrorAction Stop
    Remove-Item -Path $testFile -Force
} catch {
    $logDir = $env:TEMP
    $debugLog = Join-Path $logDir "WAPS_Debug.log"
}

function Write-Log {
    param($Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] $Message" -ForegroundColor Gray
    try {
        Add-Content -Path $debugLog -Value "[$timestamp] $Message" -Force -ErrorAction Stop
        Add-Content -Path $debugLog -Value "[$timestamp] $Message" -Force -ErrorAction Stop -Encoding UTF8
    } catch {
        # Fail silently if logging fails to avoid recursive crashes
    }
}

Write-Log "Starting WAPS Control Center..."
Write-Log "Script Path: $scriptPath"
Write-Log "StartMinimized Switch: $StartMinimized"
Write-Log "StartMinimized Switch passed: $StartMinimized (Ignoring to force visibility)"
Write-Log "Log File Path: $debugLog"

# Import the WPAS module
try {
    $modulePath = Join-Path $scriptPath "WPAS.psm1"
    if (Test-Path $modulePath) {
        Import-Module $modulePath -Force
        Write-Log "Module Loaded: $modulePath"
    } else {
        Write-Log "Warning: WPAS.psm1 not found at $modulePath"
    }
} catch {
    Write-Log "Error importing module: $_"
}

# Load Windows Forms
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
    Write-Log "Windows Forms loaded and Visual Styles enabled."
} catch {
    Write-Log "Fatal Error loading Windows Forms: $_"
    exit
}

# ------------------------------------------------------------
# Load Configuration
# ------------------------------------------------------------
$configFile = Join-Path $scriptPath "config.json"
$script:config = @{
    ACScript = "SystemACMode.ps1"
    BatteryScript = "Windows Power Saver"
    AutoStart = $false
    DarkMode = $false
    InstallPath = $scriptPath
    AutoSwitch = $false
    ShowDiskCleanup = $true
    ShowRestartExplorer = $true
    ShowCheckUpdates = $true
    ShowEventViewer = $true
    ShowEnergyReport = $true
    ShowRestorePoint = $true
    ShowLauncher = $true
    ShowLogs = $true
    ShowSysInfo = $true
    ShowNetwork = $true
    ShowServices = $true
    ShowProcesses = $true
    ShowStartupApps = $true
    ShowTasks = $true
    ShowSystemClean = $true
    ShowHardware = $true
}
if (Test-Path $configFile) {
    try {
        $jsonContent = Get-Content $configFile -Raw -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($jsonContent)) {
            $saved = $jsonContent | ConvertFrom-Json
            if ($saved.ACScript) { $script:config.ACScript = $saved.ACScript }
            if ($saved.BatteryScript) { $script:config.BatteryScript = $saved.BatteryScript }
            if ($saved.AutoStart) { $script:config.AutoStart = $saved.AutoStart }
            if ($saved.DarkMode) { $script:config.DarkMode = $saved.DarkMode }
            if ($saved.InstallPath) { $script:config.InstallPath = $saved.InstallPath }
            if ($saved.PSObject.Properties['AutoSwitch']) { $script:config.AutoSwitch = $saved.AutoSwitch }
            if ($saved.PSObject.Properties['ShowDiskCleanup']) { $script:config.ShowDiskCleanup = $saved.ShowDiskCleanup }
            if ($saved.PSObject.Properties['ShowRestartExplorer']) { $script:config.ShowRestartExplorer = $saved.ShowRestartExplorer }
            if ($saved.PSObject.Properties['ShowCheckUpdates']) { $script:config.ShowCheckUpdates = $saved.ShowCheckUpdates }
            if ($saved.PSObject.Properties['ShowEventViewer']) { $script:config.ShowEventViewer = $saved.ShowEventViewer }
            if ($saved.PSObject.Properties['ShowEnergyReport']) { $script:config.ShowEnergyReport = $saved.ShowEnergyReport }
            if ($saved.PSObject.Properties['ShowRestorePoint']) { $script:config.ShowRestorePoint = $saved.ShowRestorePoint }
            if ($saved.PSObject.Properties['ShowLauncher']) { $script:config.ShowLauncher = $saved.ShowLauncher }
            if ($saved.PSObject.Properties['ShowLogs']) { $script:config.ShowLogs = $saved.ShowLogs }
            if ($saved.PSObject.Properties['ShowSysInfo']) { $script:config.ShowSysInfo = $saved.ShowSysInfo }
            if ($saved.PSObject.Properties['ShowNetwork']) { $script:config.ShowNetwork = $saved.ShowNetwork }
            if ($saved.PSObject.Properties['ShowServices']) { $script:config.ShowServices = $saved.ShowServices }
            if ($saved.PSObject.Properties['ShowProcesses']) { $script:config.ShowProcesses = $saved.ShowProcesses }
            if ($saved.PSObject.Properties['ShowStartupApps']) { $script:config.ShowStartupApps = $saved.ShowStartupApps }
            if ($saved.PSObject.Properties['ShowTasks']) { $script:config.ShowTasks = $saved.ShowTasks }
            if ($saved.PSObject.Properties['ShowSystemClean']) { $script:config.ShowSystemClean = $saved.ShowSystemClean }
            if ($saved.PSObject.Properties['ShowHardware']) { $script:config.ShowHardware = $saved.ShowHardware }
        }
        Write-Log "Configuration loaded successfully."
    } catch {
        # Config is likely corrupt or empty; ignore and use defaults
        Write-Log "Error loading config: $_"
    }
}

# ------------------------------------------------------------
# Main Form Setup
# ------------------------------------------------------------
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Windows Power Automation Suite (WPAS)"
$mainForm.Size = New-Object System.Drawing.Size(600, 750)
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = "FixedSingle"
$mainForm.MaximizeBox = $false
$mainForm.Opacity = 1.0
Write-Log "Main Form Created."

# Set Icon (Try to use PowerShell icon)
try {
    $iconPath = Join-Path $scriptPath "app.ico"
    if (-not (Test-Path $iconPath)) {
        $iconPath = "$PSHOME\powershell.exe"
    }
    $mainForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
} catch {}

# ------------------------------------------------------------
# Tray Icon
# ------------------------------------------------------------
$trayIcon = New-Object System.Windows.Forms.NotifyIcon
$trayIcon.Icon = $mainForm.Icon
$trayIcon.Text = "Windows Power Automation Suite (WPAS)"
$trayIcon.Visible = $true

$trayMenu = New-Object System.Windows.Forms.ContextMenu
$menuItemShow = $trayMenu.MenuItems.Add("Open WPAS")
$menuItemShow.add_Click({
    $mainForm.Show()
    $mainForm.WindowState = "Normal"
    $mainForm.Activate()
})

$menuItemExit = $trayMenu.MenuItems.Add("Exit")
$menuItemExit.add_Click({
    $script:allowExit = $true
    $trayIcon.Visible = $false
    $trayIcon.Dispose()
    $mainForm.Close()
    [System.Windows.Forms.Application]::Exit()
    [Environment]::Exit(0)
})

$trayIcon.ContextMenu = $trayMenu

# Initialize exit flag
$script:allowExit = $false

# Close to Tray Logic
$mainForm.add_FormClosing({
    if (-not $script:allowExit -and $_.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing) {
        $_.Cancel = $true
        $mainForm.Hide()
        $trayIcon.ShowBalloonTip(1000, "WPAS", "App is running in the background.", [System.Windows.Forms.ToolTipIcon]::Info)
        Write-Log "Form closing intercepted. Minimized to tray."
    } else {
        Write-Log "Application exiting."
    }
})

# Handle Startup State
$mainForm.add_Shown({
    Write-Log "Form Shown event triggered."
    if ($StartMinimized) {
        $mainForm.WindowState = "Minimized"
        $mainForm.Hide()
        $timer.Start()
        Write-Log "Starting in Minimized mode (Hidden)."
    } else {
        $mainForm.WindowState = "Normal"
        $mainForm.TopMost = $true
        $mainForm.Show()
        $mainForm.Activate()
        $mainForm.Focus()
        [System.Windows.Forms.Application]::DoEvents()
        $mainForm.TopMost = $false
        $timer.Start()
        Write-Log "Starting in Normal mode."
    }
    $mainForm.WindowState = "Normal"
    $mainForm.Activate()
    $timer.Start()
    Write-Log "Form activated and Timer started."
})

# ------------------------------------------------------------
# Tab Control
# ------------------------------------------------------------
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = "Fill"
$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# ------------------------------------------------------------
# Tab 1: Dashboard
# ------------------------------------------------------------
$tabDashboard = New-Object System.Windows.Forms.TabPage
$tabDashboard.Text = "Dashboard"
$tabDashboard.UseVisualStyleBackColor = $true
$tabDashboard.AutoScroll = $true

# Power Source Label
$lblPowerSource = New-Object System.Windows.Forms.Label
$lblPowerSource.Location = New-Object System.Drawing.Point(20, 20)
$lblPowerSource.Size = New-Object System.Drawing.Size(400, 30)
$lblPowerSource.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$lblPowerSource.Text = "Power Source: Detecting..."

# Active Plan Label
$lblPlan = New-Object System.Windows.Forms.Label
$lblPlan.Location = New-Object System.Drawing.Point(20, 60)
$lblPlan.Size = New-Object System.Drawing.Size(400, 30)
$lblPlan.Text = "Active Plan: ..."

# GPU & CPU Temp Labels
$lblGPU = New-Object System.Windows.Forms.Label
$lblGPU.Location = New-Object System.Drawing.Point(20, 100)
$lblGPU.Size = New-Object System.Drawing.Size(500, 25)
$lblGPU.Text = "GPU: Detecting..."

$lblCPUTemp = New-Object System.Windows.Forms.Label
$lblCPUTemp.Location = New-Object System.Drawing.Point(20, 130)
$lblCPUTemp.Size = New-Object System.Drawing.Size(500, 25)
$lblCPUTemp.Text = "CPU Temp: Detecting..."

$lblUptime = New-Object System.Windows.Forms.Label
$lblUptime.Location = New-Object System.Drawing.Point(20, 160)
$lblUptime.Size = New-Object System.Drawing.Size(500, 25)
$lblUptime.Text = "Uptime: Calculating..."

# RAM Usage Dashboard
$lblRAMDashboard = New-Object System.Windows.Forms.Label
$lblRAMDashboard.Location = New-Object System.Drawing.Point(20, 190)
$lblRAMDashboard.Size = New-Object System.Drawing.Size(500, 25)
$lblRAMDashboard.Text = "RAM Usage: Calculating..."

$pbRAMDashboard = New-Object System.Windows.Forms.ProgressBar
$pbRAMDashboard.Location = New-Object System.Drawing.Point(20, 220)
$pbRAMDashboard.Size = New-Object System.Drawing.Size(520, 20)
$pbRAMDashboard.Style = "Continuous"

# Network Speed Label
$lblNetSpeed = New-Object System.Windows.Forms.Label
$lblNetSpeed.Location = New-Object System.Drawing.Point(20, 250)
$lblNetSpeed.Size = New-Object System.Drawing.Size(500, 25)
$lblNetSpeed.Text = "Network Speed: Calculating..."

# Disk Speed Label
$lblDiskSpeed = New-Object System.Windows.Forms.Label
$lblDiskSpeed.Location = New-Object System.Drawing.Point(20, 280)
$lblDiskSpeed.Size = New-Object System.Drawing.Size(500, 25)
$lblDiskSpeed.Text = "Disk Speed: Calculating..."

# Battery Info Group
$grpBattery = New-Object System.Windows.Forms.GroupBox
$grpBattery.Text = "Battery Health"
$grpBattery.Location = New-Object System.Drawing.Point(20, 315)
$grpBattery.Size = New-Object System.Drawing.Size(520, 160)

$lblBatDesign = New-Object System.Windows.Forms.Label
$lblBatDesign.Location = New-Object System.Drawing.Point(20, 30)
$lblBatDesign.Size = New-Object System.Drawing.Size(400, 25)
$lblBatDesign.Text = "Design Capacity: ..."

$lblBatFull = New-Object System.Windows.Forms.Label
$lblBatFull.Location = New-Object System.Drawing.Point(20, 60)
$lblBatFull.Size = New-Object System.Drawing.Size(400, 25)
$lblBatFull.Text = "Full Charge Capacity: ..."

$lblBatWear = New-Object System.Windows.Forms.Label
$lblBatWear.Location = New-Object System.Drawing.Point(20, 90)
$lblBatWear.Size = New-Object System.Drawing.Size(400, 25)
$lblBatWear.Text = "Wear Level: ..."

$lblBatPercent = New-Object System.Windows.Forms.Label
$lblBatPercent.Location = New-Object System.Drawing.Point(20, 120)
$lblBatPercent.Size = New-Object System.Drawing.Size(400, 25)
$lblBatPercent.Text = "Battery Percentage: ..."

# Auto-Switch Checkbox
$chkAutoSwitch = New-Object System.Windows.Forms.CheckBox
$chkAutoSwitch.Text = "Auto-Switch Plan on Power Change"
$chkAutoSwitch.Location = New-Object System.Drawing.Point(20, 490)
$chkAutoSwitch.Size = New-Object System.Drawing.Size(300, 30)
$chkAutoSwitch.Checked = $script:config.AutoSwitch

# Actions Panel (Flow Layout for dynamic buttons)
$flpActions = New-Object System.Windows.Forms.FlowLayoutPanel
$flpActions.Location = New-Object System.Drawing.Point(20, 530)
$flpActions.Size = New-Object System.Drawing.Size(540, 200)
$flpActions.FlowDirection = "LeftToRight"
$flpActions.WrapContents = $true
$flpActions.AutoSize = $true

# Disk Cleanup Button
$btnDiskCleanup = New-Object System.Windows.Forms.Button
$btnDiskCleanup.Text = "Disk Cleanup"
$btnDiskCleanup.Size = New-Object System.Drawing.Size(150, 35)
$btnDiskCleanup.add_Click({
    Start-Process "cleanmgr.exe" -Verb RunAs
})
$btnDiskCleanup.Visible = $script:config.ShowDiskCleanup

# Check for Updates Button
$btnCheckUpdates = New-Object System.Windows.Forms.Button
$btnCheckUpdates.Text = "Check for Updates"
$btnCheckUpdates.Location = New-Object System.Drawing.Point(20, 580)
$btnCheckUpdates.Size = New-Object System.Drawing.Size(150, 35)
$btnCheckUpdates.add_Click({
    try {
        Start-Process "ms-settings:windowsupdate"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Could not open the Windows Update settings page.", "WPAS Error", "OK", "Error")
    }
})
$btnCheckUpdates.Visible = $script:config.ShowCheckUpdates

# Event Viewer Button
$btnEventViewer = New-Object System.Windows.Forms.Button
$btnEventViewer.Text = "Event Viewer"
$btnEventViewer.Location = New-Object System.Drawing.Point(180, 580)
$btnEventViewer.Size = New-Object System.Drawing.Size(150, 35)
$btnEventViewer.add_Click({
    Start-Process "eventvwr.msc"
})
$btnEventViewer.Visible = $script:config.ShowEventViewer

# Restart Explorer Button
$btnRestartExplorer = New-Object System.Windows.Forms.Button
$btnRestartExplorer.Text = "Restart Explorer"
$btnRestartExplorer.Location = New-Object System.Drawing.Point(180, 530)
$btnRestartExplorer.Size = New-Object System.Drawing.Size(150, 35)
$btnRestartExplorer.add_Click({
    Stop-Process -Name explorer -Force
})
$btnRestartExplorer.Visible = $script:config.ShowRestartExplorer

# Energy Report Button
$btnEnergyReport = New-Object System.Windows.Forms.Button
$btnEnergyReport.Text = "Generate Energy Report"
$btnEnergyReport.Location = New-Object System.Drawing.Point(20, 630)
$btnEnergyReport.Size = New-Object System.Drawing.Size(310, 35)
$btnEnergyReport.add_Click({
    $reportPath = Join-Path $scriptPath "logs\energy-report.html"
    Start-Process "powershell.exe" -ArgumentList "-NoProfile -Command `"Write-Host 'Generating Energy Report (60 seconds)...' -ForegroundColor Cyan; powercfg /energy /output '$reportPath' /duration 60; Write-Host 'Report Generated: $reportPath' -ForegroundColor Green; Start-Process '$reportPath'; Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.MessageBox]::Show('Energy Report Generated and Opened.', 'WPAS', 'OK', 'Information')`"" -Verb RunAs
})
$btnEnergyReport.Visible = $script:config.ShowEnergyReport

# Create Restore Point Button
$btnRestorePoint = New-Object System.Windows.Forms.Button
$btnRestorePoint.Text = "Create Restore Point"
$btnRestorePoint.Location = New-Object System.Drawing.Point(340, 630)
$btnRestorePoint.Size = New-Object System.Drawing.Size(200, 35)
$btnRestorePoint.add_Click({
    $desc = "WPAS Manual Restore Point $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    $cmd = "try { Checkpoint-Computer -Description '$desc' -RestorePointType 'MODIFY_SETTINGS' -ErrorAction Stop; Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.MessageBox]::Show('Restore Point Created Successfully.', 'WPAS', 'OK', 'Information') } catch { Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.MessageBox]::Show('Error creating restore point: ' + `$_.Exception.Message, 'WPAS Error', 'OK', 'Error') }"
    Start-Process "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"$cmd`"" -Verb RunAs
})
$btnRestorePoint.Visible = $script:config.ShowRestorePoint

# Battery Report Button
$btnBatteryReport = New-Object System.Windows.Forms.Button
$btnBatteryReport.Text = "View Battery Report"
$btnBatteryReport.Size = New-Object System.Drawing.Size(150, 35)
$btnBatteryReport.add_Click({
    $script = Join-Path $scriptPath "GenerateBatteryReport.ps1"
    if (Test-Path $script) {
        Start-Process "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$script`""
    } else {
        [System.Windows.Forms.MessageBox]::Show("Script not found: GenerateBatteryReport.ps1", "WPAS Error", "OK", "Error")
    }
})

$flpActions.Controls.AddRange(@($btnDiskCleanup, $btnRestartExplorer, $btnCheckUpdates, $btnEventViewer, $btnEnergyReport, $btnRestorePoint, $btnBatteryReport))
$grpBattery.Controls.AddRange(@($lblBatDesign, $lblBatFull, $lblBatWear, $lblBatPercent))
$tabDashboard.Controls.AddRange(@($lblPowerSource, $lblPlan, $lblGPU, $lblCPUTemp, $lblUptime, $lblRAMDashboard, $pbRAMDashboard, $lblNetSpeed, $lblDiskSpeed, $grpBattery, $chkAutoSwitch, $flpActions))

# ------------------------------------------------------------
# Tab 2: Launcher
# ------------------------------------------------------------
$tabLauncher = New-Object System.Windows.Forms.TabPage
$tabLauncher.Text = "Launcher"
$tabLauncher.UseVisualStyleBackColor = $true

function Add-ScriptButton ($text, $y, $fileName) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text
    $btn.Location = New-Object System.Drawing.Point(40, $y)
    $btn.Size = New-Object System.Drawing.Size(480, 40)
    $btn.BackColor = [System.Drawing.Color]::WhiteSmoke
    $btn.add_Click({
        $fullPath = Join-Path $scriptPath $fileName
        if (Test-Path $fullPath) {
            Start-Process "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$fullPath`"" -Verb RunAs
        } else {
            [System.Windows.Forms.MessageBox]::Show("Script not found: $fileName", "WPAS Error", "OK", "Error")
        }
    })
    $tabLauncher.Controls.Add($btn)
}

Add-ScriptButton "Run Cleanup Logs" 30 "CleanupLogs.ps1"
Add-ScriptButton "Enable Performance Mode" 80 "PerformanceMode.ps1"
Add-ScriptButton "Enable System AC Mode" 130 "SystemACMode.ps1"
Add-ScriptButton "Run System Optimization" 180 "SystemPowerOptimize.ps1"
Add-ScriptButton "Check Battery Health (Console)" 230 "BatteryHealth.ps1"
Add-ScriptButton "Open Console Dashboard" 280 "Dashboard.ps1"

# ------------------------------------------------------------
# Tab 3: Settings
# ------------------------------------------------------------
$tabSettings = New-Object System.Windows.Forms.TabPage
$tabSettings.Text = "Settings"
$tabSettings.UseVisualStyleBackColor = $true
$tabSettings.AutoScroll = $true

# AC Script Selection
$lblACSet = New-Object System.Windows.Forms.Label
$lblACSet.Text = "Script to run on AC Power:"
$lblACSet.Location = New-Object System.Drawing.Point(20, 20)
$lblACSet.Size = New-Object System.Drawing.Size(300, 25)

$cmbAC = New-Object System.Windows.Forms.ComboBox
$cmbAC.Location = New-Object System.Drawing.Point(20, 50)
$cmbAC.Size = New-Object System.Drawing.Size(400, 30)
$cmbAC.DropDownStyle = "DropDownList"

# Battery Script Selection
$lblBatSet = New-Object System.Windows.Forms.Label
$lblBatSet.Text = "Script to run on Battery Power:"
$lblBatSet.Location = New-Object System.Drawing.Point(20, 100)
$lblBatSet.Size = New-Object System.Drawing.Size(300, 25)

$cmbBat = New-Object System.Windows.Forms.ComboBox
$cmbBat.Location = New-Object System.Drawing.Point(20, 130)
$cmbBat.Size = New-Object System.Drawing.Size(400, 30)
$cmbBat.DropDownStyle = "DropDownList"

# Populate Scripts
$availScripts = Get-ChildItem $scriptPath -Filter "*.ps1" | Select-Object -ExpandProperty Name
$cmbAC.Items.AddRange($availScripts)
$cmbBat.Items.Add("Windows Power Saver")
$cmbBat.Items.AddRange($availScripts)

# Set Current Selections
if ($cmbAC.Items.Contains($script:config.ACScript)) { $cmbAC.SelectedItem = $script:config.ACScript }
if ($cmbBat.Items.Contains($script:config.BatteryScript)) { $cmbBat.SelectedItem = $script:config.BatteryScript }

# Startup Checkbox
$chkStartupSet = New-Object System.Windows.Forms.CheckBox
$chkStartupSet.Text = "Run Control Center at Windows Startup"
$chkStartupSet.Location = New-Object System.Drawing.Point(20, 180)
$chkStartupSet.Size = New-Object System.Drawing.Size(400, 30)
$chkStartupSet.Checked = $script:config.AutoStart

# Dark Mode Checkbox
$chkDarkMode = New-Object System.Windows.Forms.CheckBox
$chkDarkMode.Text = "Enable Dark Mode"
$chkDarkMode.Location = New-Object System.Drawing.Point(20, 220)
$chkDarkMode.Size = New-Object System.Drawing.Size(400, 30)
$chkDarkMode.Checked = $script:config.DarkMode

# Dashboard Visibility Group
$grpDashVis = New-Object System.Windows.Forms.GroupBox
$grpDashVis.Text = "Dashboard Buttons"
$grpDashVis.Location = New-Object System.Drawing.Point(20, 260)
$grpDashVis.Size = New-Object System.Drawing.Size(540, 150)

$chkShowDisk = New-Object System.Windows.Forms.CheckBox
$chkShowDisk.Text = "Disk Cleanup"
$chkShowDisk.Location = New-Object System.Drawing.Point(20, 30)
$chkShowDisk.Checked = $script:config.ShowDiskCleanup

$chkShowExplorer = New-Object System.Windows.Forms.CheckBox
$chkShowExplorer.Text = "Restart Explorer"
$chkShowExplorer.Location = New-Object System.Drawing.Point(200, 30)
$chkShowExplorer.Checked = $script:config.ShowRestartExplorer

$chkShowUpdates = New-Object System.Windows.Forms.CheckBox
$chkShowUpdates.Text = "Check Updates"
$chkShowUpdates.Location = New-Object System.Drawing.Point(380, 30)
$chkShowUpdates.Checked = $script:config.ShowCheckUpdates

$chkShowEvent = New-Object System.Windows.Forms.CheckBox
$chkShowEvent.Text = "Event Viewer"
$chkShowEvent.Location = New-Object System.Drawing.Point(20, 70)
$chkShowEvent.Checked = $script:config.ShowEventViewer

$chkShowEnergy = New-Object System.Windows.Forms.CheckBox
$chkShowEnergy.Text = "Energy Report"
$chkShowEnergy.Location = New-Object System.Drawing.Point(200, 70)
$chkShowEnergy.Checked = $script:config.ShowEnergyReport

$chkShowRestore = New-Object System.Windows.Forms.CheckBox
$chkShowRestore.Text = "Restore Point"
$chkShowRestore.Location = New-Object System.Drawing.Point(380, 70)
$chkShowRestore.Checked = $script:config.ShowRestorePoint

$grpDashVis.Controls.AddRange(@($chkShowDisk, $chkShowExplorer, $chkShowUpdates, $chkShowEvent, $chkShowEnergy, $chkShowRestore))

# Tab Visibility Group
$grpTabVis = New-Object System.Windows.Forms.GroupBox
$grpTabVis.Text = "Tab Visibility"
$grpTabVis.Location = New-Object System.Drawing.Point(20, 420)
$grpTabVis.Size = New-Object System.Drawing.Size(540, 160)

$chkShowLauncher = New-Object System.Windows.Forms.CheckBox; $chkShowLauncher.Text = "Launcher"; $chkShowLauncher.Location = New-Object System.Drawing.Point(20, 30); $chkShowLauncher.Checked = $script:config.ShowLauncher
$chkShowLogs = New-Object System.Windows.Forms.CheckBox; $chkShowLogs.Text = "Logs"; $chkShowLogs.Location = New-Object System.Drawing.Point(200, 30); $chkShowLogs.Checked = $script:config.ShowLogs
$chkShowSysInfo = New-Object System.Windows.Forms.CheckBox; $chkShowSysInfo.Text = "System Info"; $chkShowSysInfo.Location = New-Object System.Drawing.Point(380, 30); $chkShowSysInfo.Checked = $script:config.ShowSysInfo

$chkShowNetwork = New-Object System.Windows.Forms.CheckBox; $chkShowNetwork.Text = "Network"; $chkShowNetwork.Location = New-Object System.Drawing.Point(20, 60); $chkShowNetwork.Checked = $script:config.ShowNetwork
$chkShowServices = New-Object System.Windows.Forms.CheckBox; $chkShowServices.Text = "Services"; $chkShowServices.Location = New-Object System.Drawing.Point(200, 60); $chkShowServices.Checked = $script:config.ShowServices
$chkShowProcesses = New-Object System.Windows.Forms.CheckBox; $chkShowProcesses.Text = "Processes"; $chkShowProcesses.Location = New-Object System.Drawing.Point(380, 60); $chkShowProcesses.Checked = $script:config.ShowProcesses

$chkShowStartup = New-Object System.Windows.Forms.CheckBox; $chkShowStartup.Text = "Startup Apps"; $chkShowStartup.Location = New-Object System.Drawing.Point(20, 90); $chkShowStartup.Checked = $script:config.ShowStartupApps
$chkShowTasks = New-Object System.Windows.Forms.CheckBox; $chkShowTasks.Text = "Scheduled Tasks"; $chkShowTasks.Location = New-Object System.Drawing.Point(200, 90); $chkShowTasks.Checked = $script:config.ShowTasks
$chkShowClean = New-Object System.Windows.Forms.CheckBox; $chkShowClean.Text = "System Clean"; $chkShowClean.Location = New-Object System.Drawing.Point(380, 90); $chkShowClean.Checked = $script:config.ShowSystemClean

$chkShowHardware = New-Object System.Windows.Forms.CheckBox; $chkShowHardware.Text = "Hardware Specs"; $chkShowHardware.Location = New-Object System.Drawing.Point(20, 120); $chkShowHardware.Checked = $script:config.ShowHardware

$grpTabVis.Controls.AddRange(@($chkShowLauncher, $chkShowLogs, $chkShowSysInfo, $chkShowNetwork, $chkShowServices, $chkShowProcesses, $chkShowStartup, $chkShowTasks, $chkShowClean, $chkShowHardware))

# Save Button
$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Save Settings"
$btnSave.Location = New-Object System.Drawing.Point(20, 600)
$btnSave.Size = New-Object System.Drawing.Size(150, 40)
$btnSave.add_Click({
    $script:config.ACScript = $cmbAC.SelectedItem
    $script:config.BatteryScript = $cmbBat.SelectedItem
    $script:config.AutoStart = $chkStartupSet.Checked
    $script:config.DarkMode = $chkDarkMode.Checked
    $script:config.AutoSwitch = $chkAutoSwitch.Checked
    
    $script:config.ShowDiskCleanup = $chkShowDisk.Checked
    $script:config.ShowRestartExplorer = $chkShowExplorer.Checked
    $script:config.ShowCheckUpdates = $chkShowUpdates.Checked
    $script:config.ShowEventViewer = $chkShowEvent.Checked
    $script:config.ShowEnergyReport = $chkShowEnergy.Checked
    $script:config.ShowRestorePoint = $chkShowRestore.Checked

    $script:config.ShowLauncher = $chkShowLauncher.Checked
    $script:config.ShowLogs = $chkShowLogs.Checked
    $script:config.ShowSysInfo = $chkShowSysInfo.Checked
    $script:config.ShowNetwork = $chkShowNetwork.Checked
    $script:config.ShowServices = $chkShowServices.Checked
    $script:config.ShowProcesses = $chkShowProcesses.Checked
    $script:config.ShowStartupApps = $chkShowStartup.Checked
    $script:config.ShowTasks = $chkShowTasks.Checked
    $script:config.ShowSystemClean = $chkShowClean.Checked
    $script:config.ShowHardware = $chkShowHardware.Checked

    $script:config | ConvertTo-Json | Set-Content $configFile

    # Handle Startup Shortcut
    $startupPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\Windows Power Automation Suite.lnk"
    if ($chkStartupSet.Checked) {
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($startupPath)
        $Shortcut.TargetPath = "powershell.exe"
        $Shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath\WAPS ControlCenter.ps1`" -StartMinimized"
        $Shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath\WAPS ControlCenter.ps1`""
        $Shortcut.IconLocation = "$PSHOME\powershell.exe"
        $Shortcut.Description = "Windows Power Automation Suite"
        $Shortcut.WorkingDirectory = $scriptPath
        $Shortcut.Save()
    } else {
        if (Test-Path $startupPath) { Remove-Item $startupPath }
    }

    # Apply Visibility Immediately
    $btnDiskCleanup.Visible = $script:config.ShowDiskCleanup
    $btnRestartExplorer.Visible = $script:config.ShowRestartExplorer
    $btnCheckUpdates.Visible = $script:config.ShowCheckUpdates
    $btnEventViewer.Visible = $script:config.ShowEventViewer
    $btnEnergyReport.Visible = $script:config.ShowEnergyReport
    $btnRestorePoint.Visible = $script:config.ShowRestorePoint

    Refresh-Tabs

    [System.Windows.Forms.MessageBox]::Show("Settings saved successfully.", "WPAS", "OK", "Information")
})

# Reset Button
$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Text = "Reset Defaults"
$btnReset.Location = New-Object System.Drawing.Point(180, 600)
$btnReset.Size = New-Object System.Drawing.Size(150, 40)
$btnReset.add_Click({
    $script:config.ACScript = "SystemACMode.ps1"
    $script:config.BatteryScript = "Windows Power Saver"
    $script:config.AutoStart = $false
    $script:config.DarkMode = $false
    $script:config.AutoSwitch = $false
    $script:config.ShowDiskCleanup = $true
    $script:config.ShowRestartExplorer = $true
    $script:config.ShowCheckUpdates = $true
    $script:config.ShowEventViewer = $true
    $script:config.ShowEnergyReport = $true
    $script:config.ShowRestorePoint = $true
    $script:config.ShowLauncher = $true
    $script:config.ShowLogs = $true
    $script:config.ShowSysInfo = $true
    $script:config.ShowNetwork = $true
    $script:config.ShowServices = $true
    $script:config.ShowProcesses = $true
    $script:config.ShowStartupApps = $true
    $script:config.ShowTasks = $true
    $script:config.ShowSystemClean = $true
    $script:config.ShowHardware = $true

    if ($cmbAC.Items.Contains($script:config.ACScript)) { $cmbAC.SelectedItem = $script:config.ACScript }
    if ($cmbBat.Items.Contains($script:config.BatteryScript)) { $cmbBat.SelectedItem = $script:config.BatteryScript }
    $chkStartupSet.Checked = $script:config.AutoStart
    $chkDarkMode.Checked = $script:config.DarkMode
    $chkAutoSwitch.Checked = $script:config.AutoSwitch
    
    $chkShowDisk.Checked = $true
    $chkShowExplorer.Checked = $true
    $chkShowUpdates.Checked = $true
    $chkShowEvent.Checked = $true
    $chkShowEnergy.Checked = $true
    $chkShowRestore.Checked = $true

    $chkShowLauncher.Checked = $true
    $chkShowLogs.Checked = $true
    $chkShowSysInfo.Checked = $true
    $chkShowNetwork.Checked = $true
    $chkShowServices.Checked = $true
    $chkShowProcesses.Checked = $true
    $chkShowStartup.Checked = $true
    $chkShowTasks.Checked = $true
    $chkShowClean.Checked = $true
    $chkShowHardware.Checked = $true

    $script:config | ConvertTo-Json | Set-Content $configFile

    $startupPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\Windows Power Automation Suite.lnk"
    if (Test-Path $startupPath) { Remove-Item $startupPath }

    Set-Theme $script:config.DarkMode
    
    # Apply Visibility Immediately
    $btnDiskCleanup.Visible = $true
    $btnRestartExplorer.Visible = $true
    $btnCheckUpdates.Visible = $true
    $btnEventViewer.Visible = $true
    $btnEnergyReport.Visible = $true
    $btnRestorePoint.Visible = $true

    Refresh-Tabs

    [System.Windows.Forms.MessageBox]::Show("Configuration reset to defaults.", "WPAS", "OK", "Information")
})

$tabSettings.Controls.AddRange(@($lblACSet, $cmbAC, $lblBatSet, $cmbBat, $chkStartupSet, $chkDarkMode, $grpDashVis, $grpTabVis, $btnSave, $btnReset))

# ------------------------------------------------------------
# Tab 4: Logs
# ------------------------------------------------------------
$tabLogs = New-Object System.Windows.Forms.TabPage
$tabLogs.Text = "Logs"
$tabLogs.UseVisualStyleBackColor = $true

$lblLogSelect = New-Object System.Windows.Forms.Label
$lblLogSelect.Text = "Select Log File:"
$lblLogSelect.Location = New-Object System.Drawing.Point(20, 20)
$lblLogSelect.Size = New-Object System.Drawing.Size(150, 25)

$cmbLogs = New-Object System.Windows.Forms.ComboBox
$cmbLogs.Location = New-Object System.Drawing.Point(180, 17)
$cmbLogs.Size = New-Object System.Drawing.Size(300, 30)
$cmbLogs.DropDownStyle = "DropDownList"

$btnRefreshLogs = New-Object System.Windows.Forms.Button
$btnRefreshLogs.Text = "Refresh"
$btnRefreshLogs.Location = New-Object System.Drawing.Point(490, 16)
$btnRefreshLogs.Size = New-Object System.Drawing.Size(80, 30)

$txtLogContent = New-Object System.Windows.Forms.TextBox
$txtLogContent.Location = New-Object System.Drawing.Point(20, 60)
$txtLogContent.Size = New-Object System.Drawing.Size(550, 380)
$txtLogContent.Multiline = $true
$txtLogContent.ScrollBars = "Vertical"
$txtLogContent.ReadOnly = $true
$txtLogContent.Font = New-Object System.Drawing.Font("Consolas", 9)

function Load-LogFiles {
    $cmbLogs.Items.Clear()
    $logDir = Join-Path $scriptPath "logs"
    if (Test-Path $logDir) {
        $logs = Get-ChildItem $logDir -Filter "*.log"
        foreach ($log in $logs) {
            $cmbLogs.Items.Add($log.Name)
        }
    }
    if ($cmbLogs.Items.Count -gt 0) { $cmbLogs.SelectedIndex = 0 }
}

$cmbLogs.add_SelectedIndexChanged({
    $logDir = Join-Path $scriptPath "logs"
    $selectedLog = Join-Path $logDir $cmbLogs.SelectedItem
    if (Test-Path $selectedLog) {
        $txtLogContent.Text = Get-Content $selectedLog -Raw
        $txtLogContent.SelectionStart = $txtLogContent.Text.Length
        $txtLogContent.ScrollToCaret()
    } else {
        $txtLogContent.Text = "Log file not found."
    }
})

$btnRefreshLogs.add_Click({ Load-LogFiles })
try { Load-LogFiles } catch {}

$tabLogs.Controls.AddRange(@($lblLogSelect, $cmbLogs, $btnRefreshLogs, $txtLogContent))

# ------------------------------------------------------------
# Tab 5: System Info
# ------------------------------------------------------------
$tabSysInfo = New-Object System.Windows.Forms.TabPage
$tabSysInfo.Text = "System Info"
$tabSysInfo.UseVisualStyleBackColor = $true

$lblSysCPU = New-Object System.Windows.Forms.Label
$lblSysCPU.Location = New-Object System.Drawing.Point(20, 30)
$lblSysCPU.Size = New-Object System.Drawing.Size(400, 30)
$lblSysCPU.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$lblSysCPU.Text = "CPU Usage: Calculating..."

$lblSysRAM = New-Object System.Windows.Forms.Label
$lblSysRAM.Location = New-Object System.Drawing.Point(20, 80)
$lblSysRAM.Size = New-Object System.Drawing.Size(400, 30)
$lblSysRAM.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$lblSysRAM.Text = "RAM Usage: Calculating..."

$lblSysDisk = New-Object System.Windows.Forms.Label
$lblSysDisk.Location = New-Object System.Drawing.Point(20, 130)
$lblSysDisk.Size = New-Object System.Drawing.Size(400, 30)
$lblSysDisk.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$lblSysDisk.Text = "Disk (C:) Usage: Calculating..."

$btnExportSysInfo = New-Object System.Windows.Forms.Button
$btnExportSysInfo.Text = "Export Info"
$btnExportSysInfo.Location = New-Object System.Drawing.Point(20, 180)
$btnExportSysInfo.Size = New-Object System.Drawing.Size(150, 40)
$btnExportSysInfo.add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "Text Files (*.txt)|*.txt"
    $sfd.FileName = "SystemInfo_$(Get-Date -Format 'yyyyMMdd_HHmm').txt"
    if ($sfd.ShowDialog() -eq "OK") {
        $content = "PCFix System Info Export`r`n"
        $content += "Date: $(Get-Date)`r`n"
        $content += "$($lblSysCPU.Text)`r`n"
        $content += "$($lblSysRAM.Text)`r`n"
        $content += "$($lblSysDisk.Text)`r`n"
        Set-Content -Path $sfd.FileName -Value $content
        [System.Windows.Forms.MessageBox]::Show("System info exported.", "WPAS Success", "OK", "Information")
    }
})

$tabSysInfo.Controls.AddRange(@($lblSysCPU, $lblSysRAM, $lblSysDisk, $btnExportSysInfo))

$tabNetwork = New-Object System.Windows.Forms.TabPage
$tabNetwork.Text = "Network"
$tabNetwork.UseVisualStyleBackColor = $true

$lblNetStatus = New-Object System.Windows.Forms.Label
$lblNetStatus.Location = New-Object System.Drawing.Point(20, 30)
$lblNetStatus.Size = New-Object System.Drawing.Size(400, 30)
$lblNetStatus.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$lblNetStatus.Text = "Internet Status: Checking..."

$lblNetIP = New-Object System.Windows.Forms.Label
$lblNetIP.Location = New-Object System.Drawing.Point(20, 80)
$lblNetIP.Size = New-Object System.Drawing.Size(400, 30)
$lblNetIP.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$lblNetIP.Text = "IP Address: Checking..."

$btnFlushDNS = New-Object System.Windows.Forms.Button
$btnFlushDNS.Text = "Flush DNS"
$btnFlushDNS.Location = New-Object System.Drawing.Point(20, 130)
$btnFlushDNS.Size = New-Object System.Drawing.Size(150, 35)
$btnFlushDNS.add_Click({
    $output = ipconfig /flushdns
    if ($output -match "Successfully") {
        [System.Windows.Forms.MessageBox]::Show("DNS Resolver Cache Successfully Flushed.", "WPAS Network", "OK", "Information")
    } else {
        [System.Windows.Forms.MessageBox]::Show("Failed to flush DNS cache.", "WPAS Error", "OK", "Error")
    }
})

$tabNetwork.Controls.AddRange(@($lblNetStatus, $lblNetIP, $btnFlushDNS))

# ------------------------------------------------------------
# Tab 7: Services
# ------------------------------------------------------------
$tabServices = New-Object System.Windows.Forms.TabPage
$tabServices.Text = "Services"
$tabServices.UseVisualStyleBackColor = $true

$lblSvcSelect = New-Object System.Windows.Forms.Label
$lblSvcSelect.Text = "Select Service:"
$lblSvcSelect.Location = New-Object System.Drawing.Point(20, 20)
$lblSvcSelect.Size = New-Object System.Drawing.Size(150, 25)

$cmbServices = New-Object System.Windows.Forms.ComboBox
$cmbServices.Location = New-Object System.Drawing.Point(20, 50)
$cmbServices.Size = New-Object System.Drawing.Size(300, 30)
$cmbServices.DropDownStyle = "DropDownList"
$cmbServices.Items.AddRange(@("wuauserv (Windows Update)", "Spooler (Print Spooler)", "SysMain (Superfetch)", "Themes", "AudioSrv (Windows Audio)"))

$lblSvcStatus = New-Object System.Windows.Forms.Label
$lblSvcStatus.Location = New-Object System.Drawing.Point(20, 90)
$lblSvcStatus.Size = New-Object System.Drawing.Size(400, 30)
$lblSvcStatus.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$lblSvcStatus.Text = "Status: Select a service"

$btnStartSvc = New-Object System.Windows.Forms.Button
$btnStartSvc.Text = "Start"
$btnStartSvc.Location = New-Object System.Drawing.Point(20, 140)
$btnStartSvc.Size = New-Object System.Drawing.Size(100, 35)
$btnStartSvc.add_Click({
    if ($cmbServices.SelectedItem) {
        $svcName = $cmbServices.SelectedItem.Split(" ")[0]
        Start-Service -Name $svcName -ErrorAction SilentlyContinue
    }
})

$btnStopSvc = New-Object System.Windows.Forms.Button
$btnStopSvc.Text = "Stop"
$btnStopSvc.Location = New-Object System.Drawing.Point(130, 140)
$btnStopSvc.Size = New-Object System.Drawing.Size(100, 35)
$btnStopSvc.add_Click({
    if ($cmbServices.SelectedItem) {
        $svcName = $cmbServices.SelectedItem.Split(" ")[0]
        Stop-Service -Name $svcName -ErrorAction SilentlyContinue
    }
})

$tabServices.Controls.AddRange(@($lblSvcSelect, $cmbServices, $lblSvcStatus, $btnStartSvc, $btnStopSvc))

# ------------------------------------------------------------
# Tab 8: Processes
# ------------------------------------------------------------
$tabProcesses = New-Object System.Windows.Forms.TabPage
$tabProcesses.Text = "Processes"
$tabProcesses.UseVisualStyleBackColor = $true

$lvProcesses = New-Object System.Windows.Forms.ListView
$lvProcesses.Dock = "Fill"
$lvProcesses.View = "Details"
$lvProcesses.FullRowSelect = $true
$lvProcesses.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$lvProcesses.Columns.Add("Process Name", 220) | Out-Null
$lvProcesses.Columns.Add("ID", 70) | Out-Null
$lvProcesses.Columns.Add("CPU %", 70) | Out-Null
$lvProcesses.Columns.Add("Memory (MB)", 100) | Out-Null

# Create Context Menu for Processes
$processContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$terminateMenuItem = $processContextMenu.Items.Add("Terminate Process")
$terminateMenuItem.add_Click({
    if ($lvProcesses.SelectedItems.Count -gt 0) {
        $selectedItem = $lvProcesses.SelectedItems[0]
        $procId = $selectedItem.Tag
        $procName = $selectedItem.SubItems[0].Text
        
        $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to terminate '$procName' (ID: $procId)?", "Confirm Termination", "YesNo", "Warning")
        if ($confirm -eq "Yes") {
            Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
        }
    }
})
$lvProcesses.ContextMenuStrip = $processContextMenu

$tabProcesses.Controls.Add($lvProcesses)

# ------------------------------------------------------------
# Tab 9: Startup Apps
# ------------------------------------------------------------
$tabStartupApps = New-Object System.Windows.Forms.TabPage
$tabStartupApps.Text = "Startup Apps"
$tabStartupApps.UseVisualStyleBackColor = $true

# Panel for Add Button (Dock Bottom)
$pnlStartupBottom = New-Object System.Windows.Forms.Panel
$pnlStartupBottom.Dock = "Bottom"
$pnlStartupBottom.Height = 60

$btnAddStartup = New-Object System.Windows.Forms.Button
$btnAddStartup.Text = "Add Startup App"
$btnAddStartup.Location = New-Object System.Drawing.Point(20, 15)
$btnAddStartup.Size = New-Object System.Drawing.Size(150, 30)
$btnAddStartup.add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Programs (*.exe)|*.exe|All Files (*.*)|*.*"
    $ofd.Title = "Select a program to start automatically"
    if ($ofd.ShowDialog() -eq "OK") {
        $path = $ofd.FileName
        $name = [System.IO.Path]::GetFileNameWithoutExtension($path)
        $startupDir = [System.Environment]::GetFolderPath('Startup')
        $linkPath = Join-Path $startupDir "$name.lnk"
        
        try {
            $wsh = New-Object -ComObject WScript.Shell
            $sc = $wsh.CreateShortcut($linkPath)
            $sc.TargetPath = $path
            $sc.Save()
            Load-StartupApps
            [System.Windows.Forms.MessageBox]::Show("Successfully added '$name' to startup.", "WPAS", "OK", "Information")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Could not create startup shortcut.", "WPAS Error", "OK", "Error")
        }
    }
})

$pnlStartupBottom.Controls.Add($btnAddStartup)

$lvStartupApps = New-Object System.Windows.Forms.ListView
$lvStartupApps.Dock = "Fill"
$lvStartupApps.View = "Details"
$lvStartupApps.FullRowSelect = $true
$lvStartupApps.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$lvStartupApps.Columns.Add("Name", 200) | Out-Null
$lvStartupApps.Columns.Add("Status", 80) | Out-Null
$lvStartupApps.Columns.Add("Location", 120) | Out-Null
$lvStartupApps.Columns.Add("Command", 300) | Out-Null

function Get-StartupApps {
    $startupItems = @()
    $regPaths = @{
        "HKCU Run" = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run";
        "HKLM Run" = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run";
    }

    foreach ($loc in $regPaths.Keys) {
        $path = $regPaths[$loc]
        # Enabled
        if (Test-Path $path) {
            $props = Get-ItemProperty -Path $path
            $props | Get-Member -MemberType NoteProperty | Where-Object Name -ne '(default)' | ForEach-Object {
                $startupItems += [PSCustomObject]@{ Name = $_.Name; Command = $props.$($_.Name); Location = $loc; Status = "Enabled"; Type = "Registry"; FullPath = $path }
            }
        }
        # Disabled
        $disabledPath = $path + "_Disabled"
        if (Test-Path $disabledPath) {
            $props = Get-ItemProperty -Path $disabledPath
            $props | Get-Member -MemberType NoteProperty | Where-Object Name -ne '(default)' | ForEach-Object {
                $startupItems += [PSCustomObject]@{ Name = $_.Name; Command = $props.$($_.Name); Location = $loc; Status = "Disabled"; Type = "Registry"; FullPath = $disabledPath }
            }
        }
    }

    $folderPaths = @{ "User Startup" = [System.Environment]::GetFolderPath('Startup'); "Common Startup" = [System.Environment]::GetFolderPath('CommonStartup') }
    foreach ($loc in $folderPaths.Keys) {
        if (Test-Path $folderPaths[$loc]) {
            Get-ChildItem -Path $folderPaths[$loc] -File | ForEach-Object {
                $status = if ($_.Extension -eq '.disabled') { "Disabled" } else { "Enabled" }
                $startupItems += [PSCustomObject]@{ Name = $_.BaseName.Replace(".lnk", ""); Command = $_.FullName; Location = $loc; Status = $status; Type = "File"; FullPath = $_.FullName }
            }
        }
    }
    return $startupItems | Sort-Object Name
}

function Load-StartupApps {
    $lvStartupApps.BeginUpdate()
    $lvStartupApps.Items.Clear()
    $apps = Get-StartupApps
    foreach ($app in $apps) {
        $item = New-Object System.Windows.Forms.ListViewItem($app.Name)
        $item.SubItems.Add($app.Status) | Out-Null
        $item.SubItems.Add($app.Location) | Out-Null
        $item.SubItems.Add($app.Command) | Out-Null
        $item.Tag = $app
        if ($app.Status -eq "Disabled") { $item.ForeColor = [System.Drawing.Color]::Gray }
        $lvStartupApps.Items.Add($item) | Out-Null
    }
    $lvStartupApps.EndUpdate()
}

$startupContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$startupEnable = $startupContextMenu.Items.Add("Enable")
$startupDisable = $startupContextMenu.Items.Add("Disable")
$startupDelete = $startupContextMenu.Items.Add("Delete")

$startupEnable.add_Click({
    $item = $lvStartupApps.SelectedItems[0].Tag
    if ($item.Status -eq "Disabled") {
        if ($item.Type -eq "Registry") {
            $destPath = $item.FullPath -replace "_Disabled$"
            if (-not (Test-Path $destPath)) { New-Item -Path $destPath -Force | Out-Null }
            Move-ItemProperty -Path $item.FullPath -Name $item.Name -Destination $destPath -Force
        } elseif ($item.Type -eq "File") {
            Rename-Item -Path $item.FullPath -NewName ($item.FullPath -replace '\.disabled$', '') -Force
        }
        Load-StartupApps
    }
})

$startupDisable.add_Click({
    $item = $lvStartupApps.SelectedItems[0].Tag
    if ($item.Status -eq "Enabled") {
        if ($item.Type -eq "Registry") {
            $destPath = $item.FullPath + "_Disabled"
            if (-not (Test-Path $destPath)) { New-Item -Path $destPath -Force | Out-Null }
            Move-ItemProperty -Path $item.FullPath -Name $item.Name -Destination $destPath -Force
        } elseif ($item.Type -eq "File") {
            Rename-Item -Path $item.FullPath -NewName ($item.FullPath + ".disabled") -Force
        }
        Load-StartupApps
    }
})

$startupDelete.add_Click({
    $item = $lvStartupApps.SelectedItems[0].Tag
    if ([System.Windows.Forms.MessageBox]::Show("Permanently delete '$($item.Name)'?", "Confirm Delete", "YesNo", "Warning") -eq "Yes") {
        if ($item.Type -eq "Registry") { Remove-ItemProperty -Path $item.FullPath -Name $item.Name -Force }
        elseif ($item.Type -eq "File") { Remove-Item -Path $item.FullPath -Force }
        Load-StartupApps
    }
})

$lvStartupApps.ContextMenuStrip = $startupContextMenu
$tabStartupApps.Controls.Add($pnlStartupBottom)
$tabStartupApps.Controls.Add($lvStartupApps)

# ------------------------------------------------------------
# Tab 10: Scheduled Tasks
# ------------------------------------------------------------
$tabTasks = New-Object System.Windows.Forms.TabPage
$tabTasks.Text = "Scheduled Tasks"
$tabTasks.UseVisualStyleBackColor = $true

$pnlTasksTop = New-Object System.Windows.Forms.Panel
$pnlTasksTop.Dock = "Top"
$pnlTasksTop.Height = 50

$btnRefreshTasks = New-Object System.Windows.Forms.Button
$btnRefreshTasks.Text = "Refresh List"
$btnRefreshTasks.Location = New-Object System.Drawing.Point(20, 10)
$btnRefreshTasks.Size = New-Object System.Drawing.Size(100, 30)
$btnRefreshTasks.add_Click({ Load-ScheduledTasks })

$btnRunTask = New-Object System.Windows.Forms.Button
$btnRunTask.Text = "Run Selected"
$btnRunTask.Location = New-Object System.Drawing.Point(130, 10)
$btnRunTask.Size = New-Object System.Drawing.Size(120, 30)
$btnRunTask.add_Click({
    if ($lvTasks.SelectedItems.Count -gt 0) {
        $task = $lvTasks.SelectedItems[0].Tag
        Start-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath
        [System.Windows.Forms.MessageBox]::Show("Task triggered successfully.", "WPAS", "OK", "Information")
    }
})

$pnlTasksTop.Controls.AddRange(@($btnRefreshTasks, $btnRunTask))

$lvTasks = New-Object System.Windows.Forms.ListView
$lvTasks.Dock = "Fill"
$lvTasks.View = "Details"
$lvTasks.FullRowSelect = $true
$lvTasks.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lvTasks.Columns.Add("Task Name", 250) | Out-Null
$lvTasks.Columns.Add("State", 100) | Out-Null
$lvTasks.Columns.Add("Path", 200) | Out-Null

function Load-ScheduledTasks {
    $lvTasks.Items.Clear()
    # Get tasks (filtering out some system noise if desired, but showing all here)
    $tasks = Get-ScheduledTask | Where-Object { $_.State -ne "Disabled" } | Sort-Object TaskName
    foreach ($t in $tasks) {
        $item = New-Object System.Windows.Forms.ListViewItem($t.TaskName)
        $item.SubItems.Add($t.State) | Out-Null
        $item.SubItems.Add($t.TaskPath) | Out-Null
        $item.Tag = $t
        $lvTasks.Items.Add($item) | Out-Null
    }
}

$tabTasks.Controls.Add($lvTasks)
$tabTasks.Controls.Add($pnlTasksTop)

$mainForm.Controls.Add($tabControl)

# ------------------------------------------------------------
# Tab 11: System Clean
# ------------------------------------------------------------
$tabClean = New-Object System.Windows.Forms.TabPage
$tabClean.Text = "System Clean"
$tabClean.UseVisualStyleBackColor = $true

$chkTemp = New-Object System.Windows.Forms.CheckBox
$chkTemp.Text = "Temporary Files"
$chkTemp.Location = New-Object System.Drawing.Point(20, 30)
$chkTemp.Size = New-Object System.Drawing.Size(200, 30)
$chkTemp.Checked = $true

$chkRecycle = New-Object System.Windows.Forms.CheckBox
$chkRecycle.Text = "Recycle Bin"
$chkRecycle.Location = New-Object System.Drawing.Point(20, 70)
$chkRecycle.Size = New-Object System.Drawing.Size(200, 30)
$chkRecycle.Checked = $true

$chkBrowser = New-Object System.Windows.Forms.CheckBox
$chkBrowser.Text = "Browser Cache (Edge/Chrome)"
$chkBrowser.Location = New-Object System.Drawing.Point(20, 110)
$chkBrowser.Size = New-Object System.Drawing.Size(300, 30)
$chkBrowser.Checked = $false

$btnRunClean = New-Object System.Windows.Forms.Button
$btnRunClean.Text = "Clean Selected"
$btnRunClean.Location = New-Object System.Drawing.Point(20, 160)
$btnRunClean.Size = New-Object System.Drawing.Size(150, 40)
$btnRunClean.add_Click({
    if ($chkTemp.Checked) {
        Remove-Item -Path "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "$env:windir\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
    }
    if ($chkRecycle.Checked) {
        Clear-RecycleBin -Force -ErrorAction SilentlyContinue
    }
    if ($chkBrowser.Checked) {
        # Edge
        $edgeCache = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache"
        if (Test-Path $edgeCache) { Remove-Item "$edgeCache\*" -Recurse -Force -ErrorAction SilentlyContinue }
        # Chrome
        $chromeCache = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache"
        if (Test-Path $chromeCache) { Remove-Item "$chromeCache\*" -Recurse -Force -ErrorAction SilentlyContinue }
    }
    [System.Windows.Forms.MessageBox]::Show("Cleanup completed.", "WPAS", "OK", "Information")
})

$btnClearStandby = New-Object System.Windows.Forms.Button
$btnClearStandby.Text = "Clear Standby RAM"
$btnClearStandby.Location = New-Object System.Drawing.Point(180, 160)
$btnClearStandby.Size = New-Object System.Drawing.Size(150, 40)
$btnClearStandby.add_Click({
    $script = Join-Path $scriptPath "ClearStandbyMemory.ps1"
    if (Test-Path $script) {
        Start-Process "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$script`"" -Verb RunAs -Wait
        [System.Windows.Forms.MessageBox]::Show("Standby memory cleared.", "WPAS", "OK", "Information")
    } else {
        [System.Windows.Forms.MessageBox]::Show("Script not found: ClearStandbyMemory.ps1", "WPAS Error", "OK", "Error")
    }
})

$tabClean.Controls.AddRange(@($chkTemp, $chkRecycle, $chkBrowser, $btnRunClean, $btnClearStandby))

# ------------------------------------------------------------
# Tab 12: Hardware Specs
# ------------------------------------------------------------
$tabHardware = New-Object System.Windows.Forms.TabPage
$tabHardware.Text = "Hardware Specs"
$tabHardware.UseVisualStyleBackColor = $true

$lvHardware = New-Object System.Windows.Forms.ListView
$lvHardware.Dock = "Fill"
$lvHardware.View = "Details"
$lvHardware.FullRowSelect = $true
$lvHardware.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lvHardware.Columns.Add("Component", 150) | Out-Null
$lvHardware.Columns.Add("Detail", 400) | Out-Null

function Load-HardwareSpecs {
    $lvHardware.Items.Clear()
    
    # CPU
    try {
        $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
        $lvHardware.Items.Add((New-Object System.Windows.Forms.ListViewItem @("CPU Model", $cpu.Name))) | Out-Null
        $lvHardware.Items.Add((New-Object System.Windows.Forms.ListViewItem @("CPU Cores", "$($cpu.NumberOfCores) Cores / $($cpu.NumberOfLogicalProcessors) Threads"))) | Out-Null
        $lvHardware.Items.Add((New-Object System.Windows.Forms.ListViewItem @("Max Clock Speed", "$($cpu.MaxClockSpeed) MHz"))) | Out-Null
    } catch {}

    # GPU
    try {
        $gpus = Get-CimInstance Win32_VideoController
        foreach ($gpu in $gpus) {
            $lvHardware.Items.Add((New-Object System.Windows.Forms.ListViewItem @("GPU Model", $gpu.Name))) | Out-Null
            $lvHardware.Items.Add((New-Object System.Windows.Forms.ListViewItem @("GPU Driver", $gpu.DriverVersion))) | Out-Null
            $vram = [math]::Round($gpu.AdapterRAM / 1GB, 2)
            if ($vram -gt 0) {
                 $lvHardware.Items.Add((New-Object System.Windows.Forms.ListViewItem @("VRAM", "$vram GB"))) | Out-Null
            }
        }
    } catch {}

    # Motherboard
    try {
        $board = Get-CimInstance Win32_BaseBoard | Select-Object -First 1
        $lvHardware.Items.Add((New-Object System.Windows.Forms.ListViewItem @("Motherboard", "$($board.Manufacturer) $($board.Product)"))) | Out-Null
        $lvHardware.Items.Add((New-Object System.Windows.Forms.ListViewItem @("BIOS Version", "$(Get-CimInstance Win32_BIOS | Select-Object -ExpandProperty SMBIOSBIOSVersion)"))) | Out-Null
    } catch {}

    # RAM
    try {
        $memSticks = Get-CimInstance Win32_PhysicalMemory
        $totalMem = 0
        foreach ($stick in $memSticks) {
            $capGB = [math]::Round($stick.Capacity / 1GB, 0)
            $totalMem += $stick.Capacity
            $lvHardware.Items.Add((New-Object System.Windows.Forms.ListViewItem @("RAM Stick", "$capGB GB $($stick.Manufacturer) @ $($stick.Speed) MHz"))) | Out-Null
        }
        $totalMemGB = [math]::Round($totalMem / 1GB, 2)
        $lvHardware.Items.Add((New-Object System.Windows.Forms.ListViewItem @("Total RAM", "$totalMemGB GB"))) | Out-Null
    } catch {}
}

try { Load-HardwareSpecs } catch {}

$tabHardware.Controls.Add($lvHardware)

# ------------------------------------------------------------
# Tab Refresh Logic
# ------------------------------------------------------------
function Refresh-Tabs {
    $tabControl.SuspendLayout()
    $tabControl.TabPages.Clear()
    
    $tabControl.TabPages.Add($tabDashboard)
    if ($script:config.ShowLauncher) { $tabControl.TabPages.Add($tabLauncher) }
    $tabControl.TabPages.Add($tabSettings)
    if ($script:config.ShowLogs) { $tabControl.TabPages.Add($tabLogs) }
    if ($script:config.ShowSysInfo) { $tabControl.TabPages.Add($tabSysInfo) }
    if ($script:config.ShowNetwork) { $tabControl.TabPages.Add($tabNetwork) }
    if ($script:config.ShowServices) { $tabControl.TabPages.Add($tabServices) }
    if ($script:config.ShowProcesses) { $tabControl.TabPages.Add($tabProcesses) }
    if ($script:config.ShowStartupApps) { $tabControl.TabPages.Add($tabStartupApps) }
    if ($script:config.ShowTasks) { $tabControl.TabPages.Add($tabTasks) }
    if ($script:config.ShowSystemClean) { $tabControl.TabPages.Add($tabClean) }
    if ($script:config.ShowHardware) { $tabControl.TabPages.Add($tabHardware) }
    
    $tabControl.ResumeLayout()
}

Refresh-Tabs

# ------------------------------------------------------------
# Theme Logic
# ------------------------------------------------------------
function Set-Theme ($isDark) {
    if ($isDark) {
        $backColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
        $foreColor = [System.Drawing.Color]::White
        $controlBack = [System.Drawing.Color]::FromArgb(45, 45, 48)
        $btnBack = [System.Drawing.Color]::FromArgb(60, 60, 60)
    } else {
        $backColor = [System.Drawing.Color]::White
        $foreColor = [System.Drawing.Color]::Black
        $controlBack = [System.Drawing.Color]::White
        $btnBack = [System.Drawing.Color]::WhiteSmoke
    }

    $mainForm.BackColor = $backColor
    $mainForm.ForeColor = $foreColor

    # Collect all controls to style
    $allControls = @()
    foreach ($tab in $tabControl.TabPages) {
        $allControls += $tab
        $allControls += $tab.Controls
        foreach ($c in $tab.Controls) {
            if ($c -is [System.Windows.Forms.GroupBox] -or $c -is [System.Windows.Forms.Panel]) { $allControls += $c.Controls }
        }
    }

    foreach ($ctrl in $allControls) {
        if ($ctrl -is [System.Windows.Forms.Button]) {
            $ctrl.BackColor = $btnBack
            $ctrl.ForeColor = $foreColor
            $ctrl.FlatStyle = "Flat"
        } elseif ($ctrl -is [System.Windows.Forms.TextBox] -or $ctrl -is [System.Windows.Forms.ComboBox] -or $ctrl -is [System.Windows.Forms.ListView]) {
            $ctrl.BackColor = $controlBack
            $ctrl.ForeColor = $foreColor
        } else {
            $ctrl.BackColor = $backColor
            $ctrl.ForeColor = $foreColor
        }
    }
}

# Apply initial theme
Set-Theme $script:config.DarkMode

# Update theme on checkbox change (preview)
$chkDarkMode.add_CheckedChanged({ Set-Theme $chkDarkMode.Checked })

# ------------------------------------------------------------
# Timer for Auto-Detection
# ------------------------------------------------------------
# State tracking for auto-switch
$script:lastPowerStatus = $null
$script:gpuName = $null
$script:firstTick = $true
$script:lowBatteryWarned = $false

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 3000 # 3 seconds
$timer.add_Tick({
    try {
    if ($script:firstTick) {
        Write-Log "Timer First Tick Executed."
        $script:firstTick = $false
    }
    # Update Power Source
    if ($tabControl.SelectedTab -eq $tabStartupApps -and $lvStartupApps.Items.Count -eq 0) {
        Load-StartupApps
    }
    
    # Load Tasks if tab selected and empty
    if ($tabControl.SelectedTab -eq $tabTasks -and $lvTasks.Items.Count -eq 0) {
        Load-ScheduledTasks
    }


    if (Get-Command Get-WPASPowerSource -ErrorAction SilentlyContinue) {
        $status = Get-WPASPowerSource

        # Auto-Switch Logic
        if ($chkAutoSwitch.Checked -and $script:lastPowerStatus -ne $null -and $status -ne $script:lastPowerStatus) {
            [System.Media.SystemSounds]::Exclamation.Play()
            if ($status -eq "Online") {
                # Switched to AC
                $acScript = Join-Path $scriptPath $script:config.ACScript
                if (Test-Path $acScript) {
                    Start-Process "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$acScript`"" -WindowStyle Hidden
                    Write-Log "Auto-Switch: Switched to AC Mode. Running $acScript"
                    $trayIcon.ShowBalloonTip(1000, "Power Connected", "Switching to High Performance settings.", [System.Windows.Forms.ToolTipIcon]::Info)
                }
            } else {
                # Switched to Battery
                if ($script:config.BatteryScript -eq "Windows Power Saver") {
                    powercfg /setactive SCHEME_MAX
                    Write-Log "Auto-Switch: Switched to Battery Mode. Enabled Windows Power Saver."
                    $trayIcon.ShowBalloonTip(1000, "Battery Mode", "Enabled Power Saver to extend battery life.", [System.Windows.Forms.ToolTipIcon]::Info)
                } else {
                    $batScript = Join-Path $scriptPath $script:config.BatteryScript
                    if (Test-Path $batScript) {
                        Start-Process "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$batScript`"" -WindowStyle Hidden
                        Write-Log "Auto-Switch: Switched to Battery Mode. Running $batScript"
                        $trayIcon.ShowBalloonTip(1000, "Battery Mode", "Optimizing settings for battery life.", [System.Windows.Forms.ToolTipIcon]::Info)
                    }
                }
            }
        }
        $script:lastPowerStatus = $status

        $lblPowerSource.Text = "Power Source: $status"
        if ($status -eq "Online") {
            $lblPowerSource.ForeColor = [System.Drawing.Color]::Green
        } else {
            $lblPowerSource.ForeColor = [System.Drawing.Color]::DarkOrange
        }
    }

    # Update Plan
    if (Get-Command Get-WPASActiveScheme -ErrorAction SilentlyContinue) {
        $plan = Get-WPASActiveScheme
        $lblPlan.Text = "Active Plan: $plan"
    }
    
    # Update GPU
    try {
        if (-not $script:gpuName) {
            $script:gpuName = (Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue | Select-Object -First 1).Name
        }
        # Attempt to get usage (may not work on all drivers)
        $gpuUsageObj = Get-CimInstance Win32_PerfFormattedData_GPUPerformanceCounters_GPUEngine -ErrorAction SilentlyContinue | Measure-Object -Property UtilizationPercentage -Maximum
        $gpuUsage = if ($gpuUsageObj) { $gpuUsageObj.Maximum } else { 0 }
        $lblGPU.Text = "GPU: $script:gpuName | Usage: $gpuUsage %"
    } catch {
        $lblGPU.Text = "GPU: Info Unavailable"
    }
    
    # Update CPU Temp
    try {
        $tempObj = Get-CimInstance -Namespace root/wmi -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction Stop | Select-Object -First 1
        $tempC = [math]::Round(($tempObj.CurrentTemperature - 2732) / 10, 1)
        $lblCPUTemp.Text = "CPU Temp: $tempC C"
    } catch {
        $lblCPUTemp.Text = "CPU Temp: N/A"
    }

    # Update Uptime
    $uptimeSpan = [TimeSpan]::FromMilliseconds([System.Environment]::TickCount64)
    $lblUptime.Text = "System Uptime: $($uptimeSpan.Days)d $($uptimeSpan.Hours)h $($uptimeSpan.Minutes)m"

    # Update RAM Dashboard
    try {
        $os = Get-CimInstance Win32_OperatingSystem
        $totalRam = $os.TotalVisibleMemorySize
        $freeRam = $os.FreePhysicalMemory
        $usedRam = $totalRam - $freeRam
        $ramPercent = [math]::Round(($usedRam / $totalRam) * 100)
        $lblRAMDashboard.Text = "RAM Usage: $ramPercent %"
        $pbRAMDashboard.Value = [math]::Min(100, [math]::Max(0, $ramPercent))
    } catch {
        $lblRAMDashboard.Text = "RAM Usage: N/A"
    }

    # Update Network Speed
    try {
        $netStats = Get-CimInstance Win32_PerfFormattedData_Tcpip_NetworkInterface -ErrorAction SilentlyContinue | Measure-Object -Property BytesReceivedPersec, BytesSentPersec -Sum
        if ($netStats) {
            $rxSum = ($netStats | Where-Object Property -eq "BytesReceivedPersec").Sum
            $txSum = ($netStats | Where-Object Property -eq "BytesSentPersec").Sum
            $rxKB = [math]::Round($rxSum / 1KB, 1)
            $txKB = [math]::Round($txSum / 1KB, 1)
            $lblNetSpeed.Text = "Network Speed: DL $rxKB KB/s  |  UL $txKB KB/s"
        }
    } catch {
        $lblNetSpeed.Text = "Network Speed: N/A"
    }

    # Update Disk Speed
    try {
        $diskStats = Get-CimInstance Win32_PerfFormattedData_PerfDisk_PhysicalDisk -Filter "Name='_Total'" -ErrorAction SilentlyContinue
        $diskStats = Get-CimInstance Win32_PerfFormattedData_PerfDisk_PhysicalDisk -ErrorAction SilentlyContinue | Where-Object Name -eq "_Total"
        if ($diskStats) {
            $readKB = [math]::Round($diskStats.DiskReadBytesPersec / 1KB, 1)
            $writeKB = [math]::Round($diskStats.DiskWriteBytesPersec / 1KB, 1)
            $lblDiskSpeed.Text = "Disk Speed: Read $readKB KB/s  |  Write $writeKB KB/s"
        }
    } catch {
        $lblDiskSpeed.Text = "Disk Speed: N/A"
    }

    # Update Battery
    if (Get-Command Get-WPASBatteryStatus -ErrorAction SilentlyContinue) {
        $bat = Get-WPASBatteryStatus
        if ($bat.Status -eq "Available") {
            $lblBatDesign.Text = "Design Capacity: $($bat.DesignCapacity) mWh"
            $lblBatFull.Text = "Full Charge Capacity: $($bat.FullChargeCapacity) mWh"
            $lblBatWear.Text = "Wear Level: $($bat.WearLevel) %"
            $lblBatPercent.Text = "Battery Percentage: $($bat.ChargeRemaining) %"
        } else {
            $lblBatDesign.Text = "Battery data unavailable."
        }
        
        # Low Battery Warning
        if ($bat.ChargeRemaining -lt 20 -and $status -ne "Online" -and -not $script:lowBatteryWarned) {
            [System.Media.SystemSounds]::Hand.Play()
            $trayIcon.ShowBalloonTip(3000, "Battery Low", "Battery is below 20%. Please plug in.", [System.Windows.Forms.ToolTipIcon]::Warning)
            $script:lowBatteryWarned = $true
        } elseif ($bat.ChargeRemaining -ge 20 -or $status -eq "Online") {
            $script:lowBatteryWarned = $false
        }
    }

    # Update System Info
    if ($tabControl.SelectedTab -eq $tabSysInfo) {
        try {
            $cpu = Get-CimInstance Win32_Processor | Measure-Object -Property LoadPercentage -Average | Select-Object -ExpandProperty Average
            $lblSysCPU.Text = "CPU Usage: $cpu %"
            
            $os = Get-CimInstance Win32_OperatingSystem
            $totalRam = [math]::Round(($os.TotalVisibleMemorySize * 1KB) / 1GB, 2)
            $freeRam = [math]::Round(($os.FreePhysicalMemory * 1KB) / 1GB, 2)
            $usedRam = [math]::Round($totalRam - $freeRam, 2)
            $lblSysRAM.Text = "RAM Usage: $usedRam GB / $totalRam GB"

            $disk = Get-PSDrive C
            $usedDisk = [math]::Round($disk.Used / 1GB, 2)
            $freeDisk = [math]::Round($disk.Free / 1GB, 2)
            $lblSysDisk.Text = "Disk (C:) Usage: $usedDisk GB Used / $freeDisk GB Free"
        } catch {
            $lblSysCPU.Text = "Error retrieving system info."
        }
    }

    # Update Network Info
    if ($tabControl.SelectedTab -eq $tabNetwork) {
        $isNetworkUp = [System.Net.NetworkInformation.NetworkInterface]::GetIsNetworkAvailable()
        if ($isNetworkUp) {
            $lblNetStatus.Text = "Internet Status: Connected"
            $lblNetStatus.ForeColor = [System.Drawing.Color]::Green
        } else {
            $lblNetStatus.Text = "Internet Status: Disconnected"
            $lblNetStatus.ForeColor = [System.Drawing.Color]::Red
        }

        $ip = Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike "*Loopback*" -and $_.ConnectionState -eq "Connected" } | Select-Object -ExpandProperty IPAddress -First 1
        if ($ip) { $lblNetIP.Text = "IP Address: $ip" } else { $lblNetIP.Text = "IP Address: Not Found" }
    }

    # Update Services Info
    if ($tabControl.SelectedTab -eq $tabServices -and $cmbServices.SelectedItem) {
        $svcName = $cmbServices.SelectedItem.Split(" ")[0]
        $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
        if ($svc) {
            $lblSvcStatus.Text = "Status: $($svc.Status)"
            if ($svc.Status -eq "Running") {
                $lblSvcStatus.ForeColor = [System.Drawing.Color]::Green
            } else {
                $lblSvcStatus.ForeColor = [System.Drawing.Color]::Red
            }
        }
    }

    # Update Processes Info
    if ($tabControl.SelectedTab -eq $tabProcesses) {
        try {
            $processes = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfProc_Process | Where-Object { $_.Name -notin "Idle", "_Total" } | Sort-Object -Property PercentProcessorTime -Descending | Select-Object -First 5
            $lvProcesses.BeginUpdate()
            $lvProcesses.Items.Clear()
            foreach ($p in $processes) {
                $procInfo = Get-Process -Id $p.IDProcess -ErrorAction SilentlyContinue
                if ($procInfo) {
                    $item = New-Object System.Windows.Forms.ListViewItem($procInfo.ProcessName)
                    $item.SubItems.Add($procInfo.Id)
                    $item.SubItems.Add($p.PercentProcessorTime)
                    $item.SubItems.Add([math]::Round($procInfo.WorkingSet64 / 1MB, 2))
                    $item.Tag = $procInfo.Id
                    $lvProcesses.Items.Add($item) | Out-Null
                }
            }
        } catch {} finally { $lvProcesses.EndUpdate() }
    }
    } catch {
        Write-Log "Timer Error: $_"
    }
})

# ------------------------------------------------------------
# Run
# ------------------------------------------------------------
Write-Log "Starting Application Loop..."
try {
    [System.Windows.Forms.Application]::Run($mainForm)
} catch {
    Write-Log "Fatal Error in Application Loop: $_"
}
Write-Log "Application Loop Ended."