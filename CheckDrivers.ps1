# ============================================================
# WPAS Driver Version Checker
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Driver Version Checker"
$form.Size = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text = "Scanning system drivers... This may take a moment."
$lblInfo.Location = New-Object System.Drawing.Point(20, 15)
$lblInfo.Size = New-Object System.Drawing.Size(600, 25)
$lblInfo.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$lvDrivers = New-Object System.Windows.Forms.ListView
$lvDrivers.Location = New-Object System.Drawing.Point(20, 50)
$lvDrivers.Size = New-Object System.Drawing.Size(840, 450)
$lvDrivers.View = "Details"
$lvDrivers.FullRowSelect = $true
$lvDrivers.GridLines = $true

$lvDrivers.Columns.Add("Device Name", 300) | Out-Null
$lvDrivers.Columns.Add("Driver Date", 100) | Out-Null
$lvDrivers.Columns.Add("Version", 120) | Out-Null
$lvDrivers.Columns.Add("Provider", 200) | Out-Null
$lvDrivers.Columns.Add("Age (Days)", 80) | Out-Null

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh List"
$btnRefresh.Location = New-Object System.Drawing.Point(20, 510)
$btnRefresh.Size = New-Object System.Drawing.Size(120, 35)

function Load-Drivers {
    $lvDrivers.Items.Clear()
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $lblInfo.Text = "Querying Win32_PnPSignedDriver... Please wait."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        # Get drivers (filtering for those with a valid date)
        $drivers = Get-CimInstance -ClassName Win32_PnPSignedDriver -ErrorAction SilentlyContinue | Where-Object { $_.DriverDate -ne $null -and $_.DeviceName -ne $null } | Sort-Object DriverDate
        
        $lvDrivers.BeginUpdate()
        foreach ($drv in $drivers) {
            $date = $drv.DriverDate
            # Handle CIM DateTime if not automatically converted
            if ($date -is [string]) {
                try { $date = [Management.ManagementDateTimeConverter]::ToDateTime($date) } catch {}
            }
            
            $age = (New-TimeSpan -Start $date -End (Get-Date)).Days
            
            $item = New-Object System.Windows.Forms.ListViewItem($drv.DeviceName)
            $item.SubItems.Add($date.ToString("yyyy-MM-dd")) | Out-Null
            $item.SubItems.Add($drv.DriverVersion) | Out-Null
            $item.SubItems.Add($drv.DriverProviderName) | Out-Null
            $item.SubItems.Add($age) | Out-Null
            
            # Highlight old drivers (> 2 years)
            if ($age -gt 730) {
                $item.ForeColor = [System.Drawing.Color]::DarkOrange
            }
            
            $lvDrivers.Items.Add($item) | Out-Null
        }
        $lvDrivers.EndUpdate()
        $lblInfo.Text = "Found $($drivers.Count) drivers. Orange items are older than 2 years."
    } catch {
        $lblInfo.Text = "Error retrieving drivers: $($_.Exception.Message)"
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

$btnRefresh.add_Click({ Load-Drivers })

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(760, 510)
$btnClose.Size = New-Object System.Drawing.Size(100, 35)
$btnClose.add_Click({ $form.Close() })

$form.Controls.AddRange(@($lblInfo, $lvDrivers, $btnRefresh, $btnClose))

$form.Add_Shown({ Load-Drivers })

[void]$form.ShowDialog()