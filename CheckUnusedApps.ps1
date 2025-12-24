# ============================================================
# WPAS Old/Unused Application Checker
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Potentially Unused Applications (> 6 Months Old)"
$form.Size = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text = "Scanning installed applications... Listing apps installed over 6 months ago."
$lblInfo.Location = New-Object System.Drawing.Point(20, 15)
$lblInfo.Size = New-Object System.Drawing.Size(800, 25)
$lblInfo.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$lvApps = New-Object System.Windows.Forms.ListView
$lvApps.Location = New-Object System.Drawing.Point(20, 50)
$lvApps.Size = New-Object System.Drawing.Size(840, 450)
$lvApps.View = "Details"
$lvApps.FullRowSelect = $true
$lvApps.GridLines = $true

$lvApps.Columns.Add("Application Name", 350) | Out-Null
$lvApps.Columns.Add("Install Date", 100) | Out-Null
$lvApps.Columns.Add("Age (Days)", 80) | Out-Null
$lvApps.Columns.Add("Version", 100) | Out-Null
$lvApps.Columns.Add("Publisher", 180) | Out-Null

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh List"
$btnRefresh.Location = New-Object System.Drawing.Point(20, 510)
$btnRefresh.Size = New-Object System.Drawing.Size(120, 35)

function Get-InstalledApps {
    $apps = @()
    $paths = @("HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
    
    foreach ($path in $paths) {
        Get-ItemProperty $path -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -and $_.InstallDate } | ForEach-Object {
            $apps += $_
        }
    }
    return $apps
}

function Load-Apps {
    $lvApps.Items.Clear()
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $lblInfo.Text = "Scanning Registry for installed applications..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $apps = Get-InstalledApps
        $thresholdDate = (Get-Date).AddMonths(-6)
        $count = 0

        $lvApps.BeginUpdate()
        foreach ($app in $apps) {
            try {
                # Parse yyyyMMdd
                $dateStr = $app.InstallDate
                if ($dateStr -match "^\d{8}$") {
                    $installDate = [DateTime]::ParseExact($dateStr, "yyyyMMdd", $null)
                    
                    if ($installDate -lt $thresholdDate) {
                        $age = (New-TimeSpan -Start $installDate -End (Get-Date)).Days
                        
                        $item = New-Object System.Windows.Forms.ListViewItem($app.DisplayName)
                        $item.SubItems.Add($installDate.ToString("yyyy-MM-dd")) | Out-Null
                        $item.SubItems.Add($age) | Out-Null
                        $item.SubItems.Add([string]$app.DisplayVersion) | Out-Null
                        $item.SubItems.Add([string]$app.Publisher) | Out-Null
                        
                        # Highlight very old apps (> 1 year)
                        if ($age -gt 365) { $item.ForeColor = [System.Drawing.Color]::DarkRed }
                        
                        $lvApps.Items.Add($item) | Out-Null
                        $count++
                    }
                }
            } catch {}
        }
        $lvApps.EndUpdate()
        $lblInfo.Text = "Found $count applications installed more than 6 months ago."
    } catch {
        $lblInfo.Text = "Error: $($_.Exception.Message)"
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

$btnRefresh.add_Click({ Load-Apps })

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(760, 510)
$btnClose.Size = New-Object System.Drawing.Size(100, 35)
$btnClose.add_Click({ $form.Close() })

$form.Controls.AddRange(@($lblInfo, $lvApps, $btnRefresh, $btnClose))
$form.Add_Shown({ Load-Apps })
[void]$form.ShowDialog()