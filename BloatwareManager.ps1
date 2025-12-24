# ============================================================
# WPAS Bloatware Manager
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "WPAS Bloatware Manager"
$form.Size = New-Object System.Drawing.Size(600, 500)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Label
$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text = "Select applications to remove (Current User):"
$lblInfo.Location = New-Object System.Drawing.Point(10, 10)
$lblInfo.Size = New-Object System.Drawing.Size(320, 20)
$lblInfo.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Search
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Search:"
$lblSearch.Location = New-Object System.Drawing.Point(340, 10)
$lblSearch.Size = New-Object System.Drawing.Size(50, 20)
$lblSearch.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(390, 8)
$txtSearch.Size = New-Object System.Drawing.Size(180, 25)

# ListView
$lvApps = New-Object System.Windows.Forms.ListView
$lvApps.Location = New-Object System.Drawing.Point(10, 40)
$lvApps.Size = New-Object System.Drawing.Size(560, 350)
$lvApps.View = "Details"
$lvApps.CheckBoxes = $true
$lvApps.FullRowSelect = $true
$lvApps.GridLines = $true

$lvApps.Columns.Add("Application Name", 250) | Out-Null
$lvApps.Columns.Add("Publisher", 200) | Out-Null
$lvApps.Columns.Add("Version", 80) | Out-Null

$script:cachedApps = @()

function Filter-Apps {
    $lvApps.BeginUpdate()
    $lvApps.Items.Clear()
    $filter = $txtSearch.Text
    
    $source = $script:cachedApps
    if (-not [string]::IsNullOrWhiteSpace($filter)) {
        $source = $source | Where-Object { $_.Name -like "*$filter*" -or $_.Publisher -like "*$filter*" }
    }
    
    foreach ($app in $source) {
        $item = New-Object System.Windows.Forms.ListViewItem($app.Name)
        $item.SubItems.Add($app.Publisher) | Out-Null
        $item.SubItems.Add($app.Version) | Out-Null
        $item.Tag = $app.PackageFullName
        $lvApps.Items.Add($item) | Out-Null
    }
    $lvApps.EndUpdate()
}

# Load Apps Function
function Load-Apps {
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    # Get Non-System Apps
    $script:cachedApps = Get-AppxPackage | Where-Object { $_.NonRemovable -eq $false -and $_.IsFramework -eq $false } | Sort-Object Name
    Filter-Apps
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
}

$txtSearch.Add_TextChanged({ Filter-Apps })

# Buttons
$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh List"
$btnRefresh.Location = New-Object System.Drawing.Point(10, 410)
$btnRefresh.Size = New-Object System.Drawing.Size(100, 35)
$btnRefresh.add_Click({ Load-Apps })

$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = "Remove Selected"
$btnRemove.Location = New-Object System.Drawing.Point(420, 410)
$btnRemove.Size = New-Object System.Drawing.Size(150, 35)
$btnRemove.ForeColor = [System.Drawing.Color]::White
$btnRemove.BackColor = [System.Drawing.Color]::Firebrick
$btnRemove.FlatStyle = "Flat"
$btnRemove.add_Click({
    $checked = $lvApps.CheckedItems
    if ($checked.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No applications selected.", "WPAS", "OK", "Warning")
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to remove $($checked.Count) application(s)?`nThis action cannot be easily undone.", "Confirm Removal", "YesNo", "Warning")
    
    if ($confirm -eq "Yes") {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $removedCount = 0
        foreach ($item in $checked) {
            try {
                Remove-AppxPackage -Package $item.Tag -ErrorAction Stop
                $removedCount++
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to remove $($item.Text):`n$($_.Exception.Message)", "Error", "OK", "Error")
            }
        }
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        [System.Windows.Forms.MessageBox]::Show("Successfully removed $removedCount application(s).", "WPAS", "OK", "Information")
        Load-Apps
    }
})

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Select All"
$btnSelectAll.Location = New-Object System.Drawing.Point(120, 410)
$btnSelectAll.Size = New-Object System.Drawing.Size(100, 35)
$btnSelectAll.add_Click({
    foreach ($item in $lvApps.Items) { $item.Checked = $true }
})

$form.Controls.AddRange(@($lblInfo, $lblSearch, $txtSearch, $lvApps, $btnRefresh, $btnRemove, $btnSelectAll))

# Initial Load
$form.Add_Shown({ Load-Apps })

[void]$form.ShowDialog()