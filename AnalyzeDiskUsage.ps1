# ============================================================
# WPAS Disk Usage Analyzer
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Top 5 Largest Folders (C:\)"
$form.Size = New-Object System.Drawing.Size(500, 400)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text = "Click 'Start Scan' to analyze top-level folders on C:\ drive.`nNote: This may take a minute depending on disk speed."
$lblInfo.Location = New-Object System.Drawing.Point(20, 20)
$lblInfo.Size = New-Object System.Drawing.Size(440, 40)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Start Scan"
$btnScan.Location = New-Object System.Drawing.Point(20, 70)
$btnScan.Size = New-Object System.Drawing.Size(100, 30)

$lvResults = New-Object System.Windows.Forms.ListView
$lvResults.Location = New-Object System.Drawing.Point(20, 110)
$lvResults.Size = New-Object System.Drawing.Size(440, 230)
$lvResults.View = "Details"
$lvResults.GridLines = $true
$lvResults.FullRowSelect = $true

$lvResults.Columns.Add("Folder Path", 280) | Out-Null
$lvResults.Columns.Add("Size (GB)", 100) | Out-Null

$btnScan.add_Click({
    $btnScan.Enabled = $false
    $btnScan.Text = "Scanning..."
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $lvResults.Items.Clear()
    
    # Force UI update
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $root = "C:\"
        $folders = Get-ChildItem $root -Directory -ErrorAction SilentlyContinue
        $folderSizes = @()

        foreach ($folder in $folders) {
            # Calculate size (recursive)
            # Using Measure-Object is standard, though slow for huge trees like Windows
            # We suppress errors for Access Denied
            try {
                $sizeBytes = (Get-ChildItem $folder.FullName -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                if ($sizeBytes -gt 0) {
                    $folderSizes += [PSCustomObject]@{
                        Path = $folder.FullName
                        Size = $sizeBytes
                    }
                }
            } catch {}
        }

        # Sort and take Top 5
        $top5 = $folderSizes | Sort-Object Size -Descending | Select-Object -First 5

        foreach ($item in $top5) {
            $gb = [math]::Round($item.Size / 1GB, 2)
            $lvItem = New-Object System.Windows.Forms.ListViewItem($item.Path)
            $lvItem.SubItems.Add("$gb GB") | Out-Null
            $lvResults.Items.Add($lvItem) | Out-Null
        }

        if ($top5.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Could not retrieve folder sizes. Ensure you have Administrator privileges.", "WPAS Info", "OK", "Warning")
        }

    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error during scan: $($_.Exception.Message)", "WPAS Error", "OK", "Error")
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnScan.Text = "Start Scan"
        $btnScan.Enabled = $true
    }
})

$form.Controls.AddRange(@($lblInfo, $btnScan, $lvResults))
[void]$form.ShowDialog()