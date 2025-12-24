# ============================================================
# WPAS Battery History Viewer
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Windows.Forms.DataVisualization
Add-Type -AssemblyName System.Drawing

$scriptPath = $PSScriptRoot
$csvFile = Join-Path $scriptPath "logs\BatteryHistory.csv"

# Form Setup
$form = New-Object System.Windows.Forms.Form
$form.Text = "Battery Wear Level History"
$form.Size = New-Object System.Drawing.Size(800, 500)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

if (-not (Test-Path $csvFile)) {
    [System.Windows.Forms.MessageBox]::Show("No battery history log found.`nRun 'BatteryHealth.ps1' or wait for the scheduled task to run.", "WPAS Info", "OK", "Information")
    exit
}

# Chart Setup
$chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$chart.Dock = "Fill"
$chart.BackColor = [System.Drawing.Color]::White

$chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
$chartArea.Name = "MainArea"
$chartArea.AxisX.Title = "Date"
$chartArea.AxisX.IntervalAutoMode = "VariableCount"
$chartArea.AxisX.LabelStyle.Format = "MM-dd"
$chartArea.AxisX.MajorGrid.LineColor = [System.Drawing.Color]::LightGray
$chartArea.AxisY.Title = "Wear Level (%)"
$chartArea.AxisY.MajorGrid.LineColor = [System.Drawing.Color]::LightGray
$chart.ChartAreas.Add($chartArea)

$series = New-Object System.Windows.Forms.DataVisualization.Charting.Series
$series.Name = "WearLevel"
$series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
$series.BorderWidth = 3
$series.Color = [System.Drawing.Color]::Crimson
$series.XValueType = [System.Windows.Forms.DataVisualization.Charting.ChartValueType]::DateTime
$chart.Series.Add($series)

$title = New-Object System.Windows.Forms.DataVisualization.Charting.Title
$title.Text = "Battery Wear Level Over Time"
$title.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$chart.Titles.Add($title)

# Load Data
try {
    $data = Import-Csv $csvFile
    if ($data.Count -eq 0) { throw "Empty CSV" }

    foreach ($row in $data) {
        if ($row.Timestamp -and $row.WearLevel) {
            try {
                [DateTime]$dt = $row.Timestamp
                [double]$wear = $row.WearLevel
                $series.Points.AddXY($dt, $wear) | Out-Null
            } catch {}
        }
    }
} catch {
    $lblErr = New-Object System.Windows.Forms.Label
    $lblErr.Text = "Not enough data to display graph or error reading CSV."
    $lblErr.AutoSize = $true
    $lblErr.Location = New-Object System.Drawing.Point(20, 20)
    $form.Controls.Add($lblErr)
}

# Close Button (Optional, standard form close works)

$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = "Fill"
$panel.Controls.Add($chart)

$form.Controls.Add($panel)

$form.Add_Shown({ 
    $form.Activate() 
    $chart.Focus()
})
[void]$form.ShowDialog()