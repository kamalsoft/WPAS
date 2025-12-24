# ============================================================
# Windows Power Automation Suite (WPAS) Module
# ============================================================

function Get-WPASPowerSource {
    <#
    .SYNOPSIS
        Returns the current power source (Online/Offline).
    #>
    Add-Type -AssemblyName System.Windows.Forms
    return [System.Windows.Forms.SystemInformation]::PowerStatus.PowerLineStatus
}

function Get-WPASActiveScheme {
    <#
    .SYNOPSIS
        Returns the name of the active power plan.
    #>
    $planOutput = powercfg /getactivescheme
    if ($planOutput -match "\((.*?)\)$") {
        return $matches[1]
    }
    return "Unknown"
}

function Get-WPASBatteryStatus {
    <#
    .SYNOPSIS
        Returns battery health information.
    #>
    $batteries = Get-CimInstance -Class Win32_Battery -ErrorAction SilentlyContinue
    $battery = $batteries | Select-Object -First 1
    $result = [PSCustomObject]@{
        DesignCapacity     = 0
        FullChargeCapacity = 0
        WearLevel          = 0
        Status             = "No Battery Detected"
        ChargeRemaining    = 0
    }

    if ($battery) {
        $result.DesignCapacity = $battery.DesignCapacity
        $result.FullChargeCapacity = $battery.FullChargeCapacity
        $result.Status = "Available"
        $result.ChargeRemaining = $battery.EstimatedChargeRemaining

        if ($result.DesignCapacity -gt 0 -and $result.FullChargeCapacity -gt 0) {
            $result.WearLevel = [math]::Round((1 - ($result.FullChargeCapacity / $result.DesignCapacity)) * 100, 2)
        }
    }
    return $result
}

function Get-WPASActivationStatus {
    <#
    .SYNOPSIS
        Returns the Windows Activation Status.
    #>
    try {
        $lics = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "PartialProductKey IS NOT NULL AND Name LIKE 'Windows%'" -ErrorAction SilentlyContinue
        $lic = $lics | Select-Object -First 1
        if ($lic) {
            switch ($lic.LicenseStatus) {
                0 { return "Unlicensed" }
                1 { return "Licensed" }
                2 { return "OOBE Grace" }
                3 { return "OOT Grace" }
                4 { return "Non-Genuine Grace" }
                5 { return "Notification" }
                6 { return "Extended Grace" }
                default { return "Unknown ($($lic.LicenseStatus))" }
            }
        }
    } catch {}
    return "Unknown"
}

Export-ModuleMember -Function Get-WPASPowerSource, Get-WPASActiveScheme, Get-WPASBatteryStatus, Get-WPASActivationStatus