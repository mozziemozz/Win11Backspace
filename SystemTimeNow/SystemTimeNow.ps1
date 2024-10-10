<#
    .SYNOPSIS
    This script shows the current system time in a balloon tip.

    .DESCRIPTION
    This script shows the current system time in a balloon tip and can be launched from a taskbar shortcut.

    .NOTES
    Author:             Martin Heusser | M365 Apps & Services MVP
    Version:            1.0.0
    Changes:            2024-10-10 - Initial version

    Sponsor Project:    https://github.com/sponsors/mozziemozz
    Website:            https://heusser.pro

    .NOTES
    This script has been tested on Windows 11
    The shortcut will be placed in $env:OneDrive\Desktop\SystemTimeNow.lnk.

    .EXAMPLE
    This script is not intended to be launched from PowerShell directly. Instead, use the file 'CreateShortcut.ps1' to create a shortcut on your desktop and pin it to the taskbar.
    The shortcut has to be manually pinned to the taskbar. This step is part of the setup process when you launch 'createShortcut.ps1' for the first time.

#>

if ((Get-Clipboard).Trim() -eq "SetUpSystemTimeNow") {

    $Message = "Right-click on the icon in the taskbar`nand select 'Pin to taskbar' to pin the app.`n`nClick OK when you're done."
    $Title = "System Time Now | SETUP"

    # Show hint if there is no phone number in clipboard
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [Windows.Forms.MessageBox]::Show($Message, $Title, [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information)
        
}

$dateTime = Get-Date

$dateTime.ToString("yyyy-MM-dd HH:mm:ss") | Set-Clipboard

$dateTimeString = ($dateTime | Out-String).Trim()

$timeZone = (Get-TimeZone).Id

$offset = ([DateTimeOffset]::Now).Offset.TotalHours

if ($offset -ge 1) {

    $offset = "+$offset"

}

$Message = @"

$dateTimeString
Timezone: $timeZone
UTC Offset: $offset
"@
$Title = "System Time Now"

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$balloonTip = New-Object System.Windows.Forms.NotifyIcon
$balloonTip.Icon = [System.Drawing.SystemIcons]::Information
$balloonTip.BalloonTipIcon = "Info"
$balloonTip.BalloonTipTitle = $Title
$balloonTip.BalloonTipText = $Message 
$balloonTip.Visible = $True
$balloonTip.ShowBalloonTip(30000)
