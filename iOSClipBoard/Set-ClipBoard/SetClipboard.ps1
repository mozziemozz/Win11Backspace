# Create an Outlook Application object
$Outlook = New-Object -ComObject Outlook.Application

# Create a new note item
$Note = $Outlook.CreateItem([Microsoft.Office.Interop.Outlook.OlItemType]::olNoteItem)

# Set the note's properties
$Note.Body = (Get-Clipboard | Out-String).Trim()
$Note.Save()

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$balloonTip = New-Object System.Windows.Forms.NotifyIcon
$balloonTip.Icon = [System.Drawing.SystemIcons]::Information
# $balloonTip.BalloonTipIcon = "Info"
$balloonTip.BalloonTipTitle = "Clipboard Content sent to iPhone"
$balloonTip.BalloonTipText = (Get-Clipboard | Out-String).Trim()
$balloonTip.Visible = $True
$balloonTip.ShowBalloonTip(30000)


Start-Sleep -Seconds 3

# Release the Outlook objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Note) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null

# Perform garbage collection
[GC]::Collect()
[GC]::WaitForPendingFinalizers()