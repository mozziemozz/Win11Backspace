# Create an Outlook Application object
$Outlook = New-Object -ComObject Outlook.Application

# Get the MAPI namespace
$Namespace = $Outlook.GetNamespace("MAPI")

# Access the Notes folder
$NotesFolder = $Namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderNotes)

# Retrieve all notes and sort them by last modified time in descending order
$Notes = $NotesFolder.Items | Sort-Object -Property LastModificationTime -Descending

($Notes[0].Body).Trim() | Set-Clipboard

$noteAge = (Get-Date) - $Notes[0].LastModificationTime

$text = @"
$($Notes[0].Body.Trim())
"@

if ($noteAge.TotalMinutes -le 5) {

    # Delete the note
    $Notes[0].Delete()

    $text += "`n`nNote deleted"

}

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$balloonTip = New-Object System.Windows.Forms.NotifyIcon
$balloonTip.Icon = [System.Drawing.SystemIcons]::Information
# $balloonTip.BalloonTipIcon = "Info"
$balloonTip.BalloonTipTitle = "Clipboard Content from iPhone"
$balloonTip.BalloonTipText = $text
$balloonTip.Visible = $True
$balloonTip.ShowBalloonTip(30000)

Start-Sleep -Seconds 3

# Release the Outlook objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($NotesFolder) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Namespace) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null

# Perform garbage collection
[GC]::Collect()
[GC]::WaitForPendingFinalizers()