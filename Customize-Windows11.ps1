Set-ExecutionPolicy Bypass

# Set network profile to private
Get-NetConnectionProfile | Set-NetConnectionProfile -NetworkCategory Private

# Enable ICMPv4 Echo Request
Enable-NetFirewallRule -Name "FPS-ICMP4-ERQ-In"

Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

if (-not (Test-Path -Path "C:\Temp")) {

    New-Item -Path "C:\Temp" -ItemType Directory

}

# Install Office

$checkOfficeInstall = winget list --id "O365HomePremRetail"

if (-not $checkOfficeInstall) {

    $officeSetupPath = "$env:OneDriveConsumer\Software\Windows 11\OfficeSetup.exe"
    Start-Process -FilePath $officeSetupPath   

}

$uninstallApps = @(

    "Microsoft.549981C3F5F10_8wekyb3d8bbwe", # Cortana
    "Microsoft.BingNews_8wekyb3d8bbwe", # Microsoft News
    "Microsoft.MicrosoftSolitaireCollection_8wekyb3d8bbwe", # Microsoft Solitaire Collection
    "Microsoft.WindowsFeedbackHub_8wekyb3d8bbwe", # Feedback Hub
    "microsoft.windowscommunicationsapps_8wekyb3d8bbwe", # Mail and Calendar
    "MicrosoftTeams_8wekyb3d8bbwe", # Microsoft Teams (Personal)
    "Microsoft.Getstarted_8wekyb3d8bbwe", # Microsoft Tips
    "Microsoft.GetHelp_8wekyb3d8bbwe", # Get Help
    "Microsoft.MicrosoftOfficeHub_8wekyb3d8bbwe" # Microsoft Office

)

foreach ($app in $uninstallApps) {

    winget uninstall --id $app --silent

}

$apps = @(

    "9PKVZCPRX9NV", # ShowKeyPlus
    "9N8G7TSCL18R", # NanaZip
    "9WZDNCRFJ3PS", # Remote Desktop
    "9WZDNCRFJB8P", # Surface
    "9WZDNCRFHWLH" # HP Smart
    "9NQDW009T0T5", # Omen Gaming Hub
    "9N013P0KR5VX", # Microsoft Accessory Center
    "XP89DCGQ3K6VLD", # PowerToys
    "9PGCV4V3BK4W", # DevToys
    "9P170799PF3Q", # PageFabric
    "9PC3H3V7Q9CH", # Rufus
    "9NKSQGP7F2NH", # WhatsApp
    "9wzdncrfj364", # Skype
    "XPDC2RH70K22MN", # Discord
    "9N3SQK8PDS8G", # ScreenToGif
    # "9PC7BZZ28G0X", # Custom Context Menu
    # "9PP2MWPBZFGH", # Paste File for File Explorer
    "XP9KHM4BK9FZ7Q", # VS Code
    "9PC3H3V7Q9CH", # Rufus
    "9PLDPG46G47Z", # Xbox Insider Hub
    "9NBLGGH30XJ3", # Xbox Accessories
    "XP9CDQW6ML4NQN", # Plex for Windows
    "9WZDNCRFJ3TJ" # Netflix
    "Git.Git", # Git for Windows
    "GitHub.GitHubDesktop", # GitHub Desktop
    "9MZ1SNWT0N5D", # PowerShell 7
    "Microsoft.DotNet.SDK.7", # .NET 7 SDK
    "OpenJS.NodeJS", # Node.JS
    "9NRWMJP3717K", # Python 3.11
    "OpenVPNTechnologies.OpenVPNConnect", # OpenVPN Connect
    "ImageMagick.ImageMagick" # ImageMagick

)

foreach ($app in $apps) {

    $checkApp = winget list --id $app

    if ($checkApp -match $app) {

        Write-Host "$app is already installed"

    }

    else {

        Write-Host "Installing $app..."
        winget install --id $app --silent --accept-package-agreements --accept-source-agreements

    }

}

# Create ps custom object to store app names and download links
$desktopApps = @(
    
    [pscustomobject]@{
        Name = "Bitwarden"
        DownloadLink = "https://vault.bitwarden.com/download/?app=desktop&platform=windows&variant=exe"
        FileType = "exe"
        DetectionMethod = "winget"
    },
    [pscustomobject]@{
        Name = "SDeleteGui"
        DownloadLink = "https://github.com/Tulpep/SDelete-Gui/releases/download/SDelete-Gui-1.3.4/SDelete-Gui.exe"
        FileType = "exe"
        DetectionMethod = "filepath"
    },
    [pscustomobject]@{
        Name = "VCRedist"
        DownloadLink = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
        FileType = "exe"
        DetectionMethod = "winget"
    }

)

foreach ($app in $desktopApps) {

    switch ($app.DetectionMethod) {
        winget {

            $checkApp = winget list --id $app.Name

            if ($checkApp -match $app.Name) {
        
                Write-Host "$($app.Name) is already installed"
        
            }
        
            else {
        
                $fileName = "$env:USERPROFILE\Downloads\$($app.Name).$($app.FileType)"
        
                Invoke-WebRequest -Uri $app.DownloadLink -OutFile $fileName
            
                Start-Process -FilePath $fileName -ArgumentList "/S" -Wait
            
                if ($?) {
            
                    Remove-Item -Path $fileName
            
                }
            
            }
        
        }
        filepath {

            if (-not (Test-Path -Path "C:\Windows\SysWOW64\sdelete.exe")) {

                if (-not (Test-Path -Path "C:\Temp\Apps")) {

                    New-Item -Path "C:\Temp\Apps" -ItemType Directory

                }

                $fileName = "C:\Temp\Apps\$($app.Name).$($app.FileType)"

                Invoke-WebRequest -Uri $app.DownloadLink -OutFile $fileName
            
                Start-Process -FilePath $fileName -ArgumentList "/S" -Wait

            }

            else {

                Write-Host "$($app.Name) is already installed"

            }

        }
        Default {}
    }

}

$optionalFeatures = @(

    "Microsoft-Hyper-V-All", # Hyper-V
    "Containers-DisposableClientVM" # Windows Sandbox

)

foreach ($feature in $optionalFeatures) {

    $checkFeature = Get-WindowsOptionalFeature -Online -FeatureName $feature

    if ($checkFeature.State -eq "Enabled") {

        Write-Host "$($feature) is already enabled"

    }

    else {

        Write-Host "Enabling $($feature)..."
        Enable-WindowsOptionalFeature -Online -FeatureName $feature -All -NoRestart

    }


}

$optionalApps = @(

    "App.WirelessDisplay.Connect~~~~0.0.1.0"

)

foreach ($app in $optionalApps) {

    $checkApp = Get-WindowsCapability -Online -Name $app

    if ($checkApp.State -eq "Installed") {

        Write-Host "$($app) is already installed"

    }

    else {

        Write-Host "Installing $($app)..."
        Add-WindowsCapability -Online -Name App.WirelessDisplay.Connect~~~~0.0.1.0

    }

}

# DocFx

# Reload path environment variable
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine")

dotnet tool install -g --add-source 'https://api.nuget.org/v3/index.json' --ignore-failed-sources "docfx"

# PS Modules

$requiredModules = @(

    "MicrosoftTeams",
    "ExchangeOnlineManagement",
    "Microsoft.Online.SharePoint.PowerShell",
    "Microsoft.Graph",
    "Pnp.PowerShell",
    "Az.Websites",
    "Az.Automation",
    "Az.Accounts",
    "Az.Resources",
    "AADInternals"

)

$installedModules = Get-InstalledModule

$missingModules = @()

foreach ($module in $requiredModules) {

    if ($installedModules.Name -contains $module) {

        Write-Host "$module is already installed." -ForegroundColor Green

    }

    else {

        Write-Host "$module is not installed." -ForegroundColor Yellow

        $missingModules += $module

    }

}

if ($missingModules) {

    foreach ($module in $missingModules) {

        Write-Host "Attempting to install $module..." -ForegroundColor Cyan

        Install-Module -Name $module

    }

    Write-Host "Finished installing required modules." -ForegroundColor Cyan
    
}

# Reload path environment variable
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine")

$nodePackages = @(

    "@pnp/cli-microsoft365",
    "@mermaid-js/mermaid-cli"

)

foreach ($package in $nodePackages) {

    npm -g install $package

}


# Reg Keys

$registryKeys = @(
    # Show search icon only
    [pscustomobject]@{
        RegistryKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search"
        ValueName = "SearchboxTaskbarMode"
        ValueData = "1"
        ValueType = "DWORD"
    },
    # Show file extensions
    [pscustomobject]@{
        RegistryKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        ValueName = "HideFileExt"
        ValueData = "0"
        ValueType = "DWORD"
    },
    # Open explorer to this PC
    [pscustomobject]@{
        RegistryKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        ValueName = "LaunchTo"
        ValueData = "1"
        ValueType = "DWORD"
    },
    # Disable item checkboxes
    [pscustomobject]@{
        RegistryKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        ValueName = "AutoCheckSelect"
        ValueData = "0"
        ValueType = "DWORD"
    },
    # Disable show frequent files
    # [pscustomobject]@{
    #     RegistryKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer"
    #     ValueName = "ShowFrequent"
    #     ValueData = "0"
    #     ValueType = "DWORD"
    # },
    # Disable show recent files
    # [pscustomobject]@{
    #     RegistryKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer"
    #     ValueName = "ShowRecent"
    #     ValueData = "0"
    #     ValueType = "DWORD"
    # },
    # Disable activity history
    [pscustomobject]@{
        RegistryKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"
        ValueName = "PublishUserActivities"
        ValueData = "0"
        ValueType = "DWORD"
    },
    # Rename This PC to Computer
    [pscustomobject]@{
        RegistryKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
        ValueName = "(Default)"
        ValueData = "Computer"
        ValueType = "String"
    },
    # Set Windows Terminal as default console application
    [pscustomobject]@{
        RegistryKey = "HKCU:\Console\%%Startup"
        ValueName = "DelegationConsole"
        ValueData = "{2EACA947-7F5F-4CFA-BA87-8F7FBEEFBE69}"
        ValueType = "String"
    },
    # Set Windows Terminal as default console application
    [pscustomobject]@{
        RegistryKey = "HKCU:\Console\%%Startup"
        ValueName = "DelegationTerminal"
        ValueData = "{E12CFF52-A866-4C77-9A90-F570A7AA2C6B}"
        ValueType = "String"
    },
    # Enable language bar but don't show
    [pscustomobject]@{
        RegistryKey = "HKCU:\SOFTWARE\Microsoft\CTF\LangBar"
        ValueName = "ExtraIconsOnMinimized"
        ValueData = "0"
        ValueType = "DWORD"
    },
    # Enable language bar but don't show
    [pscustomobject]@{
        RegistryKey = "HKCU:\SOFTWARE\Microsoft\CTF\LangBar"
        ValueName = "ShowStatus"
        ValueData = "3"
        ValueType = "DWORD"
    },
    # Enable language bar but don't show
    [pscustomobject]@{
        RegistryKey = "HKCU:\SOFTWARE\Microsoft\CTF\LangBar"
        ValueName = "Label"
        ValueData = "1"
        ValueType = "DWORD"
    },
    # Enable language bar but don't show
    [pscustomobject]@{
        RegistryKey = "HKCU:\SOFTWARE\Microsoft\CTF\LangBar"
        ValueName = "Transparency"
        ValueData = "255"
        ValueType = "DWORD"
    },
    # Remove chat from Taskbar
    [pscustomobject]@{
        RegistryKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        ValueName = "TaskbarMn"
        ValueData = "0"
        ValueType = "DWORD"
    },
    # Set Start Menu layout to more pins
    [pscustomobject]@{
        RegistryKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        ValueName = "Start_Layout"
        ValueData = "1"
        ValueType = "DWORD"
    },
    # Enable clipboard history
    [PSCustomObject]@{
        RegistryKey = "HKCU:\SOFTWARE\Software\Microsoft\Clipboard"
        ValueName   = "EnableClipboardHistory"
        ValueData   = "1"
        ValueType   = "DWORD"
    },
    # Disable desktop icons
    [pscustomobject]@{
        RegistryKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"
        ValueName = "NoDesktop"
        ValueData = "1"
        ValueType = "DWORD"
    },
    # Enable get me up to date
    [pscustomobject]@{
        RegistryKey = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"
        ValueName   = "IsExpedited"
        ValueData   = "1"
        ValueType   = "DWORD"
    },
    # Enable get me up to date
    [pscustomobject]@{
        RegistryKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
        ValueName   = "AllowMUUpdateService"
        ValueData   = "1"
        ValueType   = "DWORD"
    }
    
)

foreach ($item in $registryKeys) {

    #Check if Key exists
    $TestPath = Test-Path -Path $item.RegistryKey
    
        if ($TestPath -eq $true) {

            Write-Host "Registry Key already exists:" $item.RegistryKey -ForegroundColor Green

            #Get Key Properties

            $GetKeyProperties = Get-ItemProperty -Path $item.RegistryKey
            $KeyNameHelper = $item.ValueName
            $GetKeyValue = $GetKeyProperties.$KeyNameHelper

                #Check if Subkey already exists within Subkey
                if ($null -eq $GetKeyValue) {

                    Write-Host "Subkey does not exist:" $item.ValueName -ForegroundColor Yellow
                    Write-Host "Adding Subkey..." $item.ValueName "Value:" $item.ValueData 
                    New-ItemProperty -Path $item.RegistryKey -Name $item.ValueName -Value $item.ValueData -PropertyType $item.ValueType -Force

                    $GetKeyProperties = Get-ItemProperty -Path $item.RegistryKey
                    $KeyNameHelper = $item.ValueName
                    $GetKeyValue = $GetKeyProperties.$KeyNameHelper        

                }

                else {

                    Write-Host "Subkey already exists, checking values now..." -ForegroundColor Green
                }

                #Check if value of Subkey is set correctly
                if ($GetKeyValue -eq $item.ValueData) {

                    Write-Host "Key value is already set correctly. Key Name:" $item.ValueName "Value:" $GetKeyValue -ForegroundColor Green

                }

                else {

                    Write-Host "Key value not correct. Setting new value..." -ForegroundColor Yellow
                    Set-ItemProperty -Path $item.RegistryKey -Name $item.ValueName -Value $item.ValueData -Force

                }


        }

        else {

            Write-Host "Creating Registry Key:" $item.RegistryKey -ForegroundColor Yellow
            $ParentPath = Split-Path $item.RegistryKey -Parent
            $Leaf = Split-Path $item.RegistryKey -Leaf
            New-Item -Path $ParentPath -Name $Leaf -Force
            Write-Host "Setting Registry Key Values..." -ForegroundColor Yellow
            New-ItemProperty -Path $item.RegistryKey -Name $item.ValueName -Value $item.ValueData -PropertyType $item.ValueType -Force

        }

}


foreach ($registryKey in $registryKeys) {

    Set-ItemProperty -Path $registryKey.RegistryKey -Name $registryKey.ValueName -Value $registryKey.ValueData -Type $registryKey.ValueType -Force

}

Write-Host "Finished installing all apps and tools. Resarting computer in 30 seconds..." -ForegroundColor Green

shutdown.exe -r -t 30

# More features:

# Surface Pen Settings