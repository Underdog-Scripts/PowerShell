<#
.Synopsis
  This script is the Installation Template
.DESCRIPTION
   This script is the Installation Template
.PARAMETER List 
   None
.INPUTS
   None
.OUTPUTS
   Log file of actions result
.NOTES
   Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2021 v5.8.187
   Created on:   	8/24/21 3:36 PM
   Created by:   	GB
   Filename: 		HP Updates 2021.12.ps1
   Version:        	3.0
   Purpose/Change: 	Initial script development
   Modification Date: 2.14.20, 7.25.20, 7.28.20, 8.13.20, 11.16.20, 12.9.20, 2.10.21, 6.16.21, 8.24.21, 11.11.21, 12.10.21
.EXAMPLE
   ./HP Updates 2021.12.ps1
.EXAMPLE
   ./HP Updates 2021.12.ps1
#>

# Log file location and name
$ScriptName = "HP Updates 2021.12"

$ComputerName = `
(Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName').ComputerName
$LogFile = "C:\Windows\Installer\$ScriptName-$ComputerName.txt"

Start-Transcript -Path $Logfile

# Gather system information
$HDSize = Get-WmiObject -Class Win32_LogicalDisk | Where-Object {$_.Name -eq "C:"} | Select @{n="Size";e={[math]::Round($_.Size/1GB,2)}}
    $HDSize = $HDSize -Replace '@{Size=', ''
    $HDSize = $HDSize -Replace '}', ''

$HDFree = Get-WmiObject -Class Win32_LogicalDisk | Where-Object {$_.Name -eq "C:"} | Select @{n="FreeSpace";e={[math]::Round($_.FreeSpace/1GB,2)}}
    $HDFree = $HDFree -Replace '@{FreeSpace=', ''
    $HDFree = $HDFree -Replace '}', ''

$Video = Get-WmiObject -Class Win32_VideoController -ErrorAction SilentlyContinue | Select @{Label="Name";Expression={($_.Caption)}}
    $Video = $Video -Replace '@{Name=', ''
    $Video = $Video -Replace '}', ''

$CcmExec = Get-Service -Name "CcmExec" | Select-Object Status -ErrorAction SilentlyContinue
    $CcmExec = $CcmExec -Replace '@{Status=', ''
    $CcmExec = $CcmExec -Replace '}', ''

$Bits = Get-Service -Name "Bits" | Select-Object Status -ErrorAction SilentlyContinue
    $Bits = $Bits -Replace '@{Status=', ''
    $Bits = $Bits -Replace '}', ''

$WuauServ = Get-Service -Name "WuauServ" | Select-Object Status -ErrorAction SilentlyContinue
    $WuauServ = $WuauServ -Replace '@{Status=', ''
    $WuauServ = $WuauServ -Replace '}', ''

$WinMgmt = Get-Service -Name "WinMgmt" | Select-Object Status -ErrorAction SilentlyContinue
    $WinMgmt = $WinMgmt -Replace '@{Status=', ''
    $WinMgmt = $WinMgmt -Replace '}', ''

$RemoteRegistry = Get-Service -Name "RemoteRegistry" | Select-Object Status -ErrorAction SilentlyContinue
    $RemoteRegistry = $RemoteRegistry -Replace '@{Status=', ''
    $RemoteRegistry = $RemoteRegistry -Replace '}', ''

$BiosSerial = `
(Get-WmiObject Win32_BIOS).SerialNumber

$BiosVer = `
(Get-WmiObject Win32_BIOS).SMBIOSBIOSVersion

$BuildLabEx = `
(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').BuildLabEx

$CompMfr = `
(Get-WmiObject Win32_ComputerSystem).Manufacturer

$CompModel = `
(Get-WmiObject Win32_ComputerSystem).Model

$CompositionEditionID = `
(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').CompositionEditionID

$ComputerName = `
(Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName').ComputerName

$CPU = `
(Get-WmiObject Win32_OperatingSystem).OSArchitecture

$CurrentVersion = `
(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').CurrentVersion

$ProductName = `
(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').ProductName

$ReleaseID = `
(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').ReleaseId

# Add system information to log file
Write-Output "BIOS Serial Number: $BiosSerial"
Write-Output "BIOS Version: $BiosVer"
Write-Output "BITS Status: $Bits"
Write-Output "CcmExec Status: $CcmExec"
Write-Output "Composition Edition: $CompositionEditionID"
Write-Output "Computer Manufacturer: $CompMfr"
Write-Output "Computer Model: $CompModel"
Write-Output "Computer Name: $ComputerName"
Write-Output "CPU: $CPU"
Write-Output "Current Version: $CurrentVersion"
Write-Output "Domain: $env:USERDOMAIN"
Write-Output "Hard Drive Size: $HDSize"
Write-Output "Hard Drive Free Space: $HDFree"
Write-Output "Product Name: $ProductName"
Write-Output "ScriptName: $ScriptName"
Write-Output "Release ID: $ReleaseID"
Write-Output "Remote Registry Status: $RemoteRegistry"
Write-Output "User Name: $env:USERNAME"
Write-Output "Video: $Video"
Write-Output "Windows BuildLabEx: $BuildLabEx"
Write-Output "WinMgmt Status: $WinMgmt"
Write-Output "WuauServ Status: $WuauServ"
Write-Output '========================= End of system information ========================='
Write-Output ' '

###########################################################################################
# Customization for this script

# Select the model to assign command line switches
Switch ($CompModel){

    #'HP Compaq Pro 6300 MT' {
    #	Write-Output "Detected HP Compaq Pro 6300 MT"
    #}

    #'HP EliteBook 820 G1' {
    #    Write-Output "Detected HP EliteBook 820 G1"
    #}

    'HP EliteBook 820 G3' {
        Write-Output "Detected HP EliteBook 820 G3"
        Start-Process -FilePath "$PSScriptRoot\SP113020\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP113020\install.cmd"
    }

    'HP EliteBook 830 G5' {
        Write-Output "Detected HP EliteBook 830 G5"
        Start-Process -FilePath "$PSScriptRoot\SP113020\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP113020\install.cmd"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP133131\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP133131\install.cmd"
    }   

    'HP EliteBook 830 G6' {
        Write-Output "Detected HP EliteBook 830 G6"
        Start-Process -FilePath "$PSScriptRoot\SP113020\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP113020\install.cmd"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP123471\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP123471\install.cmd"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP134448\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP134448\install.cmd"
    }
    
    'HP EliteBook 830 G8 Notebook PC' {
        Write-Output "Detected HP EliteBook 830 G8 Notebook PC"
        Start-Process -FilePath "$PSScriptRoot\SP112603\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP112603\install.cmd"
    }

    'HP EliteBook Revolve 810 G3' {
        Write-Output "Detected HP EliteBook Revolve 810 G3"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP118367\igxpin.exe -s -Report C:\Windows\Installer\IntelGFX.log"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP118367\igxpin.exe -s -report C:\Windows\Installer\IntelGFX.log"
    }

    'HP EliteDesk 800 G3 DM 35W' {
        Write-Output "HP EliteDesk 800 G3 DM 35W"
        Start-Process -FilePath "$PSScriptRoot\SP113020\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP113020\install.cmd"
    }

    'HP ProBook 650 G1' {
        Write-Output "Detected HP ProBook 650 G1"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP118367\igxpin.exe -s -Report C:\Windows\Installer\IntelGFX.log"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP118367\igxpin.exe -s -report C:\Windows\Installer\IntelGFX.log"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP133190\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP133190\install.cmd"
    }
    
    'HP ProBook 650 G2' {
        Write-Output "Detected HP ProBook 650 G2"
        Start-Process -FilePath "$PSScriptRoot\SP113020\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP113020\install.cmd"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP115295\igxpin.exe -s -Report C:\Windows\Installer\IntelGFX.log"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP115295\igxpin.exe -s -report C:\Windows\Installer\IntelGFX.log"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP133190\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP133190\install.cmd"
    }

    'HP ProBook 650 G4' {
        Write-Output "Detected HP ProBook 650 G4"
        Start-Process -FilePath "$PSScriptRoot\SP113020\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP113020\install.cmd"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP115295\igxpin.exe -s -Report C:\Windows\Installer\IntelGFX.log"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP115295\igxpin.exe -s -report C:\Windows\Installer\IntelGFX.log"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP134238\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP134238\install.cmd"
    }

    'HP ProBook 650 G5' {
        Write-Output "Detected HP ProBook 650 G5"
        Start-Process -FilePath "$PSScriptRoot\SP113020\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP113020\install.cmd"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP123471\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP123471\install.cmd"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP134238\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP134238\install.cmd"
    }

    'HP ProBook 650 G8 Notebook PC' {
        Write-Output "Detected HP ProBook 650 G8 Notebook PC"
        Start-Process -FilePath "$PSScriptRoot\SP113020\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP113020\install.cmd"
    }

    'HP ProBook x360 11 G6 EE' {
        Write-Output "Detected HP ProBook x360 11 G6 EE"
        Start-Process -FilePath "$PSScriptRoot\SP113020\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP113020\install.cmd"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP123471\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP123471\install.cmd"
    }


    'HP ProDesk 600 G1 TWR' {
        Write-Output "Detected HP ProDesk 600 G1 TWR"
    }

    'HP ProDesk 600 G2 MT' {
        Write-Output "Detected HP ProDesk 600 G2 MT"
    }
	
    'HP ProDesk 600 G3 MT' {
        Write-Output "Detected HP ProDesk 600 G3 MT"
    }

    'HP ProDesk 600 G4 MT' {
        Write-Output "Detected HP ProDesk 600 G4 MT"
        Start-Process -FilePath "$PSScriptRoot\SP113020\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP113020\install.cmd"
    }
    
    'HP ProDesk 600 G4 SFF' {
        Write-Output "Detected HP ProDesk 600 G4 SFF"
        Start-Process -FilePath "$PSScriptRoot\SP113020\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP113020\install.cmd"
    }
    
    'HP ProDesk 600 G5 MT' {
        Write-Output "Detected HP ProDesk 600 G5 MT"
        Start-Process -FilePath "$PSScriptRoot\SP113020\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP113020\install.cmd"
    }    

    'HP Z4 G4 Workstation' {
        Write-Output "Detected HP Z4 G4 Workstation"
        Start-Process -FilePath "$PSScriptRoot\SP112211\irmfuu_7.1.19_win64.exe"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP112211\irmfuu_7.1.19_win64.exe"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP132922\Setup.exe -Install"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP132922\Setup.exe -Install"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP135935\ThunderboltSoftwareInstaller.exe /install /quiet /Log C:\Windows\Installer\Thunder.txt"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP135935\ThunderboltSoftwareInstaller.exe /install /quiet /Log C:\Windows\Installer\Thunder.txt"
    }
    
    'HP Z8 G4 Workstation' {
        Write-Output "Detected HP Z8 G4 Workstation"
        Start-Process -FilePath "$PSScriptRoot\SP112211\irmfuu_7.1.19_win64.exe"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP112211\irmfuu_7.1.19_win64.exe"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP132922\Setup.exe -Install"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP132922\Setup.exe -Install"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP135935\ThunderboltSoftwareInstaller.exe /install /quiet /Log C:\Windows\Installer\Thunder.txt"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP135935\ThunderboltSoftwareInstaller.exe /install /quiet /Log C:\Windows\Installer\Thunder.txt"
    }    

    #'HP Z420 Workstation' {
    #    Write-Output "Detected HP Z420 Workstation"
    #}

    'HP Z440 Workstation' {
        Write-Output "Detected HP Z440 Workstation"
        Start-Process -FilePath "$PSScriptRoot\SP132922\Setup.exe -Install"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP132922\Setup.exe -Install"
    }

    #'HP ZBook 17' {
    #    Write-Output "Detected HP ZBook 17"
    #}

    'HP ZBook 17 G2' {
        Write-Output "Detected HP ZBook 17 G2"
        Start-Process -FilePath "$PSScriptRoot\SP113697\Setup.exe -Install"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP113697\Setup.exe -Install"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP118367\igxpin.exe -s -Report C:\Windows\Installer\IntelGFX.log"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP118367\igxpin.exe -s -report C:\Windows\Installer\IntelGFX.log"
    }

    'HP ZBook 17 G3' {
        Write-Output "Detected HP ZBook 17 G3"
        Start-Process -FilePath "$PSScriptRoot\SP113020\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP113020\install.cmd"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP115295\igxpin.exe -s -Report C:\Windows\Installer\IntelGFX.log"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP115295\igxpin.exe -s -report C:\Windows\Installer\IntelGFX.log"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP133131\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP133131\install.cmd"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP133190\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP133190\install.cmd"
    }
    
    'HP ZBook 17 G5' {
        Write-Output "Detected HP ZBook 17 G5"
        Start-Process -FilePath "$PSScriptRoot\SP113020\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP113020\install.cmd"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP115295\igxpin.exe -s -Report C:\Windows\Installer\IntelGFX.log"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP115295\igxpin.exe -s -report C:\Windows\Installer\IntelGFX.log"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP133131\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP133131\install.cmd"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP134238\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP134238\install.cmd"
    }

    'HP ZBook 17 G6' {
        Write-Output "Detected HP ZBook 17 G6"
        Start-Process -FilePath "$PSScriptRoot\SP113020\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP113020\install.cmd"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP123471\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP123471\install.cmd"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\SP134448\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP134448\install.cmd"
    }

    'HP ZBook Fury 17 G7 Mobile Workstation' {
        Write-Output "Detected HP ZBook Fury 17 G7 Mobile Workstation"
        Start-Process -FilePath "$PSScriptRoot\SP113020\install.cmd"
        Write-Output "Executing - Start-Process -FilePath $PSScriptRoot\SP113020\install.cmd"
    }
    
    Default {
    $CompModel = "No Model"
    }

}

# Update Log file
If ($CompModel -ne "No Model") {
    Write-Output "$ScriptName Applied"

} Else {
	
	Write-Output "No HP update required for model $CompModel"
}

# Add file ending to log file
Write-Output "$ScriptName complete"

Stop-Transcript
