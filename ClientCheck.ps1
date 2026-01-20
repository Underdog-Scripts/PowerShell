<#
.Synopsis
   This script performs a Client Check
.DESCRIPTION
   This script performs a Client Check
.PARAMETER List
   None
.INPUTS
   Computer Name
.OUTPUTS
   Dialog display
.NOTES
   Version:        1.0
   Author:         GB
   Creation Date:  5.22.23
   Purpose/Change: Initial script development
   Modification Date: 5.23.23, 6.5.23, 6.7.23, 6.14.23, 6.15.23, 7.19.23,
                      7.20.23, 8.28.23, 9.1.23, 9.5.23, 9.21.23, 10.16.23,
                      10.4.24, 2.3.25, 2.10.25
   Purpose/Change: Modified for USPS
.EXAMPLE
   ./ClientCheck.ps1
.EXAMPLE
   ./ClientCheck.ps1
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SystemName
)

$SystemName = $SystemName.ToUpper()
$Remote = New-PSSession -ComputerName $SystemName
Invoke-Command -Session $Remote -ScriptBlock {

    # Log writing function
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogFile = "C:\Temp\$Using:SystemName-ClientCheck.log"
    Function Write-Text {
        param (
            [string]$Message
        )

        $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $LogMessage = "$TimeStamp - $Message"
        $LogMessage | Add-Content -Path $LogFile
    }

    # Check status of services ###################################################################################################
    Write-Host "Services --------------------------------"
    Write-Text "Services --------------------------------"
    $ServiceArray = "AdaptivaClient","BFE","BITS","CcmExec","LanManServer","LanManWorkstation","Riposte","WinMgmt"

    ForEach ($Service in $ServiceArray) {
        $ServiceStatus = (Get-Service -Name $Service -ErrorAction SilentlyContinue).Status
        $ServiceStartType = (Get-Service -Name $Service -ErrorAction SilentlyContinue).StartType
  
        If($ServiceStatus -eq "1") {$ServiceStatus = "Stopped"}
        If($ServiceStatus -eq "4") {$ServiceStatus = "Running"}
        If($ServiceStatus -eq "2") {$ServiceStatus = "Start Pending"}
        If($ServiceStartType -eq "2") {$ServiceStartType = "Automatic"}
        If($ServiceStartType -eq "3") {$ServiceStartType = "Manual"}
        If($ServiceStartType -eq "4") {$ServiceStartType = "Disabled"}
        If($Null -eq $ServiceStatus) {$ServiceStatus = "Null"}
        If($Null -eq $ServiceStartType) {$ServiceStartType = "Null"}

        If ($ServiceStatus -eq "Null" -or $ServiceStartType -eq "Null"){
        } ElseIf ($ServiceStartType -eq "Disabled" -or $ServiceStartType -eq "Manual" -or $ServiceStatus -eq "Stopped") {

            Write-Host "$Service - $ServiceStatus - $ServiceStartType"
            Write-Text -Message "$Service - $ServiceStatus - $ServiceStartType"
        } Else {
            Write-Host "$Service - $ServiceStatus - $ServiceStartType"
            Write-Text "$Service - $ServiceStatus - $ServiceStartType"
        }

    }

    # Check status of processes ##################################################################################################
    Write-Host "Processes -------------------------------"
    Write-Text "Processes -------------------------------"
    $ProcessArray = "AdaptivaClientService","Desktop","DesktopShell","Riposte"

    ForEach ($Process in $ProcessArray) {
        $ProcessName = (Get-Process -Name $Process -ErrorAction SilentlyContinue).Name
        If ($Process -ne "$ProcessName"){
            Write-Host "Inactive = $Process"
            Write-Text "Inactive = $Process"
        } Else {

            Write-Host "Active = $ProcessName"
            Write-Text "Active = $ProcessName"
        }

    }

    # Check status of registry ###################################################################################################
    Write-Host "Registry --------------------------------"
    Write-Text "Registry --------------------------------"
    $AdaptivaServer = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Adaptiva\Client\' -ErrorAction SilentlyContinue)."server_locator.server_name"
    If($Null -eq $AdaptivaServer) {
        $AdaptivaServer = "Null"
        Write-Host "Adaptiva Server = $AdaptivaServer"
        Write-Text "Adaptiva Server = $AdaptivaServer"
    } Else {

        Write-Host "Adaptiva Server = $AdaptivaServer"
        Write-Text "Adaptiva Server = $AdaptivaServer"
    }

    $AdaptivaSetup = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Adaptiva\Client\' -ErrorAction SilentlyContinue)."setup.server_host_name"
    If($Null -eq $AdaptivaSetup) {
        $AdaptivaSetup = "Null"
        Write-Host "Adaptiva Setup = $AdaptivaSetup"
        Write-Text "Adaptiva Setup = $AdaptivaSetup"
    } Else {

        Write-Host "Adaptiva Setup = $AdaptivaSetup"
        Write-Text "Adaptiva Setup = $AdaptivaSetup"
    }           

    $MPList = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\CCM\' -ErrorAction SilentlyContinue)."LookUpMPList"
    If($Null -eq $MPList) {
        $MPList = "Null"
        Write-Host "MP List = $MPList"
        Write-Text "MP List = $MPList"
    } Else {

        Write-Host "MP List = $MPList"
        Write-Text "MP List = $MPList"
    }

    $SMSSLP = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\CCM\' -ErrorAction SilentlyContinue)."SMSSLP"
    If($Null -eq $SMSSLP) {$SMSSLP = "Null"
        Write-Host "SMSSLP = $SMSSLP"
        Write-Text "SMSSLP = $SMSSLP"
    } Else {

        Write-Host "SMSSLP = $SMSSLP"
        Write-Text "SMSSLP = $SMSSLP"
    }

    # Check last ime SCCM logged actions #########################################################################################
    Write-Host "SCCM ------------------------------------"
    Write-Text "SCCM ------------------------------------"       
        $LastCacheCount = (Get-ChildItem -Path "C:\Windows\ccmcache" -ErrorAction SilentlyContinue).Count
    If($LastCacheCount -eq 0) {
    $LastCacheCount = "Null"
        Write-Host "Number of ccmcache folders = $LastCacheCount"
        Write-Text "Number of ccmcache folders = $LastCacheCount"
    } Else {

        Write-Host "Number of ccmcache folders = $LastCacheCount"
        Write-Text "Number of ccmcache folders = $LastCacheCount"
    }

    $LastCache = (Get-ChildItem -Path "C:\Windows\ccmcache" -ErrorAction SilentlyContinue).LastWriteTime | Sort-Object -Descending | Select-Object -First 1
    If($Null -eq $LastCache) {$LastCache = "Null"
        Write-Host "Last ccmcache folder write = $LastCache"
        Write-Text "Last ccmcache folder write = $LastCache"
    } Else {

        Write-Host "Last ccmcache folder write = $LastCache"
        Write-Text "Last ccmcache folder write = $LastCache"
    }

    $TaskLog = (Get-ChildItem -Path "C:\Windows\CCM\Logs\CITaskMgr.log" -ErrorAction SilentlyContinue).LastWriteTime
    If($Null -eq $TaskLog) {$TaskLog = "Null"
        Write-Host "Last CITaskMgr log write = $TaskLog"
        Write-Text "Last CITaskMgr log write = $TaskLog"
    } Else {

        Write-Host "Last CITaskMgr log write = $TaskLog"
        Write-Text "Last CITaskMgr log write = $TaskLog"
    }

    $TransferLog = (Get-ChildItem -Path "C:\Windows\CCM\Logs\DataTransferService.log" -ErrorAction SilentlyContinue).LastWriteTime
    If($Null -eq $TransferLog) {$TransferLog = "Null"
        Write-Host "Last DataTransferService log write = $TransferLog"
        Write-Text "Last DataTransferService log write = $TransferLog"
    } Else {

        Write-Host "Last DataTransferService log write = $TransferLog"
        Write-Text "Last DataTransferService log write = $TransferLog"
    }

    Write-Host "Other -----------------------------------"
    Write-Text "Other -----------------------------------"

    # Check drive free space #####################################################################################################
    $FreeSpace = (Get-WmiObject -Class Win32_LogicalDisk | Where-Object {$_.DeviceID -eq "C:"}).FreeSpace
    $FreeC = [math]::Round($FreeSpace/1GB,2)
    Write-Host "C: Drive Free Space = $FreeC GBs"
    Write-Text "C: Drive Free Space = $FreeC GBs"

    # Check version of PowerShell ################################################################################################
    $PSVersionIs = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine').PowerShellVersion
    Write-Host "PowerShell Version = $PSVersionIs"
    Write-Text "PowerShell Version = $PSVersionIs"

    # Check status of WMI ########################################################################################################
    $VerifyWMI = WinMgmt /VerifyRepository
    Write-Host "$VerifyWMI"
    Write-Text "$VerifyWMI"

    # Check reboot status ########################################################################################################
    If ((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing").RebootPending) {
        Return "Reboot Pending = $True"
        Write-Text "Reboot Pending = $True"
    } ElseIf ((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update").RebootRequired) {

        Return "Reboot Pending = $True"
        Write-Text "Reboot Pending = $True"
    } ElseIf ((Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager").PendingFileRenameOperations) {

        Return "Reboot Pending = $True"
        Write-Text "Reboot Pending = $True"
    } Else { 

        Return "Reboot Pending = $false"
        Write-Text "Reboot Pending = $false"
    }

}

$LogAction = Read-Host "Log file created at ClientCheck.log please select an option
1 = Delete log file
2 = Open log file `n"
# C:\Temp\$Using:SystemName-
If (1 -eq $LogAction){
    Write-Host "Delete"
} Else {
    Write-Host "Open"
}