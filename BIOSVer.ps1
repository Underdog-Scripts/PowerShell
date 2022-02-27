<#           
.NOTES
===========================================================================
Created with:    SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.173
Created on:      4/6/2021 3:04 PM
Created by:      Gary Brunner
Organization:    Gary Brunner
Filename:        RemoteDiag.ps1 
===========================================================================
.DESCRIPTION
Provides output from SetupDiag executed on remote system
#>

$DeviceList = Get-Content -Path ".\HP Update 2021-07 Systems.txt"

ForEach ($Device in $DeviceList) {
    If (Test-Connection -ComputerName $Device -Count 1 -Quiet) {
        Get-WmiObject -Class Win32_BIOS -ErrorAction SilentlyContinue -ComputerName $Device |`
        Select PSComputerName, Name |`
        Format-List |`
        Out-File -FilePath ".\HP Update 2021-07 Report.txt" -NoClobber -Append
    } Else {

        Write-Output "$Device Offline" | Out-File -FilePath ".\HP Update 2021-07 Report.txt" -NoClobber -Append
    }

}
