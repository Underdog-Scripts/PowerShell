<#           
.NOTES
===========================================================================
Created with:    SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.173
Created on:      2/16/2022
Last Modified:   2/16/2022
Created by:      Gary Brunner
Organization:    Gary Brunner
Filename:        Timestamp.ps1 
===========================================================================
.DESCRIPTION
Provides daily SCCM reports
#>

<#

$ScriptName = "Timestamp"

$ComputerName = `
(Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName').ComputerName
$LogFile = ".\$ScriptName-$ComputerName.txt"

Start-Transcript -Path $Logfile

#>

#__________________________________________________________________________________________________


$TimeIs = Get-Date -Format "yyyyMMddhhmmss"

# Time description Y2022 M02 D16 H10 M55S 00
While($TimeIs -le 20220216110900){
    Start-Sleep -Seconds 10
    $TimeIs = Get-Date -Format "yyyyMMddhhmmss"
    Write-Host $TimeIs

}

Write-Host "Executing something"     


#__________________________________________________________________________________________________

<#

# Add file ending to log file
Write-Output "$ScriptName complete"

Stop-Transcript

#>