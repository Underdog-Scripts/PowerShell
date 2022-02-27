<#           
.NOTES
===========================================================================
Created with:    SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.173
Created on:      2/16/2022
Last Modified:   2/16/2022
Created by:      Gary Brunner
Organization:    Gary Brunner
Filename:        TimeUntil.ps1 
===========================================================================
.DESCRIPTION
Provides daily SCCM reports
#>

$ScriptName = "TimeUntil"

$LogFile = ".\$ScriptName.txt"

Start-Transcript -Path $Logfile

#__________________________________________________________________________________________________

# Time description: Year 2022 Month 01 Day 01 Hour 12 Minute 00 Second 00
$TimeTo = 20220216033000

$TimeIs = Get-Date -UFormat "%Y%m%d%H%M%S"
Write-Output "TimeIs = $TimeIs"
Write-Output "TimeTo = $TimeTo"
Write-Output ""

While($TimeIs -le $TimeTo){
    Start-Sleep -Seconds 3600
    $TimeIs = Get-Date -UFormat "%Y%m%d%H%M%S"

}

# Start-Process -FilePath "C:\Users\Gary\Documents\SCCMServer\Toys\MakeZip.exe"
# Write-Output "Executing - Start-Process -FilePath "C:\Users\Gary\Documents\SCCMServer\Toys\MakeZip.exe"

#__________________________________________________________________________________________________

# Add file ending to log file
Write-Output "$ScriptName complete"

Stop-Transcript
