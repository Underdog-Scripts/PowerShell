<#           
.NOTES
===========================================================================
Created with:    SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.173
Created on:      3/10/2021
Last Modified:   9/15/2021, 11.10.2021, 2.8.22
Created by:      Gary Brunner
Organization:    Gary Brunner
Filename:        Daily Reports.ps1 
===========================================================================
.DESCRIPTION
Provides daily SCCM reports
#>

$ScriptName = "Daily Reports 1Hr"

$ComputerName = `
(Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName').ComputerName
$LogFile = ".\$ScriptName-$ComputerName.txt"

Start-Transcript -Path $Logfile

#__________________________________________________________________________________________________

Function DeployReport() {
$TodayIs = Get-Date -UFormat "%m.%d.%y"
$DeploymentList = Get-WMIObject -Namespace root\sms\site_LW2 -class SMS_DeploymentSummary
$SelectedDeployment = $DeploymentList | Select @{Label= "Deployments" ;Expression={($_.SoftwareName)}},`
@{Label = "Complete %" ;Expression={($_.NumberSuccess/($_.NumberTargeted-$_.NumberUnknown))}},`
@{Label = "Target" ;Expression={($_.NumberTargeted)}},`
@{Label = "Complete" ;Expression={($_.NumberSuccess)}},`
@{Label = "Error" ;Expression={($_.NumberErrors)}},`
@{Label = "In Progress" ;Expression={($_.NumberInProgress)}},`
@{Label = "Unknown" ;Expression={($_.NumberUnknown)}},`
@{Label = "Start Date" ; Expression = {$([Management.ManagementDateTimeConverter]::toDateTime($_.DeploymentTime))}},`
@{Label = "Deadline"; Expression = { $([Management.ManagementDateTimeConverter]::toDateTime($_.EnforcementDeadline))}} |`
Sort-Object -Property "Start Date" -Descending | Export-Csv -LiteralPath ".\Deployments $TodayIs.csv" -NoClobber -Append -NoTypeInformation
}

Function OSReport() {
$TodayIs = Get-Date -UFormat "%m.%d.%y"
Start-Sleep -Seconds 2
Invoke-Sqlcmd "USE CM_LW2
Select v_r_system.Name0 as 'Computer Name',
v_r_system.User_Name0 as 'User Name',
v_GS_COMPUTER_SYSTEM.Manufacturer0 as 'Manufacturer',
v_GS_COMPUTER_SYSTEM.Model0 as 'Model',
OPSYS.Caption0 as 'Operating System',
OPSYS.Version0 as 'Version',
OPSYS.BuildNumber0 as 'Build Number',
v_r_system.BuildExt as 'Build Version',
OPSYS.LastBootUpTime0 as 'Last BootUp Time',
HWSCAN.LastHWScan as 'Last HW Inventory',
SWSCAN.LastScanDate as 'Last SW Inventory'
from v_r_system
left join v_GS_COMPUTER_SYSTEM on v_r_system.resourceid = v_GS_COMPUTER_SYSTEM.resourceid
left JOIN v_FullCollectionMembership fcm on v_r_system.ResourceID = fcm.ResourceID
left JOIN v_GS_OPERATING_SYSTEM OPSYS on v_r_system.ResourceID = OPSYS.ResourceID
LEFT JOIN v_GS_LastSoftwareScan SWSCAN on v_r_system.ResourceID = SWSCAN.ResourceID
LEFT JOIN v_GS_WORKSTATION_STATUS HWSCAN on v_r_system.ResourceID = HWSCAN.ResourceID
WHERE fcm.CollectionID='SMSDM003'
order by OPSYS.Caption0 desc" -QueryTimeout 120 | Export-Csv -Path ".\OS Report $TodayIs.csv" -NoClobber -NoTypeInformation -Append
}

Start-Sleep -Seconds (1 * 60 * 60)
OSReport
DeployReport

$TodayIs = Get-Date -UFormat "%m.%d.%y"
Start-Sleep -Seconds 2
If (Test-Path ".\Deployments $TodayIs.csv") {
    Copy-Item ".\Deployments $TodayIs.csv" -Destination "\\IL7354-964\C$\Users\2brunnga\Documents\Reports\New\Deployments $TodayIs.csv"
} Else {

Write-Output "Deployments $TodayIs.csv not found"
}

Start-Sleep -Seconds 2
If (Test-Path ".\OS Report $TodayIs.csv") {
Copy-Item ".\OS Report $TodayIs.csv" -Destination "\\IL7354-964\C$\Users\2brunnga\Documents\Reports\New\OS Report $TodayIs.csv"
} Else {

Write-Output "OS Report $TodayIs.csv not found"
}

#__________________________________________________________________________________________________

# Add file ending to log file
Write-Output "$ScriptName complete"

Stop-Transcript