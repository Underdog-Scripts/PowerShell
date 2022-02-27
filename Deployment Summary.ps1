<#           
.NOTES
===========================================================================
Created with:    SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.173
Created on:      3/10/2021
Last Modified:   9/15/2021
Created by:      Gary Brunner
Organization:    Gary Brunner
Filename:        Deployment Summary.ps1 
===========================================================================
.DESCRIPTION
Provides list of SCCM deployments
#>

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

Start-Sleep -Seconds 2
Copy-Item ".\Deployments $TodayIs.csv" -Destination "\\IL7354-964\C$\Users\2brunnga\Documents\Reports\New\Deployments $TodayIs.csv"

Start-Sleep -Seconds 2
If (Test-Path "FileSystem::\\IL7354-964\C$\Users\2brunnga\Documents\Reports\New\Deployments $TodayIs.csv") {
    Remove-Item ".\Deployments $TodayIs.csv"
}