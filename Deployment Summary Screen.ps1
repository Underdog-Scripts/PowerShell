<#           
.NOTES
===========================================================================
Created with:    SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.173
Created on:      01/17/2022
Last Modified:   01/17/2022
Created by:      Gary Brunner
Organization:    Gary Brunner
Filename:        Deployment Summary Screen.ps1 
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
Sort-Object -Property "Start Date" -Descending | Out-GridView -Title "Deployment Summary for $TodayIs"
