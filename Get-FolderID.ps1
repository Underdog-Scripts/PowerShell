<#        
    .SYNOPSIS
     A brief summary of the commands in the file.

    .DESCRIPTION
    A detailed description of the commands in the file.

    .NOTES
    ========================================================================
         Windows PowerShell Source File 
         Created with SAPIEN Technologies PrimalScript 2019
         
         NAME: 
         
         AUTHOR: Altria , Altria
         DATE  : 12/28/2018
         
         COMMENT: 
         
    ==========================================================================
#>
Select [ContainerNodeID]
 ,[Name]
 ,[ObjectType]
 ,[ParentContainerNodeID]
 ,[FolderGuid]
 FROM [CM_P01].[dbo].[vSMS_Folders]
 WHERE Name like ‘Retired’
 
 $FolderID = XXXXXXXX
 $Sitecode = “XXX”
$CmServer = “XXXXXXXXXXX”
$Instancekeys = (Get-WmiObject -Namespace “ROOT\SMS\Site_$Sitecode” -ComputerName $CmServer -query “select InstanceKey from SMS_ObjectContainerItem where ObjectType=’6000′ and ContainerNodeID=’$FolderID'”).instanceKey
 foreach ($key in $Instancekeys)
 {
 (Get-WmiObject -Namespace “ROOT\SMS\Site_$Sitecode” -ComputerName $CmServer -Query “select * from SMS_Applicationlatest where ModelName = ‘$key'”).LocalizedDisplayName
 }
 
 
$SiteCode = ‘P01’

$FolderName = ‘Retired’

Import-Module (Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH –parent) ConfigurationManager.psd1)

set-location ${SiteCode}:

$FolderObj = Get-WmiObject -Class SMS_ObjectContainerNode -Namespace Root\SMS\Site_$SiteCode -filter “Name=’$FolderName‘” | foreach { $_.ContainerNodeID }

Get-WmiObject -Class SMS_ObjectContainerItem -Namespace Root\SMS\Site_$SiteCode -filter “ContainerNodeID=’$FolderObjID‘” | ForEach-Object {

    Get-CMDeviceCollection -CollectionID $_.InstanceKey | foreach { $_.Name }

}
