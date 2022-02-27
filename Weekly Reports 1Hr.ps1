<#           
.NOTES
===========================================================================
Created with:    SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.173
Created on:      7/14/2021 5:50 PM
Modified on:     11/27/2021 7:32 PM, 2.8.22 3:32 PM
Created by:      Gary Brunner
Filename:        Weekly Reports.ps1 
===========================================================================
.DESCRIPTION
Provides All Weekly Reports for environment
#>

Start-Sleep -Seconds (8 * 60 * 60)
$ScriptName = "Weekly Reports 1Hr"

$ComputerName = `
(Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName').ComputerName
$LogFile = ".\$ScriptName-$ComputerName.txt"

Start-Transcript -Path $Logfile

#__________________________________________________________________________________________________

$TodayIs = Get-Date -UFormat "%m.%d.%y"
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

Start-Sleep -Seconds 1
Invoke-Sqlcmd "USE CM_LW2
SELECT DISTINCT sys.Netbios_Name0 as 'Machine Name',
sys.User_Name0 as 'User Name', 
sys.AD_Site_Name0 as 'Site Name',
os.Caption0 as 'Operating System',
cs.Manufacturer0 as 'Manufacturer',
proc2.SystemType0 as 'Operating System Type',
os.Version0 as 'Operating System Version Number',
bio.Name0 as 'BIOS Name',
bio.SMBIOSBIOSVersion0 as 'BIOS Version',
ld.FreeSpace0 as 'Disk Space Available',
cpu.Name0 as 'Processor Information',
mem.Capacity0 as 'RAM (MB)',
vg.Name0 as 'Video Graphics Adapter',
vg.AdapterRAM0 as 'Video Memory',
vg.InstalledDisplayDrivers0 as 'Video Graphics Adapter Driver Name',
ENC.SMBIOSAssetTag0 as 'Asset Tag',
ENC.SerialNumber0 as 'Serial Number',
Case
When D_MachineType.Desktops = 1  THEN 'Desktop'
When D_MachineType.Laptops = 1 THEN 'Laptop'
Else 'Unknown'
end as 'Type',
DateDiff(Day,WK.LastHWScan,GetDate()) as 'Days since HW Inv'
FROM 
v_R_System as sys
left join 
(Select * from v_GS_SYSTEM_ENCLOSURE WHERE V_GS_SYSTEM_ENCLOSURE.Tag0 = 'System Enclosure 0') ENC on sys.ResourceID = ENC.ResourceID
left join v_GS_WORKSTATION_STATUS WK on sys.ResourceID = WK.ResourceID
left join v_gs_video_controller vg on sys.resourceid = vg.resourceid
left join v_GS_COMPUTER_SYSTEM cs on sys.resourceid = cs.resourceid
LEFT JOIN 
(Select ResourceID, Max(Name0) as Name0, NormSpeed0, Is64Bit0, AddressWidth0, DataWidth0 
from v_GS_PROCESSOR 
group by ResourceID, NormSpeed0, Is64Bit0, AddressWidth0, DataWidth0) CPU on SYS.ResourceID = CPU.ResourceID
left join v_gs_pc_bios bio on sys.resourceid = bio.resourceid
LEFT join (select * from v_GS_LOGICAL_DISK
where 
DriveType0 = 3 and DeviceID0 = 'C:') as ld on SYS.ResourceID = ld.ResourceID
left join v_GS_COMPUTER_SYSTEM PROC2 on sys.ResourceID = PROC2.ResourceID
left join v_GS_OPERATING_SYSTEM OS on sys.ResourceID = OS.ResourceID
left join v_GS_PHYSICAL_MEMORY MEM on sys.ResourceID = MEM.ResourceID
LEFT JOIN (SELECT ResourceID,
    'Laptops' = CASE
        WHEN (ChassisTypes0 = '8')
            OR (ChassisTypes0 = '9') 
            OR (ChassisTypes0 = '10')  
            OR (ChassisTypes0 = '11') 
            OR (ChassisTypes0 = '12') 
            OR (ChassisTypes0 = '14') 
            OR (ChassisTypes0 = '18') 
            OR (ChassisTypes0 = '21') THEN 1
       ELSE 0 END,
    'Desktops' = CASE
        WHEN (ChassisTypes0 = '8') 
            OR (ChassisTypes0 = '9') 
            OR (ChassisTypes0 = '10') 
            OR (ChassisTypes0 = '11') 
            OR (ChassisTypes0 = '12') 
            OR (ChassisTypes0 = '14') 
            OR (ChassisTypes0 = '18') 
            OR (ChassisTypes0 = '21') THEN 0
        WHEN (ChassisTypes0 = '17') 
            OR (ChassisTypes0 = '23') THEN 0
       ELSE 1 END
FROM V_GS_SYSTEM_ENCLOSURE
WHERE V_GS_SYSTEM_ENCLOSURE.Tag0 = 'System Enclosure 0') D_MachineType ON sys.ResourceID = D_MachineType.ResourceID
left join v_FullCollectionMembership vfm on vfm.ResourceID=sys.ResourceID
Where 
sys.Operating_System_Name_and0 NOT LIKE '%SERVER%'
AND vfm.CollectionID='SMSDM003'
Order by sys.Netbios_Name0" -QueryTimeout 120 | Export-Csv -Path ".\Hardware Report Full List of Hardware $TodayIs.csv" -NoClobber -NoTypeInformation -Append

Start-Sleep -Seconds 1
Invoke-Sqlcmd "USE CM_LW2
SELECT 
SYS.Netbios_Name0 as 'Netbios Name', 
INST.SMS_Installed_Sites0 as 'Installed Site Code', 
SYS.AD_Site_Name0 as 'AD Site Name',
SYS.User_Domain0 as 'User Domain',
SYS.User_Name0 as 'User Name',
SYS.Resource_Domain_OR_Workgr0 as 'Computer Domain',
HWSCAN.LastHWScan as 'Last Hardware Scan',
  DateDiff(Day, 
     CASE WHEN IsNULL(HWSCAN.LastHWScan,'1/1/1980') > IsNULL(SWSCAN.LastScanDate,'1/1/1980') 
          THEN HWSCAN.LastHWScan
          ELSE SWSCAN.LastScanDate
     END,
GetDate()) AS 'Number of Days Since Inventoried',
SWSCAN.LastScanDate as 'Last Software Inventory',
SWSCAN.LastCollectedFileScanDate as 'Last File Collection'
FROM v_R_System SYS
LEFT JOIN v_GS_LastSoftwareScan SWSCAN on SYS.ResourceID = SWSCAN.ResourceID
LEFT JOIN v_GS_WORKSTATION_STATUS HWSCAN on SYS.ResourceID = HWSCAN.ResourceID
LEFT OUTER JOIN v_RA_System_SMSInstalledSites INST ON SYS.ResourceID = INST.ResourceID 
WHERE 
   DateDiff(Day, 
     CASE WHEN IsNULL(HWSCAN.LastHWScan,'1/1/1980') > IsNULL(SWSCAN.LastScanDate,'1/1/1980') 
          THEN HWSCAN.LastHWScan
          ELSE SWSCAN.LastScanDate
     END
      ,GetDate()) >= 7
AND Client0=1
Order By INST.SMS_Installed_Sites0, SYS.Netbios_Name0" -QueryTimeout 120 | Export-Csv -Path ".\CR026 Computers Not Actively Reporting $TodayIs.csv" -NoClobber -NoTypeInformation -Append

Start-Sleep -Seconds 1
Invoke-Sqlcmd "USE CM_LW2
declare @CollectionID varchar(10)
set @CollectionID = 'SMS00001' --ALL

declare @locale varchar(10)
set @locale = 'en-US'
     
declare @RscID int
set @RscID = 16777218

declare @lcid as int set @lcid = dbo.fn_LShortNameToLCID(@locale)
declare @ci table(CI_ID int primary key, CI_UniqueID nvarchar(256), Title nvarchar(512), ArticleID nvarchar(64), BulletinID nvarchar(64), InfoURL nvarchar(512), DatePosted DateTime)
insert @ci(CI_ID, CI_UniqueID, Title, ArticleID, BulletinID, InfoURL, DatePosted)
select ci.CI_ID, ci.CI_UniqueID, ci.Title, ci.ArticleID, ci.BulletinID, ci.InfoURL, ci.DatePosted
from fn_UpdateInfo(@lcid) ci
left join v_CICategories_All ven on ven.CI_ID=ci.CI_ID and ven.CategoryInstanceID= 0 --@VendorID 
left join v_CICategories_All cls on cls.CI_ID=ci.CI_ID and cls.CategoryInstanceID= 0 --@ClassID 
left join v_CICategories_All prd on prd.CI_ID=ci.CI_ID and prd.CategoryInstanceID= 0 --@ProductID
where ci.CIType_ID in (1,8) and ci.IsExpired=0 
order by ci.DateRevised desc
select distinct
DatePosted,
BulletinID 'Bulletin ID', 
ArticleID 'Article ID', 
Installed=NumPresent , 
Reboot=NumPending,  --Needs Reboot
Missing=(NumMissing - NumPending) , --Needed
InScope=NumPresent + NumMissing, --InScope 
'% Complete'= Str((NumPresent * 100.00 / (NumPresent + NumMissing)),6,2),
'% Complete after reboot' = Str((((NumPresent + NumPending) * 100.00) / (NumPresent + NumMissing)),6,2),
Title
from @ci ui
left join fn_CICategoryInfo_All(@lcid) ven on ven.CI_ID=ui.CI_ID and ven.CategoryTypeName='Company' 
left join fn_CICategoryInfo_All(@lcid) cls on cls.CI_ID=ui.CI_ID and cls.CategoryTypeName='UpdateClassification' 
left join v_CITargetedCollections tc on tc.CI_ID=ui.CI_ID and tc.CollectionID= @CollectionID --@CollID
left join v_UpdateSummaryPerCollection cs on cs.CI_ID=ui.CI_ID and cs.CollectionID= @CollectionID --@CollID 
left join v_CICategories_All CIALL on ui.CI_ID=CIALL.CI_ID
where 
-- CIALL.CategoryInstanceID = 31 AND
((NumPresent > 0)  OR (NumMissing > 0)) AND
DatePosted >= '2010' AND
Title NOT LIKE '%Security Only%'
order by 
DatePosted,
BulletinID, 
ArticleID" -QueryTimeout 180| Export-Csv -Path ".\CR063 Patch Levels $TodayIs.csv" -NoClobber -NoTypeInformation -Append

Start-Sleep -Seconds 1
Invoke-Sqlcmd "USE CM_LW2
declare @Vendor varchar(20);
declare @UpdateClass varchar(20);

Set @Vendor = 'Microsoft';
Set @UpdateClass = 'Security Updates';

select 
SYS.Netbios_Name0 as 'Machine Name',
ui.BulletinID,
ui.ArticleID,
ui.Title as 'Title',
DateDiff(Day,WK.LastHWScan,GetDate()) as 'Days since last Hardware Scan'
from 
v_r_System SYS
left join v_GS_WORKSTATION_STATUS WK on SYS.ResourceID = WK.ResourceID
left join v_UpdateComplianceStatus css on SYS.ResourceID = css.ResourceID
join v_UpdateInfo ui on ui.CI_ID=css.CI_ID
join v_CICategories_All catall on catall.CI_ID=ui.CI_ID 
join v_CategoryInfo catinfo on catall.CategoryInstance_UniqueID = catinfo.CategoryInstance_UniqueID and catinfo.CategoryTypeName='Company' 
join v_CICategories_All catall2 on catall2.CI_ID=ui.CI_ID 
join v_CategoryInfo catinfo2 on catall2.CategoryInstance_UniqueID = catinfo2.CategoryInstance_UniqueID and catinfo2.CategoryTypeName='UpdateClassification' 
where  
(css.Status=2)
and (@Vendor = '' or catinfo.CategoryInstanceName = @Vendor)
and (@UpdateClass = '' or catinfo2.CategoryInstanceName = @UpdateClass)
order by 
SYS.Netbios_Name0,
ui.BulletinID,
ui.ArticleID" -QueryTimeout 600 | Export-Csv -Path ".\CR064 Computers Missing Patches $TodayIs.csv" -NoClobber -NoTypeInformation -Append

Start-Sleep -Seconds 1
Invoke-Sqlcmd "USE CM_LW2
select
SYS.Name0 as 'Name', 
SYS.AD_Site_Name0 as 'AD Site Name',
SYS.Resource_Domain_OR_Workgr0 as 'Computer Domain',
SYS.User_Name0 as 'User Name',
WK.LastHWScan as 'Last Hardware Scan',
DateDiff(Day,WK.LastHWScan,GetDate()) as 'Days Since HW Inv',
Case
WHEN (substring(AccountName0,charindex('Domain=',Accountname0)+8,(charindex('Name=',Accountname0)-charindex('Domain=',Accountname0)-10))) = SYS.Name0 THEN
  (substring(AccountName0,len(AccountName0)-charindex('`"',reverse(AccountName0),2)+2,charindex('`"',reverse(AccountName0),2)-2))
ELSE
(substring(AccountName0,charindex('Domain=',Accountname0)+8,(charindex('Name=',Accountname0)-charindex('Domain=',Accountname0)-10))) + '\' +
  (substring(AccountName0,len(AccountName0)-charindex('`"',reverse(AccountName0),2)+2,charindex('`"',reverse(AccountName0),2)-2)) 
End as 'Local Admin Group Member'
from
v_R_System AS SYS 
left JOIN v_GS_WORKSTATION_STATUS WK on SYS.ResourceID = WK.ResourceID
left join
(select *
from v_GS_Local_Admins 
where 
  (AccountName0 not like '%Administrator%' 
    AND AccountName0 not like '%Domain Admins%')
) LocAdm ON  SYS.ResourceID = LocAdm.ResourceID 
where
DateDiff(Day,WK.LastHWScan,GetDate()) <= '30'
and SYS.Operating_System_Name_and0 NOT LIKE '%Server%'
and SYS.Client0 = '1' 
and WK.LastHWScan IS NOT NULL
order by
SYS.Name0" -QueryTimeout 120 | Export-Csv -Path ".\CR066 Admin User Accounts $TodayIs.csv" -NoClobber -NoTypeInformation -Append

Start-Sleep -Seconds 1
Invoke-Sqlcmd "USE CM_LW2
SELECT
SYS.Netbios_Name0 as 'Netbios Name', 
INST.SMS_Installed_Sites0 as 'SMS Installed Sites', 
SYS.AD_Site_Name0 as 'AD Site Name',
SYS.User_Domain0 as 'User Domain', 
SYS.User_Name0 as 'User Name',
' ' as MYID,
' ' as 'Cost Center',
OS.Caption0 'OS',
OS.CSDVersion0 'Service Pack',
SYS.Resource_Domain_OR_Workgr0 as 'Computer Domain', 
HWSCAN.LastHWScan as 'Last Hardware Scan',
DateDiff( 
Day, 
CASE WHEN IsNULL(HWSCAN.LastHWScan,'1/1/1980') > IsNULL(SWSCAN.LastScanDate,'1/1/1980') 
     THEN HWSCAN.LastHWScan
     ELSE SWSCAN.LastScanDate
END,
GetDate() ) AS C053,
SWSCAN.LastScanDate
FROM v_R_System SYS
LEFT JOIN v_GS_LastSoftwareScan SWSCAN on SYS.ResourceID = SWSCAN.ResourceID
LEFT JOIN v_GS_WORKSTATION_STATUS HWSCAN on SYS.ResourceID = HWSCAN.ResourceID
LEFT OUTER JOIN v_RA_System_SMSInstalledSites INST ON SYS.ResourceID = INST.ResourceID 
LEFT JOIN v_GS_OPERATING_SYSTEM OS on SYS.ResourceID = OS.ResourceID
WHERE 
DateDiff(
Day, 
CASE WHEN IsNULL(HWSCAN.LastHWScan,'1/1/1980') > IsNULL(SWSCAN.LastScanDate,'1/1/1980') 
     THEN HWSCAN.LastHWScan
     ELSE SWSCAN.LastScanDate 
END,
GetDate() ) >= 7
AND 
Client0=1
Order By INST.SMS_Installed_Sites0, SYS.Netbios_Name0" -QueryTimeout 120 | Export-Csv -Path ".\CR067 Computer Utilization (CR026) $TodayIs.csv" -NoClobber -NoTypeInformation -Append

Start-Sleep -Seconds 1
Invoke-Sqlcmd "USE CM_LW2
SELECT SYS.Name as 'Name',
SYS.SiteCode as 'Site Code',
LDISK.Description0 as 'Description',
LDISK.DeviceID0 as 'Device ID',
LDISK.VolumeName0 as 'Volume Name',
LDISK.FileSystem0 as 'File system',
LDISK.Size0 as 'Size',
LDISK.FreeSpace0 as 'Free Space'
FROM v_FullCollectionMembership SYS
join v_GS_LOGICAL_DISK LDISK on SYS.ResourceID = LDISK.ResourceID
WHERE 
  LDISK.DriveType0 = 3 AND
  LDISK.FreeSpace0 < 4000 AND
  LDISK.DeviceID0 = 'C:' AND
SYS.CollectionID = 'LW20001A'
ORDER BY SYS.Name" -QueryTimeout 120 | Export-Csv -Path ".\CR068 Disk Space Usage $TodayIs.csv" -NoClobber -NoTypeInformation -Append

Start-Sleep -Seconds 2
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

Start-Sleep -Seconds 2
If (Test-Path ".\Hardware Report Full List of Hardware $TodayIs.csv") {
    Copy-Item ".\Hardware Report Full List of Hardware $TodayIs.csv" -Destination "\\IL7354-964\C$\Users\2brunnga\Documents\Reports\New\Hardware Report Full List of Hardware $TodayIs.csv"
} Else {

    Write-Output "Hardware Report Full List of Hardware $TodayIs.csv not found"
}

Start-Sleep -Seconds 2
If (Test-Path ".\CR026 Computers Not Actively Reporting $TodayIs.csv") {
    Copy-Item ".\CR026 Computers Not Actively Reporting $TodayIs.csv" -Destination "\\IL7354-964\C$\Users\2brunnga\Documents\Reports\New\CR026 Computers Not Actively Reporting $TodayIs.csv"
} Else {

    Write-Output "CR026 Computers Not Actively Reporting $TodayIs.csv not found"
}

Start-Sleep -Seconds 2
If (Test-Path ".\CR063 Patch Levels $TodayIs.csv") {
    Copy-Item ".\CR063 Patch Levels $TodayIs.csv" -Destination "\\IL7354-964\C$\Users\2brunnga\Documents\Reports\New\CR063 Patch Levels $TodayIs.csv"
} Else {

    Write-Output "CR063 Patch Levels $TodayIs.csv not found"
}

Start-Sleep -Seconds 2
If (Test-Path ".\CR064 Computers Missing Patches $TodayIs.csv") {
    Copy-Item ".\CR064 Computers Missing Patches $TodayIs.csv" -Destination "\\IL7354-964\C$\Users\2brunnga\Documents\Reports\New\CR064 Computers Missing Patches $TodayIs.csv"
} Else {

    Write-Output "CR064 Computers Missing Patches $TodayIs.csv not found"
}

Start-Sleep -Seconds 2
If (Test-Path ".\CR066 Admin User Accounts $TodayIs.csv") {
    Copy-Item ".\CR066 Admin User Accounts $TodayIs.csv" -Destination "\\IL7354-964\C$\Users\2brunnga\Documents\Reports\New\CR066 Admin User Accounts $TodayIs.csv"
} Else {

    Write-Output "CR066 Admin User Accounts $TodayIs.csv not found"
}

Start-Sleep -Seconds 2
If (Test-Path ".\CR067 Computer Utilization (CR026) $TodayIs.csv") {
    Copy-Item ".\CR067 Computer Utilization (CR026) $TodayIs.csv" -Destination "\\IL7354-964\C$\Users\2brunnga\Documents\Reports\New\CR067 Computer Utilization (CR026) $TodayIs.csv"
} Else {

    Write-Output "CR067 Computer Utilization (CR026) $TodayIs.csv not found"
}

Start-Sleep -Seconds 2
If (Test-Path ".\CR068 Disk Space Usage $TodayIs.csv") {
    Copy-Item ".\CR068 Disk Space Usage $TodayIs.csv" -Destination "\\IL7354-964\C$\Users\2brunnga\Documents\Reports\New\CR068 Disk Space Usage $TodayIs.csv"
} Else {

    Write-Output "CR068 Disk Space Usage $TodayIs.csv not found"
}

#__________________________________________________________________________________________________

# Add file ending to log file
Write-Output "$ScriptName complete"

Stop-Transcript 

