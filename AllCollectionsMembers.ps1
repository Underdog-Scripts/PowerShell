# Script to  create a text file for each collection with the members in tht file.

<#
$DeviceCollection ="2018-01 Non-Complaint"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="2018-01 Non-Compliant 03-21-2018"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="2018-04 Non-Compliant @ 8~36pm"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="2018-10 Non-Complaint"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="2018-11 - Unknown"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="2018-11 12 - Success"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="2018-11 Non-Complaint"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AD Password Reset Pre-2016"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALC SMS Budget Worksheet 1.0 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALC SMS EMC Corporation (Thawte Code Signing CA) 3.0 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALC SMS HyperSnap DX 5.0 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALC SMS Mapped Drives (Corporate Affairs) Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALC SMS Mapped Drives (PMCC Finance) INU Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALC SMS Oracle & Sybase & Brio Config Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALC SMS Oracle ADI 7.1 and Oracle 8.0.5 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALC SMS Oracle for Windows 1.0 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALC SMS QuarkXPress 4.1 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALC SMS Transition App Removal Tool Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALC SMS WSLink 2.16 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALCS CP Security Updates Lab Computers Pilot"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALCS FSF Special Pools Setter Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALCS Log File Collector Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALCS SMS Adobe Captivate 5.0 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALCS Vantive Fix"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALL Active Clients with new CA Loadset"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Active Directory Security Groups (RCH)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="ALL CA Windows Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Desktop and Server Clients"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Devices with no Value in PwdLastSet"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Field Sales Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Field Sales Force Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Field Sales Force Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Field Sales Office M50 WXP Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Field Sales Office T41 WXP Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Field Sales Office X61 WXP Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Lenovo M52 WXP Desktop Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Lenovo T60 and T60p WXP Laptop Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Lenovo X1 Tablet Gen 1 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Non-RIC Campus, CAB, RHQ WXP Computers RCH"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All PMUSA NON SP3 SMS Clients"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All PMUSA Salaried Employees"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Richmond HQ WXP Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All SCCM Site System"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All SMS Clients with Duplicate SMS GUIDs"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Systems"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Systems Limited Collection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Systems~1~"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Unknown Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All User Groups"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Users (RCH)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Windows 2000 Professional Systems (RCH)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Windows 2000 Server Systems (RCH)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Windows 7 Devices (with and without SCCM Client)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Windows 7 Workstations"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Windows 8 Workstations"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Windows 8.x Devices (with or without SCCM Client)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Windows Mobile Devices (RCH)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Windows Mobile Pocket PC 2003 Devices (RCH)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Windows NT Workstation 4.0 Systems (RCH)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Windows Server Systems (RCH)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Workstations - Patch Limited"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="All Workstations (With or Without Client Install)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="App Fac-SW Field Sales Self Service Application Catalog Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac - SW Turning Technologies TurningPoint 2008 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac- SW Trend OfficeScan IDF ReRegister Utility Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac- SW Trend OfficeScan Intrusion Defense Firewall Client Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-ALC SMS Altria Fonts 1.0 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW 3M Detection Management Software 1.6.67 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ABBYY FineReader 11 Corporate Edition ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append



$DeviceCollection ="AppFac-SW ABBYY FineReader 12 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ACL Analytics 10.5 Desktop Edition ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ACL Analytics 10.5 Desktop Edition ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Actual Search and Replace 2.8.9 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Adobe Acrobat DC 2015 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Adobe Acrobat Professional 11.0.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Adobe Creative Suite 6 Design & Web Premium ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Adobe Illustrator CS5 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Adobe InDesign CS6 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Adobe Photoshop CS6 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Adobe Reader XI 11.0.06 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Adobe Shockwave Player 12.2.4.194 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW AGDC FieldEDGE File Update 1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ALCS Crafts OTS ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ALCS FieldEdge Reporting Utility Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFAc-SW ALCS GAPAssessment 2.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFAc-SW ALCS GAPAssessment 2.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ALCS ICIS 4.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ALCS Kiosk IE Shortcut 082013 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ALCS Kiosk IE Shortcut 082013 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ALCS MMS ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ALCS MMS Richmond 6.17 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ALCS StagFont 1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ALCS StagFont 1.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ALCS Trouble Call LineAssignment 3.0.0.73 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Altria Docs Desktop (EC 10.5) ENT Computers~1~"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Austin Digital ADI Wasabi 4.6 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Autodesk AutoCAD 2015 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Autodesk Design Review 2013 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW AutoDesk DesignReview 2012 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Autodesk Product Design Suite 2015 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Autodesk Suite 2012 REMOVAL ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Bentley Systems MicroStation 8.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW BNA Corporate Tax Analyzer 2012.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW CA ERwin Data Modeler 9.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW CambridgeSoft Chemdraw AX 14.0 ELN ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW CambridgeSoft ChemDraw Pro v15 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW CCH Small Firms Intelliforms 2009 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Cigarette Design Model 4.3.2 ENT Users~1~"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Cigarette Design Model 4.4 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Circos 0.69 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Cisco AnyConnect Secure Mobility Client 3.1.04072 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Cisco AnyConnect Secure Mobility Client 3.1.04072 ENT Removal Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Cisco AnyConnect Secure Mobility Client 3.1.06073 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Cisco AnyConnect Secure Mobility Client 4.2.01022 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Cisco Proximity for Desktop 1.0.0.118 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Cisco Supervisor Desktop 6.0.1.15 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Cisco Supervisor Desktop 9 Suite ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Cisco WebEx Meeting Center 30.3.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Cisco WebEx Meeting Center 8.23.2500 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Cisco WebexProductivityTools 2.4 ENT Users~1~"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Citrix Desktop Viewer 12.1.19 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Citrix Desktop Viewer 12.1.19 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Citrix Online Plugin 12.3 ENT Computers~1~"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Citrix Online Plugin 12.3 ENT Computers~1~~1~"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Citrix Receiver 14.3.0.5014 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW CLCBio CLCSequenceViewer 6.9 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Compare Spreadsheets for Excel 1.1.7 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Compusense 5.0.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Corel VideoStudio Pro X7 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Critical Tools WBS Chart Pro 4.9 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW CriticalTools WBS Schedule Pro 5.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Crystal Decisions Crystal Reports 9 Professional 9.2.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Dedicated Micros DM NetVu Observer 1.9.4 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Dedicated Micros DM NetVu Observer 1.9.4 ENT Computers_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Dedicated Micros DM NetVu Observer 1.9.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Dedicated Micros DM NetVu Observer 1.9.4 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Dell RemoteScan Universal 10 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Dell RemoteScan Universal 10 ENT Computers_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Dell RemoteScan Universal 10 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Design Safety Engineering Designsafe 5 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW DNASTAR Lasergene 12.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Documentum DM Drag and Drop ActiveX 1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW eDoc CIC Signit 9.0 AGDC Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW eDocument Setup 4.9.4.0 AGDC Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Embarcadero Technologies WebSphere 5.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Enterprise Connect Redirect Utility 1.0.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Environmental Systems Research Institute ArcGIS Desktop 10.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ePad Driver Suite 1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ePad Driver Suite 2.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ePad Driver Suite 2.0 ENT Computers_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ExpertGPS 4.89 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW FABES MigratestEXP ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW FileMaker FileMaker Pro 13 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW FileMaker FileMaker Pro 13 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW FileOpen Systems FileOpen Plug-In 1.0.872 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Google EarthPro 5.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Gotham Narrow Fonts Bundle 1.0.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Hansen-Solubility HSPiP 4.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Helios TextPad 7.1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW HP AssetManager 9.31 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW HPC KeyTrail 3.6.13 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Hummingbird Connectivity Secure Shell 2007 12.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Hummingbird Hummingbird Exceed 7.1.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Hummingbird Hummingbird Exceed 7.1.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW IBM iSeries Access for Windows 5.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW IBM iSeries Access for Windows 5.4 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW IBM SPSS Data Collection (CATTS) 6.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW IBM SPSS Decision Trees v23 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW IDM Computer Solutions UEStudio 12.10.1003 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW IDM Computer Solutions UltraCompare 8.50.1026 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW IDM Computer Solutions UltraEdit 20.00.1056 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Inmagic Inmagic/TextWorks Setup Workstation 13 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\"AppFac-SW Inmagic Inmagic'/TextWorks Setup Workstation 13 ENT Users.txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Insperity OrgPlus 9 Plug-In 9.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Insperity OrgPlus 9 Plug-In 9.1 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Insperity OrgPlus 9.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Inspyder Web2Disk 4.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW InstallShield Software Install Script MSI Engine 3.00.185 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Intel JPEG Library 1.5.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Ipswitch WS_FTP Pro 12.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW IRI Advantage Office 1.0.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW JD Edwards World Vision 2.0.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW KnowWare International The Qi Macros for Excel 2010 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Kutools For Word 7.10 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Labtronics Collect 6.1.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW LabWare AA PROD LIMS 3.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW LabWare BI DEV LIMS 1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW LabWare BI DEV LIMS 2.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW LabWare BI DEV LIMS 2.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW LabWare BI PROD LIMS 1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW LabWare Crystal Reports Runtime 9.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW LabWare Crystal Reports Runtime 9.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW LabWare FLVR PROD LIMS 2.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW LabWare FLVR PROD LIMS 2.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Labware QAPR DEV LIMS 1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Labware QAPR PROD LIMS 2.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Lexis Nexis Bridger Insight 1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Lhasa KB Database Update 2.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Lhasa KB Database Update 2.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW MAPILab Mail Merge Toolkit 2.6.1 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW MapMechanics Optisite 2.1.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW MDT Software AutoSave Windows Client 5.07 Custom ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Mettler-Toledo AG LabX Direct pH ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Microsoft C++ 2010 x64 Redistributable 10.0.40219 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Microsoft Dynamics CRM 2013 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Microsoft Dynamics CRM 2015 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Microsoft Excel 2013 Standalone Side By Side ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Microsoft Lync 2010 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Microsoft Office 2003 Web Components ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Microsoft Office Viewers 2010 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Microsoft OneDrive For Business FSF Migration Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Microsoft OneDrive For Business Migration (Available) Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Microsoft OneDrive For Business Migration Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Microsoft OneDrive for Business QA Install Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Microsoft Project Standard 2010 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Microsoft Visio Professional 2010 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Microsoft Visio Standard 2010 UNINSTALL Users DO NOT USE"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Mincom Linkone Publisher 5.0.1140.10 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Mincom LinkOne WinView 5.11 Individual ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW MindJet MindManager 14.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW MindJet MindManager 16 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Mindjet MindManager 9.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Minitab V17 Concurrent User License ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW MyID 1.0.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW NaturalReader 12 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW NeuexpowerNXPowerLite6 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Non_Conforming Products Update ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW NuMark Fonts 1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW OmniRIM Solutions Barcode Font 2.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW OmniRIM Solutions Barcode Font 2.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Open Text SOCKS Client 14.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW OpenText Enterprise Connect 10.3 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW OpenText SOCKS 14.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Oracle Administrator Client 11.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Oracle Client Administrator 10.2.0.5 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Oracle Essbase Spreadsheet Add-in Fusion Edition 11.1.1.4 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Oracle Financial Reporting Studio 11.1.2.3.500 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Oracle Financial Reporting Studio 11.1.2.3.700 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Oracle Hyperion EAS Console 11.1.2.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Oracle Hyperion Smart View Fusion Edition 11.1.2.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Oracle Hyperion Smart View Fusion Edition 11.1.2.5 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Oracle Hyperion Spreadsheet Addin 11.1.1.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Oracle Hyperion Suite 11.1.2.3 Admins ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Oracle Instant Client 11.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Oracle Oracle Client 8.1.7.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Oracle Oracle Client 9.2.0.8 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW OrgPublisher OrgChart Plugin ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW PMDF Budget Management 8.0.2 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW PMDF Budget Management 8.0.2 Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW PMDF CSD 2.6.0 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW PMDF CSD 2.7.0 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW PMDF Sales Management 9.0.0 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW PMUSA HMI Support Setup 1.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW PMUSA Holdwork Tools 1.5.5 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW PMUSA PPO 3.1.19 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW PMUSA WeighBelt Explorer 1.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Prezi PreziDesktop 4.7.5 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Proginy TimelineMaker Pro ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Prometheus Group Print Diversion 4.3.7 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Prometheus to PDF 2.0.0.1 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW QDAMiner 4.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW QlikTech QlikViewDesktop 10 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Quest PuTTY 0.3.0.129 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW QuickTime v7.74 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Real VNC 4.6.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW RealLegal eViewer 6.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Records Management Confidentiality Tool 1.3 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW RedPrairie JDA Client 2013.2.2 EGL Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW RedPrairie JDA Client 2013.2.2 EGL Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW RedPrairie JDA Client 2013.2.2 NSH Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Refresh Workstation ConfigurationENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Sales Launch Pad 4.5.2 AGDC Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW SAP BEX 7.2 SP10 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW SAP BEX 7.3 Patch 10 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW SAP BusinessObjects BI platform 4.1 Client Tools SP3 Patch 2 ENT Computers_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW SAP BusinessObjects BI platform 4.1 Client Tools SP3 Patch 2 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW SAP BusinessObjects Dashboards 4.1 SP3 Patch 2 Update 14.1.3.1334 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW SAP Crystal Reports 2011 14.0.7.1147 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW SAP EHS Expert 3.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW SAP EHS Expert 3.2 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW SAP GUI for Windows 7.20 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW SAP GUI for Windows 7.30 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW ScanSoft OmniPage Pro Office 14 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Seagull Software BlueZone 6.1.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW SigmaPlot 12.5 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Skype 7.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Skype 7.7 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Sopheon Accolade 10.1.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Sopheon Accolade 8.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW SQL Server Management Studio 2012 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW SSRS 2012 ActiveX Plugin ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Sungard Treasury Systems MetaView 6.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Synactive GuiXT 4.2.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW TCDI Cvedit 2.5.09 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Teammate R10.2.4-ConfigCopier ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW TeamMate R10.4.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW TechSmith Camtasia Studio 7 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW TechSmith Camtasia Studio 8 ENT Uninstall Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW TechSmith Camtasia Studio 8 ENT Uninstall Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Thawte Consulting EMC Corporation (Thawte Code Signing CA) 3.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Thompson Reuters Reference Manager 12 Workstation 12.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Thomson Reuters West Case Timeline 2.5 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Turning Technologies TurningPoint 4.3.2.1178 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Uninstall IDM Computer Solutions UltraCompare 8.50.1026 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Uninstall MDT Software AutoSave Windows Client 5.07 Custom ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Uninstall MindJet MindManager 15 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Uninstall SAP BusinessObjects BI platform 4.1 Client Tools SP3 Patch 2 ENT Computer"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Uninstall Tableau Desktop 8.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Uninstall Workshare Professional 7.0.10000 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Update for Microsoft Windows (KB2775511) ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW USST Vision Supervisory 1.2.0 NSH ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW UTS Data Transfer ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Visual Lighting Software Visual Lighting Software 2.06 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW VmWare Player 3.1.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Warren Selbert LINK 2.16 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW WinShuttle Runner 10.6 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW WireShark Network Analyzer 2.0.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Workshare License Utility 3.0.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-SW Workshare Workshare Professional 7.0.10000 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-Trend IPXFER Filter Collection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AppFac-Trend Micro IPXFER Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="AutoScriptTest"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CA Windows 10 x86 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CA Windows 7 SP1 x64 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CA Windows 7 SP1 x86 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CA Windows 7 SP1 x86 Reduced Core (No MBAM) Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CJ - Altria Docs ~ Enterprise Connect - Size limit x64"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CJ - Virtual Machines AD Group"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CJTST - MS Project Pro"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CJTST - Womdpws 10 Project testing"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Client Health ~ Client Verison Not Latest (1710)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Client Health ~ Clients Inactive"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Client Health ~ Disabled"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CP Aviation FOS Kiosk Workstation Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CP CITRIX XenDesktop Non-Persistent Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CP COC Kiosk Workstation Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CP Enterprise Site Discovery Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CP Executive Support Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CP Igloo Kiosk Workstation Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CP NuMark Miami Returns Shared Workstations"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CP PMU Direct Materials HMI Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CP PMUSA ACL Shared Workstation Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CP PMUSA Fixer Kiosk Workstation Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CP PMUSA Park 500 Shared Workstation Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CP PMUSA Precon NCP Kiosk Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CP PMUSA TTC Shared Workstation Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CP USST Manufacturing Shared Workstation Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CP Virtual Workstation Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="CP XD IE11 Testing Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="DENY User Application Deployment Client Settings"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Deploy Windows 10 CBB x86"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Deploy Windows 10 LTSB x86"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Deploy Windows 7 SP1 CRT Lab x86"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Deploy Windows 7 SP1 ENT x64"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Deploy Windows 7 SP1 ENT x86 QA"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Deploy Windows 7 SP1 with IE 8 ENT x64"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Deployed Users not in the new Lync Collection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="DO NOT USE - EXEQA - AppFac- SW Trend IPXfer v1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="DO NOT USE - EXEQA - AppFac-SW - TrendIPXfer ENT Computers (2nd)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="DO-NOT-USE-TST-All - Workstations - Patching - PROD"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="DVC-CB-Win10Refresh-CustomData"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EDDIE TEST2 - DO NOT USE"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EUO - All Workstations (Limited to EUO)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXCLUDE - AppFac-SW Adobe Reader ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXCLUDE - AppFac-SW Cisco AnyConnect 4.2 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXCLUDE - AppFac-SW JDA ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXCLUDE - AppFac-SW LHS (Legal Hold System) 5.3 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXCLUDE - AppFac-SW Proximity 2.0.2 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXCLUDE Altria Docs Desktop 10.5 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXCLUDE Altria Docs Desktop 10.5 ENT Computers~2~"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXCLUDE Cisco AnyConnect Uninstaller"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXCLUDE Microsoft Internet Explorer 11.0 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXCLUDE SW eDocument Setup Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXCLUDE SW Executive Users Master Collection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXCLUDE SW LAP Reset Utility ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXCLUDE SW Sales Launch Pad Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXCLUSION - AppFac-Trend Micro IPXFER Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXCLUSION AppFac-SW OpenText WebDAV Client 10.5 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE -AppFac-SW Easy GIF Animator 7.0 ENT User-Retired"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW - Uninstall Microsoft Visio (All Versions) 1.0.0_0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW ActivePerl 5.24.1 Build 2402 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Addin Technology KuTools for Excel 18.00 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Adobe Acrobat DC Professional 2017-AVAILABLE-ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Adobe Air 30.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Adobe Captivate 2017 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Adobe Captivate 2017 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Adobe Creative Cloud Suite v2018 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Adobe Creative Cloud Suite v2018 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Adobe Digital Editions 4.5 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW ALCS Trend VP ReRegister Computers Phase 3"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW ALCS UninstallKiosk IE Shortcut 082013 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Analytic Insight Market Solutions 2.0 Patch ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW ArsClip5.0.8 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Articulate Storyline 3 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW AutoCAD 2017 ENT Computers_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW AutoCAD 2017 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW AutoDesk 2017 License Server Environment Variable-AutoCAD ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW AutoDesk 2017 License Server Environment Variable-AutoCAD ENT User_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW AutoDesk 2017 License Server Environment Variable-Vault ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW AutoDesk 2017 License Server Environment Variable-Vault ENT Computers_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW AutoDesk 2017 License Server Environment Variable-Vault ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW AutoDesk Product Design Suite 2017 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW AutoDesk Product Design Suite 2017 ENT Computers_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW AutoDesk Vault Professional 2017 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW AutoDesk Vault Professional 2017 ENT Computers_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Berkeley Madonna 9.1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Bios Update 1.22 For Lenovo T450 EN Device"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Blumentals Easy GIF Animator v7.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Blumentals Easy GIF Animator v7.3 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Business Objects 4.2 Update (New Users) ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Business Objects 4.2 Update (New Users) ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Business Objects 4.2 Update ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Business Objects 4.2 Update ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Change-Pro Premier Suite 10.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Citrix Receiver 4.6 ENT Computers (COC)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Creative Cloud Suite Creative Cloud Edition ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Crystal Reports 2016 Update ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Crystal Reports 2016 Update ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW DirectLink 6.0 ACL Plug-in Windows 10 Uninstall ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Easy GIF Animator Pro v6"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Environmental Systems Research Institute ArcMap Basic v10.5.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Environmental Systems Research Institute ArcMap Basic v10.5.1 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW ePad IntegriSign Desktop 12.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Epson WF7620 Printer_Scanner Drivers ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Google Chrome x64 70.0.3538.77 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW GraphPad Prism 7 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Horizon Client for Windows 4.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-Sw IBM SPSS Smartreader X64 ENT User"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-Sw IBM SPSS Smartreader X64 ENT User_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW IDM UltraCompare x64 18.00.0.86 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW IDM UltraCompare x64 18.00.0.86 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW IDM UltraEdit x64 25.10.0.48 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - Appfac-SW iMacros Version 12.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - Appfac-SW iMacros Version 12.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW JMP PRO 14 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW June 2016 update for Lync 2010 64 bit ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Lenovo w540 Wireless Drivers 18 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW MATLAB 2018b v9.5 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW MATLAB 2018b v9.5 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Metasys Launcher 1.6 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Microsoft .NET Framework 4.6.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Microsoft Dynamics AX 2012 R3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Microsoft Photos v. 17.214.10010.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Microsoft Power BI Desktop 2.63.3272.40461 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Microsoft Project Standard 2016 16.0.9126.2210 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - Appfac-Sw Microsoft Project Standard 2016 16.0.9126.2336 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Microsoft Reporting Services Client Print 10.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Microsoft SQL Server Management Studio 2017 v14.0.17285.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Microsoft SQL Server Management Studio 2017 v14.0.17285.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Microsoft Visio Professional 2016 16.0.9126.2210 AVAILABLE ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Microsoft Visio Professional 2016 16.0.9126.2210 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Microsoft Visio Professional 2016 16.0.9126.2336 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Microsoft Visio Viewer 2016 x64 v16.0.4339.1001 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Microsoft Visio Viewer 2016 x86 16.0.4339.1001 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Microsoft Visio Viewer 2016 x86 16.0.4339.1001 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE – AppFac–SW Microsoft Visual C++ 2015 Redistributable (x64) - 14.0.24215 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Microsoft Visual Studio Tools for Office Runtime 2010 10.0.60828 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Mindjet MindManager 2017 17.0.290 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW MMS 10 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW MorningStar Direct 3.19.00 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW MorningStar Direct 3.19.00 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Mozilla Firefox 58.0.2 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Mozilla Firefox 58.0.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - Appfac-SW Notepad++ x64 7.5.6 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW OpenText Rightfax 16.4 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW OpenText RightFax Suite 10.6 Update AVAILABLE ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW OpenText RightFax Suite 10.6 Update ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW OpenText RightFax Suite 10.6 Windows 10ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Oracle Client 12c (32-bit) 12.1.0.2.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Oracle Client 12c (32-bit) 12.1.0.2.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Oracle Hyperion Suite 11.1.2.4 Admins ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Oracle Hyperion Suite 11.1.2.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Oracle Smart View v11.1.2.5.720 AVAILABLE ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Oracle_Admin_11.2_X86 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW PC Miler v30 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Phish Alarm 3.1.29 ENT Users (ALG UR & ELT Deployment)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Phish Alarm 3.1.29 ENT Users (ENT UR Hourly Employees)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Phish Alarm 3.1.29 ENT Users (PMU UR Loc PMUSA P500 Chester VA)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Phish Alarm 3.1.29 ENT Users (UST UR USSTP Hourly Employees)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW PMDF CSD Query Manager 2.7.0 Reports ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Progeny Timeline Maker Pro v4.5.32.16 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Prometheus_PrintDiversion_5.0.9 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW reaConverter 7 Pro Edition 7.3.99 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW RMG Networks StudioDesign 12.02 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW RMG Networks StudioDesign 12.02 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SAP GUI 7.2~Rollback~ ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SAP GUI 7.2 ENT Computers_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SAP GUI 7.4 Compilation 3 Patch Level 10 ENT Computers (Kiosks) Phase C"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SAP GUI 7.4 Compilation 3 Patch Level 10 ENT Computers (Kiosks) Phase D"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SAP GUI 7.4 Compilation 3 Patch Level 10 ENT Computers (Kiosks) Phase E"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SAP GUI 7.4 Compilation 3 Patch Level 10 ENT Users Phase3"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SAP GUI 7.4 Compilation 3 Patch Level 10 ENT Users Phase6"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SAP GUI 7.4 Compilation 3 Patch Level 10 ENT Users Phase7"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SAP GUI 7.4 Compilation 3 Patch Level 10 Required ENT (Shared) Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SAP GUI 7.4 Compilation 3 Patch Level 10 Required ENT Computers (Kiosks) Phase A"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SAP GUI 7.4 Compilation 3 Patch Level 10 Required ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SAP GUI v7.50 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SAS ACCESS Interface to ODBC v9.4 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SAS ETS v9.4 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SAS OR v9.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SecuriTree 4.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SnagIT 13 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SolidWorks SP2 v2018 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SolidWorks SP2 v2018 ENT Users_AVAILABLE_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Sopheon Accolade Microsoft Add-ins 11.2.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW SystemTools Hyena v13.00.2000 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Tableau Desktop 2018.1.6 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Tableau Desktop 2018.1.6 ENT Users - OLD"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Tableau Reader 10.2.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Tableau Reader 2018.2.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Teradata Tools and Utiliteies v.15.10"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Teradata Tools and Utiliteies v.15.10 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Toad for Oracle 13.1.0.78 X86 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Toad for Oracle 13.1.0.78 X86 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Toad License Capture 1.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Trend Vulnerability Protection 9.6.7599 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW UninstallBlackboard CollaborateLauncher 1.6.5.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW UninstallSAS ETS v9.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Acronis Backup Agent 11.5 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Adobe Acrobat DC Professional 2017 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Adobe Captivate 2017 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Adobe Captivate 2017 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Adobe RoboHelp 2019 14.0.1.39 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Altria Docs 16.2.4 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall ArcMap Basic v10.5.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall ArcMap Professional v10.5.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Articulate Storyline 3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall AutoCAD 2017 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall AutoDesk Product Design Suite 2017 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall AutoDesk Vault Professional 2017 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall BACIS FRM 2001 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - Appfac-SW Uninstall Balsamiq Mockups v3.5.15 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Berkeley Madonna 9.1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Blumentals Easy GIF Animator v7.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Citrix Receiver 4.11 v14.11.0.17061 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Crystal Reports 2016 Update ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Dell RemoteScan Universal 10 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Dell RemoteScan Universal 10 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Forefront Identity Manager 2010 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall IDM UltraCompare x64 18.00.0.86 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Intuit Quickbooks Enterprise Solutions 13 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Java 8 Update 172 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall JMP PRO 14 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Lexis Nexis Bridger Insight 1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Microsoft Power BI Desktop 2.63.3272.40461 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Microsoft Project Standard 2010 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Microsoft Project Standard 2016 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Microsoft SQL Server Management Studio 2017 v14.0.17285.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE – AppFac–SW Uninstall Microsoft Visual C++ 2015 Redistributable (x86) - 14.0.24215 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall OpenText RightFax Suite 10.6 Update ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Oracle Hyperion Smart View Fusion Edition 11.1.2.5 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Oracle Smart View v11.1.2.5.720 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Oracle_Admin_11.2_X86 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall PC Miler v30 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Primary Portal Web Controls 1.0 PMUSA Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Progeny Timeline Maker Pro v4.5.32.16 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Prometheus_PrintDiversion_5.0.9 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall RMG Networks DesignStudio 12 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall RStudio_1.1.456 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall SAP GUI 7.5 PL4 R1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall SAP_GUI_for_Windows_C3_7.20 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall SAS ACCESS Interface to ODBC v9.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall SAS Analytics Pro v9.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall SAS Analytics Pro v9.4 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall SAS QC v9.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall SnagIT 13 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall SSRS 2012 ActiveX Plugin ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall SystemTools Hyena v13.00.2000 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Tableau Desktop v.10.2.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall TechSmith Snagit 2018 v.18.2.1.1590 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Timeline Maker Pro 4.2.39.14 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Timeline Maker Pro 4.2.39.14 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall Toad for Oracle 13.1.0.78 X86 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall USEPA BMDS v2.7.0.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall UTS Data Transfer ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall VMWare Horizon Client 4.8.0.1562 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Uninstall VNC Viewer 6.18.625 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW USEPA BMDS v2.7.0.4 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW UTS Data Transfer AVAILABLE ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW VLC Media Player 2.2.4 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW VLC Media Player 2.2.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW VLC Media Player 3.0.3 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW VNC Viewer 6.18.625 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW VNC Viewer 6.18.625 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW Windows Movie Maker 2012 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW WinShuttle Runner 10.7.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW WinShuttle Runner 10.7.3 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW WinShuttle Studio 10.7.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW WinShuttle Studio 10.7.3 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW WorkShare License Utility 5.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW-DDM Inventory Agent Uninstall Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW-SAS STATPLUS v9.4 ENT User"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFac-SW-SAS STATPLUS v9.4 ENT User_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFaq-SW Balsamiq Mockups 3.5.15 R1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFaq-SW Balsamiq Mockups 3.5.15 R1 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFaq-SW SmartBear TestComplete 12 12.50.4142.7 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFaq-SW Uninstall Workshare Professional 9.5.787.0 ENT R1 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFaq-SW WorkShare License Utility 4.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFaq-SW Workshare Professional 9.5.787.0 R1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFax-SW Cisco AnyConnect Secure Mobility Client 4.2.01022 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFax-SW Uninstall CorpTax 2016 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppFax-SW Uninstall OpenText Rightfax 16.4 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - AppHac-SW WinZip Pro v.22.0 Required ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - Cisco AnyConnect Secure Mobility Client 4.2.01022 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - Microsoft Configuration Manager Console Install - v1806"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - Microsoft Online Services Sign-In Assistant 7.250.4303.0 (Enterprise)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - Office 365 Client Installer - Monthly Channel - v1802 Build 16.0.9029.2167 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE - Phase 4 Phish Alarm 3.1.29 - 6~25"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXE- AppFac-SW Uninstall Rcore_RforWindows_3.5.1_0"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA -AppFac-SW Adobe Acrobat Reader v18.011.20035 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW Altria Docs Desktop 10.5.2.809 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW Articulate Storyline 2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW Map Print Queue RRMC_STKRM_BW3 v 1.0"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW Microsoft Dynamics CRM 2016 for MS Office Outlook with offline access v.8.0.0.1088 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW Microsoft SQL Server 2012Express SP2 EXT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW Microsoft Windows Camera v.2017.619.10.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW PC Miler v30 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW PC Miler v30 ENT Computers_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW SAP GUI 7.4 Compilation 3 Patch Level 10ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW SAP GUI 7.4 Compilation 3 Patch Level 10 (Kiosk) ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW SAP GUI Upgrade Alert 7.40 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW Uninstall Altria Docs Desktop 10.5.2.474 ENT Computers_OLD"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW Uninstall Altria Docs Desktop 10.5.2.474_N ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW Uninstall Microsoft SQL Server Management Studio Uninstall All v1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW Uninstall MS Visio Professional 2016 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW Uninstall PPO v3.1.25 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW Uninstall RedPrairie JDA Client 2013.2.2R1 Nash ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - AppFac-SW Uninstall SAP GUI 7.4 Compilation 3 Patch Level 10 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEQA - Honeywell MAXPRO VMS R3.10.326 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT- AppFac-SW as Install Uninstall SAP Business Objects BI Platform 4x Client Tools Uninstall All v1.00 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT- AppFac-SW Uninstall SAP Business Objects BI Platform 4x Client Tools Uninstall All v1.00 ENT User"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac- SW IBM iSeries 5.4 R2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW - Uninstall Adobe Acrobat 1.0.0_0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Adobe_PremiereElements_2018_16.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Alteryx Designer 2018 3.4.51585ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Altria Docs Desktop 10.5 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW ArcMap Professional v10.5.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Caliper Maptitude 2016 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Caliper Maptitude 2016 Uninstall ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Cisco AnyConnect 4.3 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Cisco AnyConnect Secure Mobility Client 4.5.03040 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Citrix Receiver 4.12 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Citrix Receiver 4.6 ENT Computers-OLD"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW EMS 1.1.86 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW GraphPad Prism 7 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW IBM SPSS Statistics 25 PKG ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW IHSM_Eviews_10.00 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Infinote Finder v3.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Information Mapping FS Pro 6.06 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Information Mapping FS Pro 6.07 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Microsoft Dynamics AX 2012 R3 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Microsoft SQL Server Compact 3.5 SP2 (X64) 3.5.8080.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW OpenText RightFax Suite 10.6 Update ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Oracle Crystal Ball 32bit v11.1.4176.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Oracle Hyperion SmartView v11.1.2.5.520 ENT User"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Postman 6.4.4 x64 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW RedPrairie JDA Client 2013.2.2R1 Eagleone ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW RedPrairie JDA Client 2013.2.2R1 Hopskin ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - Appfac-SW Rockwell FactoryTalk Activation Manager 4.02 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW SAP BusinessObjects BI Platform 4.2 Client Tools SP5 P4 v.14.2.5.2813 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW SAP_GUI_for_Windows_C3_7.20 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW SAS 9.4 Serial Updater ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Smartbear License Manager 2.90 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Statsolutions_nQuery_nTerim4.0_1.1.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Tableau Add-In 7 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Tableau Desktop2018.1.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Tableau Reader 2018 .2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Tableau Reader 2018 1.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Trend Micro Vulnerability Protection 9.6.6400 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Trend Micro Vulnerability Protection 9.6.7516 RDE ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW TSI Detection Management Software v3.0.0 ENT User"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW uninstall Addinsoft XLSTAT 2018 v20.4.51329 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Uninstall ArcMap Professional v10.5.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Uninstall Certara_Phoenix_8.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Uninstall CorelDRAW Graphics Suite 2018 20.1.0.708 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Uninstall Easy GIF Animator Pro v6 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Uninstall Hygiena SureTRend Data Analysis v.4.1.194 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Uninstall IBM SPSS Statistics 25 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Uninstall Integra VIALAB 1.0.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Uninstall Multicase Skin and Eye Toxicity Models x64 v1.5.2.0 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Uninstall Oracle Crystal Ball 32bit v11.1.4176.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Uninstall RedPrairie JDA Client 2013.2.2 HPK Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Uninstall RedPrairie JDA SCA Client UST_RIC 2013.2.2 ENT User"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Uninstall SAP GUI 7.5 PL4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Uninstall Video Studio Pro X7 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFac-SW Uninstall VNC Viewer 6.18.625 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFaq-SW Corptax Client 2016 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFaq-SW Microsoft Visual Studio Tools for Office Runtime 2010 10.0.60828 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFaq-SW Uninstall Corptax Client 2016 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - AppFaq-SW Uninstall VideoStudio Pro X3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - Microsoft Windows Update KB3159398 ENT Computers - Uninstall"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - SAP GUI 7.4 Compilation 3 Patch Level 10 (Kiosk) ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT - Visual C++ redistributable 2015 (x64) v.14.0.23026 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="EXEUAT- AppFac-SW Oracle Client 12c (64-bit) ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Hardware Inventory ~ Active Clients Not Reporting HW Inventory Greater than 7 Days and Have Logged In within 30 Days"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Hotfix for Windows XP (KB969557) IE8 Bluescreen"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Imaging servers Patch KB4103723"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="KyleTestExclude"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Lenovo Primax Keyboard Firmware Update 2.0.47.060 and 2.00.19.21"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - 3rd Party Patch Collections"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - 3rd Party Unknown"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - NOV2017PatchIssue1"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - NOV2017PatchIssue3"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - NOV2017PatchIssue5"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - Tableau Desktop Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - Test AD Container Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - Test ChemDraw ActiveX 14 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - TEST CI TLS 1.1~1.2 Reg Change Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - Test Flash"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - Test Project Pro Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - Test Users AD Group Name"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - Test VisioPro Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - Test Windows Patches Computer - Scan Tool"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - Test Windows Patches Computer Unknown"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - Test_AD_User_Add_UsingGroupName"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - Test_AD_User_Add_UsingSecurityGroup"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO - Test_AD_User_AGDC_UR_FSF_All"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO- POGTool v16.2.2 Users Query"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="LO- Test WinZip Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Machines with out a client (REM)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Machines with out old client (REM)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MDT Applicaiton Layering Only"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MDT Application Layering"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Migrated collection PMU SMS Access Connections Profile Update and sub-collections"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Migrated collection TESTING ONLY-New OpCo Collection and sub-collections"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - Appfac - SW Uninstall Adobe Photoshop CC 2018 v19.1.5 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac SW Uninstall Prometheus to PDF 2.0.0.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SWUninstall Mindjet MindManager 2018 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW ACL for Windows v12.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW ACL for Windows v12.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW ACL for Windows v12.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Adobe Illustrator CC 2018 v22.1 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Adobe LiveCycle v11 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Adobe LiveCycle v11 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - Appfac-SW Adobe Photoshop CC 2018 v19.1.5 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Adobe Robohelp 2015 x64 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Adobe Shockwave Player 12.2.4.194 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Adode Robohelp 2017 v13.0.2.334 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW AIS CSD Management 3.0.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFAc-SW ALCS GAPAssessment 2.0 AVAILABLE ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW ALCS Trouble Call System RIC 5.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW ALCS Trouble Call Workstation 5.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW ALCS Uninstall Moisture Meter 2.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Altria International Sales 1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Analytic Insight Market Solutions (AIMS) 3.0.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Analytic Insight Market Solutions 1.5.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Beyond Compare 4.1.9 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Beyond Compare 4.1.9 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Bluebeam Revu ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Brother P-Touch Editor 5.1 ENT User_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Camtasia Studio 18.0.4.3822 AVAILABLE ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Cisco IP Communicator 8.6 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Cisco Unified WFO Monitoring and Recording Prod 115.1.865 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Cisco Unified WFO Monitoring and Recording Prod 115.1.865 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Critical Tools WBS Schedule Pro v.5.1.0024 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW DBText 16 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW DBText 16 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW EndNote X7 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW EndNote X8 v18.0.1.10444 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Equipment Management System v9.8 ENT User"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Filemaker Pro 17.0.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW FileOpen Client (x64) B9793.0.149.979 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - Appfac-SW GagePack v12.0.18086.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW HireVue Live 2.2.3 AVAILABLE ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW iQOS v1.0.17.10 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW IRI Unify Office 3.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Java 1.8.0.172 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Java 8 Update 1.8_111 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Java 8 Update 1.8_111 ENT HQA computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Java 8 Update 1.8_111 MMS 10 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Kodak Smart Touchv1.8.55.331 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Kodak Smart Touchv1.8.55.331 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW LabWare AA DEV LIMS 2.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW LabWare FLV QA LIMS 4.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW LabWare FLV QA LIMS 4.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW LabWare QAPR DEV LIMS 2.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW MatrixStl FlowCode v8.0.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Microsoft Power BI Desktop 2.39.4526.362 Required ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Microsoft Report Services Client Print 10.0.2841.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW MPlus 8.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW MSIQA - IC3M_Hybrid_App"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Office Timeline 3.16 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW OpenText Socks Client 14.0.14 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW PMUSA UninstalleContract Signature Pad 1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW PMUSA Uninstall WeighBelt Explorer 1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW POGTool 16.4.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW POGTool 17.4.0 Uninstall ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW POGTool v17.4.0 ENT Users_2"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Proximity 2.0.2 ENT Computers HQ&HQA"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Proximity 2.0.2 ENT Computers IT OPS"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Proximity 2.0.2 ENT Computers Middleton & USSTC"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW RSA SecurID Software Token 5.0.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW SafeNet MobilePASS 8.4.4.99 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Sales Application Launch Pad 4.6.2 ENT Users~1~"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Sales Application Launch Pad 4.6.5 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Sales Launch Pad 4.6.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Sales Launch Pad 5.0.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Sales Launch Pad 5.0.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW SAP Crystal Reports Runtime Engine 64bit 13.0.16 (Version 2.7.0) ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW SecureCRT 8.3.2"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW SecureCRT 8.3.2 AVAILABLE ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW SecureCRT 8.3.2_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Seiko Instruments Smart Label Creator 1.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Seiko Instruments Smart Label Creator 1.2 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Symantec Enterprise Vault Compliance Accelerator Client 11.0.1 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Symantec Enterprise Vault Discovery Accelerator Client 11.0.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Symantec Enterprise Vault Discovery Accelerator Client 11.0.1 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW T440s Wireless Drivers ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW TeamMate Desktop R11.2- Available-ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Teklynx Label Matrix 2015 V.15.00.00 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Teklynx Label Matrix 2015 V.15.00.00 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Training Attendance Management v1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Training Attendance Management v1.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Unify Office 3.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW UninstallBrother P-Touch Editor 5.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall .NET Provider for Teradata 14 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Acronis Backup Management Console 11.5 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Adobe Acrobat DC Professional 2017 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Adobe Robohelp 2015 x64 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Adode Robohelp 2017 v13.0.2.334 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall ALCS Trouble Call Workstation 5.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Altria AIC Configuration Utility v1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Beyond Compare 4.1.9 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Bluebeam Revu ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Camtasia Studio 18.0.4.3822 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Camtasia Studio 18.0.4.3822 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Dedicated Micros DM NetVu Observer 1.9.4 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Dedicated Micros DM NetVu Observer 1.9.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall EndNote X8 v18.0.1.10444 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall HireVue Live 2.2.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Inmagic Inmagic~TextWorks Setup Workstation 13 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall JD Edwards World Vision 2.0.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Labtronics Collect 6.1.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall LabWare AA DEV LIMS 2.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Labware FLV DEV LIMS 4.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall LabWare FLVR PROD LIMS 2.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall LabWare M0094 Quality Analyst Interface OCX 1.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall LabWare QAPR DEV LIMS 2.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall LabWare QAPR QA LIMS 2.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Leffingwell & Associates Flavor Base 10th Edition ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall MacroWorks 3.1 v1.1.1.88 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Microsoft Lync 2010 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Microsoft Power BI Desktop 2.39.4526.362 Required ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Mindjet MindManager2019 19.0.306 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall MPlus 8.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall OmniRIM Solutions Barcode Font 2.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Oracle Hyperion EPMA Clients 11.1.23500 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Prometheus Group Print Diversion 4.3.7 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Proximity 2.0.2 VDI ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Quest Toad for Oracle 12.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Sales Launch Pad 5.0.3 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append
#>
$DeviceCollection ="MSI - AppFac-SW Uninstall SAP EHS Expert 3.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall TeamMate Desktop R11.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Teklynx Label Matrix 2015 V.15.00.00 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Turning Technologies TurningPoint 2008 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Unify Office 3.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Vista POG Tool 18.0.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall WinShuttle_Studio_10.07.0003-(Direct-Query-Transaction) ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstall Zebra Printer Drivers 4.5.0.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Uninstaller Cisco IP Communicator v.8.6.6.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW Vista POGTool 18.1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW WebEx Productivity Tools 31.7.2.77 ENT Computers CRT, Park&MC"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW WebEx Productivity Tools 31.7.2.77 ENT Computers IT OPS"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW WebEx Productivity Tools 32.2.0.47 ENT Users FSF Region 2"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW WebEx Productivity Tools 32.7.0.47 ENT Computers CRT, Park&MC"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW WebEx Productivity Tools 32.7.0.47 ENT Computers HQ&HQA"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW WebEx Productivity Tools 32.7.0.47 ENT Computers Middleton & USSTC"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW WebEx Productivity Tools 32.7.047 ENT Users FSF Regions1&4"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - Appfac-SW WebEx Productivity Tools 33.0.1.66 (IS) ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFac-SW WebEx Productivity Tools ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFaq-SW Altria AIC Configuration Utility v1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFaq-SW Altria AIC Configuration Utility v1.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - Appfaq-SW TechSmith Snagit 2018 18.2.2.2240 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - Appfaq-SW TechSmith Snagit 2018 18.2.2.2240 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFaq-SW Uninstall TechSmith Snagit 2018 18.2.2.2240 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFax-SW Go Paperless 1.0.8.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - AppFax-SW Go Paperless 1.0.8.0 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - SCCM Console 1706"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - Uninstall AppFac-SW ALCS StagFont 1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - Uninstall AppFac-SW PST Viewer 6.0.357.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSI - Uninstall AppFax-SW Go Paperless 1.0.8.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIQA - AppFac-SW Acronis Backup Agent 11.5 QA ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIQA - AppFac-SW FieldEdge Shortcut ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIQA - AppFac-SW Microsoft Report Viewer 2012 Runtime ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIQA - AppFac-SW Microsoft SQL Server Compact 4.0 SP1 x64 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIQA - AppFac-SW Oracle Financial Reporting Studio 11.1.2.4.000 ENT Device"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIQA - AppFac-SW Oracle Hyperion Offline Planning v11.1.1.4 ENT Devices"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIQA - AppFac-SW PMUSA Uninstall Non-Conforming Product 3.5.13 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIQA - AppFac-SW Uninstall Adobe Acrobat DC PRO 2015 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIQA - AppFac-SW Uninstall WinShuttle_Query_10.07.0003 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIQA - AppFac-SW Uninstall WinShuttle_Transaction_10.07.0003 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIQA - AppFac-SW WinShuttle_Transaction_10.07.0003 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT -AppFac–SW Uninstall Compatibility Pack for the 2007 Office system ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-Phish Alarm v3.0.44 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Cisco Unified WFO Monitoring and Recording Recording v115.1.865 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Cisco WebEx Meeting Center 31.5.1.12 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Cisco WebEx Meeting Center 31.5.1.12 Required ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Core Apps QA ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW ePadLink ePad Driver 12.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Equipment Management System v9.8 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW ERI Salary Assessor 1.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Google Earth Pro 7.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW HPC KeyTrail 3.6.13 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW iQOS v1.0.17.10 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW IRI Unify Office 3.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Lhasa KB Database Update 2.0 R1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW MDT Software AutoSave Windows Client 5.07 Custom ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Oracle Financial Reporting Studio 11.1.2.4.000 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Phish Alarm v3.0.44 ENT User"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Trouble Call LineAssignment v3.0.0.73 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Uninstall ALCS Trouble Call Workstation 5.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Uninstall DecisionTools Suite 6.3.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Uninstall DecisionTools Suite ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Uninstall DecisionTools Suite ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Uninstall Equipment Management System v9.8 ENT User"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Uninstall Lhasa KB Database Update 2.0 R1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Uninstall West Case Timeline ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Vista POGTool 18.1.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW VMTools v10.1.5 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW Win TMC 3.1.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW WinShuttle Runner 10.7.3R1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFac-SW WinZip Pro 20 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - AppFaq-SW Altria AIC Configuration Utility v1.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT - MSIUAT -AppFac–SW Compatibility Pack for the 2007 Office system ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSIUAT -l AppFax-SW Go Paperless 1.0.8.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="MSUAT - AppFac-SW .NET Provider for Teradata 14 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="NET FRAMEWORK 4.5"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Office 365 Update"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="OSD Computer Test Collection-Win10 Sales Refresh"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PATCH - EXCLUDE WSUS Security Updates Production (RCH)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PATCH - Security Update Pilot Group - Application Testing"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PATCH - WSUS All Existing Workstations"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PATCH - WSUS Patch Office 2003 SP3"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PATCH - WSUS Security Updates Pilot"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PATCH - WSUS Security Updates Pilot WSS Managed Servers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PATCH - WSUS Security Updates Production (RCH) - RS TEST"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PATCH - WSUS Security Updates WSS Managed Servers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PATCH - WSUS Testing"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMDF Budget Application v8.0.2 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMDF Sales Management v9.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS Exchange Migration Utility Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS IBM Access Connections 4.41 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS Internet Explorer 8.0 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS Market Master 3.1.25 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS Microsoft MapPoint 2009 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS Microsoft Office 2007 SP2 Systems"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS Microsoft Silverlight 2.0 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS Microsoft Silverlight 3.0 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS Microsoft Silverlight 4.0 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS Sales Development Associate (SDA) Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS SAP 7.1 Computers~1~"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS SAP 7.1 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS SAPLOGON CMMS Configuration Cabarrus Changes"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS SAPLOGON HRP SHARP Cleanup Configuration Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS SQL Server 2005 Management Studio Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS SUN JRE 1.4.1 & 1.5.0 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS Windows XP SP2 Installation Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS Windows XP SP3 Stage Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU SMS Zebra Printer Installer 1.0 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="PMU UR OneDrive Users (GEO)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Remove PST Mappings from Shared Workstations"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="RETIRED - AppFac-SW VanDyke Software SecureCRT 7.3.6 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Sales AGDC Win10 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Sales Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="sd test collection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Semi-Managed Systems"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Servers - Patching - Pilot (WSS Managed Servers)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SG-Test-DeviceCollection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW ABBY Reader Professional 8.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Active X Print Control 1.0 AGDC Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Adobe Acrobat 5 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Adobe Acrobat Professional 7.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Adobe Acrobat Professional 9.1.3 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Adobe Creative Suite 6.0 Design & Web Premium"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Adobe Creative Suite 6.0 Master Collection ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Adobe Dreamweaver CS6 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Adobe Illustrator CS6 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Adobe Illustrator Removal 1.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Adobe InDesign 2.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Adobe Photoshop 7.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Adobe Photoshop CS2 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW ALCS LAP Reset Utility 2.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW ALCS LAP Reset Utility ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW ALCS MIPIA 3.5i ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW ALCS MIPIA 3.5k ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW ALSLS SSL Registry Fix 1.0 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Alterna TIFF 2.0.4.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Altria Workstation Refresh Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Altria Workstation Refresh Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Apple Quicktime Player 7.7.3 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Apple Quicktime Player 7.7.4 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Autodesk DWG TrueView 2012 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Autodesk Printer Discovery 1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Autodesk Suite 2012 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Autodesk Suite 2012(x64) ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Azalea Tools C128 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Bentley BentleyView 8.05 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW BNA Corporate Tax Analyzer 2012.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Brio Report Viewer 6.21 for Sybase ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW CA ERwin Data Modeler 9.37 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Captaris RightFax 9.4 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW CDM 4.4 patch ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Cisco Agent Desktop (Agent) 6.0.1.15 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Citrix Web Client 12.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW CLCBio CLCSequenceViewer 6.9 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Crystal Reports Enterprise 11.5 r2 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Crystal Reports Runtime 9.0 (LIMS ONLY) ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Crystal Reports Runtime 9.1 (LIMS ONLY) ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW DataWorks Active Factory 8.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW DEP REG Fix 1.0 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW DirectLink 6.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW eContract Signature Pad v1.0.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW EMS Front End 1.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW EndPoint Encryption Administration Console 5.2.5 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW ENT Windows Analytics Configuration Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Enterprise Site Discovery DISABLE 2.0.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW EPB2 Web Reporting ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW EPK Client Controls 6.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Epson WP-4533 Driver Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW FloorSol Util 1.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW GoToMeeting 4.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Honeywell MAXPRO r300 SP2 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Honeywell ProWatch 3.81 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Honeywell ProWatch Software Suite 4.2.0.10835 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Honeywell ProWatch Software Suite 4.2.0.10835 ENT Users_W10PRJ"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW HP DDM Inventory Agent 9.31 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW HP Quality Center Client 10.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW HyperSnap 6.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW IBM iSeries Access v5.3 UST Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW IBM ThinInstaller Update 1.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW ICIS 4.4 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW IE Utility 1.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Ignite Clearstream Removal 1.0.0 AGDC Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Information Mapping FS Pro 4.2.0.5 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Intel Management Engine 9.5 Software 9.5.22.1760 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Intel Wireless Bluetooth 4.0 Adapter Software 17.0.1401.0428 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW IRI PlusSuite Removal 1.0.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW iRise Reader 8.10 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW JDE Worldvision 2.0 UST Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW LabWare AA PROD LIMS 3.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW LabWare AA QA LIMS 5.02 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW LabWare FLV QA LIMS 3.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW LabWare Lims Base Package 5.02 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW LabWare NP PROD LIMS 2.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Lakeside Systrack 8.2 Troubleshooting Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW LaserFiche Quest Workspace Client 7.6.305.791 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Lenovo Hotkey Features Integration 5.30.0000 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Lenovo Intel Wireless LAN Driver Only 16.8.0 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Lenovo Power Management Driver 1.67.04.05 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Sw Lenovo ThinkVantage Communcations Utility 3.1.12.0 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Lexmark Pro 901 Drivers ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Liquidware Labs Connector ID Removal ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW MalwareBytes Uninstaller ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Microsoft 2003 Admin Tools ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Microsoft MapPoint 2011 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Microsoft Project 2002 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Microsoft Project 2003 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Microsoft SCCM 2012 Client ALCS Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Microsoft SCCM 2012 Client Push ALCS Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Microsoft Silverlight 5.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Microsoft Skype for Business 2016 (MSI) x86 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Microsoft Skype for Business Pro Plus Version 1803 x86 (leave Lync) Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Microsoft SQL Server Express 2008 R2 (x64) AGDC Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Microsoft Visio Professional 2002 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Microsoft Visio Standard 2002 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Microsoft Visio Standard 2010 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Microsoft Windows 7 SP1 Only Installation ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Minitab v16 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Moisture Meter 1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW NetOp Remote Control 6.50 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW NetPlay Instant Demo v1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW NetVu Observer 1.9.4 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW ODBC importCSV 1.0 ALSD Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW OmniRim Barcode Fonts 2.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW OmniRim Barcode Fonts 2.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW OneDrive Directory Utility 1.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW OpenText WebDAV ActiveX Control 10.0 AGDC Win7 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Oracle Basic Instant Client 11.2 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Oracle Client 9.2.0.8 Removal LIMS ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Oracle Java Runtime Env SE 1.6.0.43 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW OTS FloorSol 5.24 PMUSA Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW OTS FloorSol 5.24 PMUSA Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW PCOE 1.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW PMCS 3.3.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW PMDF Customer Supplied Data v2.0.0 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Print From Home POC Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Prometheus Work Order Print Manager 3.6.45 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW QlikTech QlikViewDesktop 10 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW RealLegal eViewer 6.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Realtek High Definition Audio Driver 6.0.1.7188 Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW RNA 7.01 Update ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW SalesEdge Training 2.0.0 AGDC Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW SAP 7.2 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW SAP GUI for Windows 7.20 OFFLINE INSTALLATION Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW SAP WWI 2.7b ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW SAPGuiXT with Input Assistant 4.2.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW SAS System Viewer 8.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW SCCM 2012 SP1 CU2 - Client Update x64"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW SMART SynchronEyes Student Edition 7.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Snagit 6.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Sopheon Accolade 8.4.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW StagFont 1.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Statgraphics Plus 5.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW TeamMate R10.2.4 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW TechSmith Snagit 9.11 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW TFPP (PMUSA) 5.0.0.27 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW ThinkPad Camera GRANT ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW TimeVision OrgPublisher 4.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW tricerat ScrewDrivers 4.6 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Trouble Call RIC Workstation 5.0 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Uninstall Adobe Acrobat Professional 11.0.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Uninstall Honeywell ProWatch Software Suite 4.2.0.10835 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Vantive Client 8.5.5.13v2 (Contact & TCM) ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Vantive Client 8.5.5.13v2 (Contact) ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Vantive Client 8.5.5.13v2 (Helpdesk & TCM) ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Vantive Client 8.5.5.13v2 (HR & HD & Contact & TCM) ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Vantive Client 8.5.5.13v2 (TCM) ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Vantive Client Removal 1.0 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Vantive Developer 8.5.5.13 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW WBS Chart Pro 4.9 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW Windows 10 Application Testing Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="SW WinSCP 4.2.5 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Test - AD Group Collection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Test - Applications (Hardware Refresh)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST - Chris Avery test collection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST - Chris Muller test collection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST - EXCLUDE CP TEST CRT Lab Pilot Pilot Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST - Field Sales Force Testing User Accounts"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST - Fix DNS Issue"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Test - Kyle Uninstall Collection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST - MMS Shortcut Removal ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST - Muller Test Collection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST - Muller Test Collection 2 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST - Muller Test Collection Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST - Rima"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST - SMS Pilot"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST - Steve N test collection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Test - Steve N Test Collection 2"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST - SW Lync 2010 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST - TM Upgrade Pilot Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="test c++2013"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Test Device Collection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Test For Mike-N-Matt"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST- SW Microsoft SCCM 1602 Client Upgrade"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST_AppFac-SW LabWare LIMS 5.0.2.1 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST-ALCS Log File Collector Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST-Baseline for Internet Explorer MaxScriptStatements ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TESTCOMPUTERS-SERBIC"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST-CU2 Replication Test"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST-Dynamic Query All Workstation Without Reader9.5.5"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TESTING All Laptops"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TESTING Autodesk 2015"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TESTING ONLY-SW MFT Platform Server v7.1.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TESTING ONLY-SW Service Pack 2 for MS Office 2010 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST-Kyles Parent Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST-KylesStagingCollection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST-MS Updates June 2014 Task Sequence"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST-NEWCOMBTEST"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST-SW Adobe Core Software ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST-SW Citrix Receiver 4.6 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST-SW Internet Explorer HKCU Registry Fixes ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST-SW Microsoft App-V Client 5.1 ENT Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST-SW Microsoft Office 2016 (KMS) Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST-SW MindJet MindManager 15 UNINSTALL ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST-SW MTI POGTool UAT TESTERS Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST-SW Sales Launch Pad 4.6.0 AGDC Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TEST-WEISIGERTEST"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TS - Deploy Win10 1709 Upgrade - Required - V2.3 - Phase 10"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TS - Deploy Win10 1709 Upgrade - Required - V2.3 - Sales Dist. 1 - Group 1"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TS - Deploy Win10 1709 Upgrade - Required - V2.3 - Sales Dist. 1 - Group 3"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TS - Deploy Win10 1709 Upgrade - Required - V2.3 - Sales Dist. 1 - Group 5"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TS - Deploy Win10 1709 Upgrade - Required - V2.3 - Sales Dist. 2"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TS - Deploy Win10 1709 Upgrade - Required - V2.3 - Sales Dist. 2 - Group 2"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TS - Deploy Win10 1709 Upgrade - Required - V2.3 - Sales Dist. 2 - Group 4"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TS - Deploy Win10 1709 Upgrade - Required - V2.3 - Sales Dist. 3"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TS - Deploy Win10 1709 Upgrade - Required - V2.3 - Sales Dist. 3 - Group 2"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TS - Deploy Win10 1709 Upgrade - Required - V2.3 - Sales Dist. 3 - Group 4"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TS - Deploy Win10 1709 Upgrade - Required - V2.3 - Sales Dist. 4 - Group 1"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TS - Deploy Win10 1709 Upgrade - Required - V2.3 - Sales Dist. 4 - Group 3"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TS - Deploy Win10 1709 Upgrade - Required - V2.3 - Sales Dist. 4 - Group 5"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TS-AGDC Migration Testing (CDB)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TS-AGDC Migration Testing (JFH)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TS-Standalone Application Layering Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TS-SW AIS Data Migration Utility Computers"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - AD Group Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - AllWindows 10 Systems"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - AppFac-SW SAP GUI for Windows 7.20 ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - Deploy X1 Tablet BIOS Update 1.63 (Pilot 2)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - Deploy X1 Tablet BIOS Update 1.68-121"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - Deploy X1 Tablet Driver Package Update (Pilot 2)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - Deploy X1 Tablet Primax Keyboard Firmware Update"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - Deploy X1 Tablet Primax Keyboard Firmware Update (Round 2)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - EUO-Script-Testing ENT Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - EXE - Adobe Acrobat Reader v18.011.20035"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - EXE - Teradata Tools and Utiliteies v.15.10"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - Machine - AGDC Migration Log Copy Utility"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - MS Updates reporting (Required)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - MS Updates reporting (Required) CP CRT"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - Trend IPXFER Utilty - ALCRTPTRDP02"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - TS - Deploy Win10 1709 Upgrade - PILOT - Required - V2.3"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - TST - Tableau Desktop 9 Uninstall Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - Updates Reboots"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - Win 10 Project testing"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - X1 Tablet Stand-Alone Installer Testing Collection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - X1 Tablets to Exclude from Deployment"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - X1 Updates Pilot Collection - July 2017"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST - X1 Yoga Gen 2 - Sierra Wireless Driver Update"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="TST-UNINSTALL-MICROSOFT-OFFICE-PROJECT-DEV"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Unknown - Nov Patch Machines"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="UR ENT CITRIX WINDOWS 10 Application Testing Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="UR-Win10Ref-MyDocsCleanup"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="USMT Data Capture - Shared Workstations"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="USMT Data Capture - UNC Ver 2.0"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="USMT Data Capture - UNC Ver 4.0"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="USMT Data Restore and Application Install"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="USMT Data Restore and Application Install Ver 2.0 - YLT Test"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="USMT Data Restore and Application Install Ver 4.0"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="UST SMS IBM iSeries Access v5.3 Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="WDigest Reg Setting (REM)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Windows 10 Servicing - 1709 - Post Install Tasks"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Windows 10 Servicing - 1709 - QA_Available"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Windows 10 Servicing - PROD - Phase1"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Windows 10 Servicing - PROD - Phase2"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Windows 10 Servicing Parent Collection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Windows 10 Workstation Version 1607 LTSB Limiting Collection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Windows 8 Tablets"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Workstation - Patching - OSD Bois - PROD"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Workstation - Patching - OSD Bois - QA"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Workstations - Patch - Levnovo Patch Tool (TST)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Workstations - Patching - 3rd Party Java Development 8.181"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Workstations - Patching - Pilot (COC)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Workstations - Patching - Pilot (Existing)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Workstations - Patching - PROD (RCH)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Workstations - Patching - PROD (VDI2)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Workstations - Patching - PROD (VDI4)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Workstations - Patching - QA (WSS)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Workstations - Patching - Test (OOB)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Workstations - Patching Collections - Reports"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Workstations - Patching Out of Band (ODC) - KB4343894"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Workstations Excluded from SCCM Client Install"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Workstations NOT JUSTIFIED"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="Workstations with Client Is NULL"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="WSS Test Users"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="WSS Test VM's"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="X1 Limiting Collection"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="X1 Tablet August Updates (Pilot 2)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="X1 Tablet August Updates (Pilot 3)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="X1 Tablet Docking Station Firmware Update 2.33"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="X1 Tablet Gen 1 System Info"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="X1 Tablet Updates (Pilot 3)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append

$DeviceCollection ="X1 Tablet Updates Paulsboro (Pre-Deployment)"
$OutFilePath = "C:\Documents\Scripts\PowerShell\Collections\" + $DeviceCollection + ".txt"
$DeviceCollection | Out-File -FilePath $OutFilePath -Width 512 
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName | Out-File -FilePath $OutFilePath -Append