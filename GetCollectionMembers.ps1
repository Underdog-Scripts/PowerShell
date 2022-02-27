# Script to  create a text file for each collection with the members in tht file.

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
