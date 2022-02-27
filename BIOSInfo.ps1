<#
$IsPasswordSet = (Get-Item -Path DellSmbios:\Security\IsAdminPasswordSet).currentvalue 
If($IsPasswordSet -eq $true) {
    Write-Host "Password is configured"
} Else {
    Write-Host "No BIOS password"
} 
#>

Function Get_Dell_BIOS_Settings {
    $WarningPreference='silentlycontinue'
    If (Get-Module -ListAvailable -Name DellBIOSProvider) {
    } Else {
    
        Install-Module -Name DellBIOSProvider -Force
    }
  
    Get-Command -module DellBIOSProvider | out-null
    $Script:Get_BIOS_Settings = Get-ChildItem -path DellSmbios:\ | Select-Object category | 
  
    ForEach {
        Get-ChildItem -path @("DellSmbios:\" + $_.Category)  | Select-Object attribute, currentvalue 
    } 

    $Script:Get_BIOS_Settings = $Get_BIOS_Settings |  % { 
        New-Object psobject -Property @{
        Setting = $_."attribute"
        Value = $_."currentvalue"
    }}  | select-object Setting, Value 
    $Get_BIOS_Settings
} 
 
Get_Dell_BIOS_Settings | Out-GridView

#Set-Item -Path Dellsmbios:\POSTBehavior\FnLock -Value Enabled
Set-Item -Path Dellsmbios:\POSTBehavior\Camera -Value Disabled

<#
$IsPasswordSet = Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class HP_BIOSSettingInterface
If($IsPasswordSet -eq 1) {
    Write-Host "Password is configured"
} Else {
    Write-Host "No BIOS password"
} 

Function Get_HP_BIOS_Settings { 
    $Script:Get_BIOS_Settings = Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class hp_biosEnumeration -ErrorAction SilentlyContinue |  % { 
        New-Object psobject -Property @{    
        Setting = $_."Name"
        Value = $_."currentvalue"
        Available_Values = $_."possiblevalues"
        }}  | Select-Object Setting, Value, possiblevalues
    $Get_BIOS_Settings
}

Get_HP_BIOS_Settings | Out-GridView

#$BIOS = Get-WmiObject -Namespace root/hp/instrumentedBIOS -Class HP_BIOSSettingInterface
#$BIOS.setbiossetting("Fast Boot", "Disable","") | out-null 

#>
 