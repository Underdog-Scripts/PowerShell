$SiteCode = "LW2"
$ProviderMachineName = "HIISCCLEW002.ingallscorp.com"
$initParams = @{}
if((Get-Module ConfigurationManager) -eq $null) {Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams}
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams}
Set-Location "$($SiteCode):\" @initParams

$DeviceList = Get-Content -Path "C:\Users\zzbrunnga\Documents\One.txt"

ForEach ($Device in $DeviceList) {
     Get-WmiObject Win32_OperatingSystem -ComputerName $Device -ErrorAction SilentlyContinue | Select Csname, @{Label="LastBootUpTime";Expression={$_.ConverttoDateTime($_.lastbootuptime)}}
     
}