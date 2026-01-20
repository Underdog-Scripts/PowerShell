# Site configuration
$SiteCode = "RSS" # Site code
$ProviderMachineName = "EAGNMNWBD13B.DEVSUB.DEV.DCE.USPS.GOV" # SMS Provider machine name

# Customizations
$initParams = @{}
#$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
#$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

# Do not change anything below this line

# Import the ConfigurationManager.psd1 module
if($null -eq (Get-Module ConfigurationManager)) {Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams}

# Connect to the site's drive if it is not already present
if($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams

$PackageName = Read-Host "Please enter package name"

Get-Content "C:\Users\xfnjp0\Documents\FromCollections.txt" | ForEach-Object {
    Remove-CMDeviceCollectionDirectMembershipRule -CollectionName "$PackageName" -ResourceID (Get-CMDevice -Name $_).ResourceID -Force }

