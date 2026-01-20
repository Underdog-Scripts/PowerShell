<#
.Synopsis
    This script adds or removes systems from a collection
.DESCRIPTION
    This script adds or removes systems from a collection
.PARAMETER List 
    Collection name
.INPUTS
    Collection name
.OUTPUTS
    Output to SCCM server
.NOTES
    Version:        1.0
    Author:         GTB
    Creation Date:  4.20.25
    Purpose/Change: Initial script development
    Update Date:    4.20.25, 5.9.25
    Purpose/Change: Modified to use lmenu listbox
.EXAMPLE
    ./CollectionMenu.ps1
.EXAMPLE
    ./CollectionMenu.ps1
#>

# Where the dialog box is created
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Form = New-Object System.Windows.Forms.Form
$Form.Text = 'Select a Collection'
$Form.Size = New-Object System.Drawing.Size(525,285)
$Form.StartPosition = 'CenterScreen'

$AddButton = New-Object System.Windows.Forms.Button
$AddButton.Location = New-Object System.Drawing.Point(70,180)
$AddButton.Size = New-Object System.Drawing.Size(120,30)
$AddButton.Text = 'Add Systems'
$AddButton.DialogResult = [System.Windows.Forms.DialogResult]::Ok
$Form.AcceptButton = $AddButton
$Form.Controls.Add($AddButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(330,180)
$CancelButton.Size = New-Object System.Drawing.Size(120,30)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$Form.CancelButton = $CancelButton
$Form.Controls.Add($CancelButton)

$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(200, 180)
$Button.Size = New-Object System.Drawing.Size(120,30)
$Button.Text = 'Remove Systems'
$Button.DialogResult = [System.Windows.Forms.DialogResult]::No
$Form.Button = $Button
$Form.Controls.Add($Button)

$Label = New-Object System.Windows.Forms.Label
$Label.Location = New-Object System.Drawing.Point(6,6)
$Label.Size = New-Object System.Drawing.Size(510,23)
$Label.Text = 'Please select a collection - System names read from file'
$Form.Controls.Add($Label)

$Label0 = New-Object System.Windows.Forms.Label
$Label0.Location = New-Object System.Drawing.Point(6, 30)
$Label0.Size = New-Object System.Drawing.Size(516, 23)
$Label0.Text = 'E:\Packages\Scripts\[From To]Collection.txt'
$Form.Controls.Add($Label0)

$ListBox = New-Object System.Windows.Forms.ListBox
$ListBox.Location = New-Object System.Drawing.Point(10,60)
$ListBox.Size = New-Object System.Drawing.Size(482,20)
$ListBox.Height = 120

# ---------- Beginning of crap needed to talk to SCCM
# Site configuration
$SiteCode = "RSS" # Site code
$ProviderMachineName = "EAGNMNWBD13B.DEVSUB.DEV.DCE.USPS.GOV" # SMS Provider machine name
# Customizations
$initParams = @{}
# Do not change anything below this line
# Import the ConfigurationManager.psd1 module
If($null -eq (Get-Module ConfigurationManager)) {Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams}
# Connect to the site's drive if it is not already present
If($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams}
# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams
# ---------- End of crap needed to talk to SCCM

# Retrieves the last 50 collection names from SCCM
$CollectionList = Get-CMDeviceCollection -Name Retail-202* | Select-Object -Property Name | Select-Object -Last 50

# Remove extra stuff from the get list of collections
ForEach ($Item in $CollectionList) {
    $Item = $Item -Replace '@{', ''
    $Item = $Item -Replace '}', ''
    [void] $ListBox.Items.Add($Item)
}

# To use a testing collection
[void] $ListBox.Items.Add("Testing Collection")

# Presents the dialog box to the user
$Form.Controls.Add($ListBox)
$Form.Topmost = $true
$Result = $Form.ShowDialog()

# Log list of members before execution
$Now = Get-Date
$PackageName = $ListBox.SelectedItem
Write-Output "$Now - $PackageName - BEFORE Members List" | Out-File "E:\Packages\Scripts\$PackageName.log" -NoClobber -Append
Get-CMCollectionMember -CollectionName "$PackageName" | Select-Object -Property Name | Out-File "E:\Packages\Scripts\$PackageName.log" -NoClobber -Append

# When Add Systems button is pressed ...
If ($Result -eq [System.Windows.Forms.DialogResult]::Ok){
    $ComputerList = Get-Content "E:\Packages\Scripts\ToCollection.txt"
    $PackageName = $ListBox.SelectedItem
    Write-Host $PackageName "Add"
    ForEach ($Computer in $ComputerList) {
        Add-CMDeviceCollectionDirectMembershipRule -CollectionName "$PackageName" -ResourceID (Get-CMDevice -Name $Computer).ResourceID
    }

}

# When Remove Systems button is pressed ...
If ($Result -eq [System.Windows.Forms.DialogResult]::No){
    $ComputerList = Get-Content "E:\Packages\Scripts\FromCollection.txt"
    $PackageName = $ListBox.SelectedItem
    Write-Host $PackageName "Remove"
    ForEach ($Computer in $ComputerList) {
        Remove-CMDeviceCollectionDirectMembershipRule -CollectionName "$PackageName" -ResourceName $Computer -Force
    }

}

# Log list of members after execution 
$Now = Get-Date
Write-Output "$Now - Start-Sleep -Seconds 10" | Out-File "E:\Packages\Scripts\$PackageName.log" -NoClobber -Append
Start-Sleep -Seconds 10

$Now = Get-Date
Write-Output "$Now - Invoke-CMCollectionUpdate -Name $PackageName" | Out-File "E:\Packages\Scripts\$PackageName.log" -NoClobber -Append
Invoke-CMCollectionUpdate -Name "$PackageName"

$Now = Get-Date
Write-Output "$Now - Start-Sleep -Seconds 10" | Out-File "E:\Packages\Scripts\$PackageName.log" -NoClobber -Append
Start-Sleep -Seconds 10

$Now = Get-Date
Write-Output "$Now - $PackageName - AFTER Members List" | Out-File "E:\Packages\Scripts\$PackageName.log" -NoClobber -Append
Get-CMCollectionMember -CollectionName "$PackageName" | Select-Object -Property Name | Out-File "E:\Packages\Scripts\$PackageName.log" -NoClobber -Append   

$Now = Get-Date
Write-Output "$Now - Script complete" | Out-File "E:\Packages\Scripts\$PackageName.log" -NoClobber -Append
    