<#
.Synopsis
   This script allows for the mapping of drives or execution of programs with different user credentials
.DESCRIPTION
   This script allows for the mapping of drives or execution of programs with different user credentials
.EXAMPLE
   ./ShareMenu.ps1
.EXAMPLE
   ./ShareMenu.ps1
#>

# Updated with IP addresses 6.20.18

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Check for username text file
If (Test-Path $env:TEMP\ShareMenu.txt){
	$User = Get-Content -Path $env:TEMP\ShareMenu.txt
}

# Check if username has "zz" in it
If ($User -Notmatch 'zz'){
	[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
	$User = [Microsoft.VisualBasic.Interaction]::InputBox("The current ID = $User `r`n `r`nTo change, delete the $env:TEMP\ShareMenu.txt file", "New User Detected", "Enter new ID")
	Set-Content -Path $env:TEMP\ShareMenu.txt -Value $User
}

$form = New-Object System.Windows.Forms.Form
$form.Text = ''
$form.Size = New-Object System.Drawing.Size(200,490)
$form.StartPosition = 'CenterScreen'

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(12,404)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(94,404)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,10)
$label.Size = New-Object System.Drawing.Size(270,20)
$label.Text = 'Please select a Share:'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,30)
$listBox.Size = New-Object System.Drawing.Size(163,20)
$listBox.Height = 376

# List of items to present in dialog box
[void] $listBox.Items.Add('COVRSICES-NAS04')
[void] $listBox.Items.Add('10.192.2.53')
[void] $listBox.Items.Add('COVSMICES-WPS02')
[void] $listBox.Items.Add('10.192.32.68')
[void] $listBox.Items.Add('DEV-SOFTENG')
[void] $listBox.Items.Add('10.194.9.66')
[void] $listBox.Items.Add('NG00181421')
[void] $listBox.Items.Add('NG00222429')
[void] $listBox.Items.Add('10.194.9.216')
[void] $listBox.Items.Add('WAP01237')
[void] $listBox.Items.Add('10.192.32.105')
[void] $listBox.Items.Add('WAP01239')
[void] $listBox.Items.Add('10.192.32.107')
[void] $listBox.Items.Add('WAP01559')
[void] $listBox.Items.Add('10.192.32.18')
[void] $listBox.Items.Add('WAP01761')
[void] $listBox.Items.Add('10.192.32.20')
[void] $listBox.Items.Add('WAP03492')
[void] $listBox.Items.Add('10.194.11.14')
[void] $listBox.Items.Add('MDT')
[void] $listBox.Items.Add('EXPLORER')
[void] $listBox.Items.Add('HYPER-V')
[void] $listBox.Items.Add('MDT_MSC')
[void] $listBox.Items.Add('POWERSHELL')
[void] $listBox.Items.Add('POWERSHELL_ISE')
[void] $listBox.Items.Add('CMD_PROMPT')
[void] $listBox.Items.Add('IVANTI')

# Grab item selected my user
$form.Controls.Add($listBox)
$form.Topmost = $true
$result = $form.ShowDialog()
$Share = ( $listBox.SelectedItem )

# Convert displayed item to usable format for the New-PSDrive action
switch ( $Share ){
    COVRSICES-NAS04 {$Share = '\\COVRSICES-NAS04'}
	10.192.2.53     {$Share = '\\10.192.2.53'}
    COVSMICES-WPS02 {$Share = '\\covsmices-wps02.cov.virginia.gov'}
	10.192.32.68    {$Share = '\\10.192.32.68'}
    DEV-SOFTENG     {$Share = '\\DEV-SOFTENG'}
	10.194.9.66     {$Share = '\\10.194.9.66'}
    NG00181421      {$Share = '\\NG00181421\c$'}
	10.192.32.105   {$Share = '\\10.192.32.105\c$'}
    NG00222429      {$Share = '\\NG00222429\c$'}
	10.194.9.216    {$Share = '\\10.194.9.216\c$'}
    WAP01237        {$Share = '\\WAP01237'}
	10.192.32.105   {$Share = '\\10.192.32.105'}
    WAP01239        {$Share = '\\WAP01239'}
	10.192.32.107   {$Share = '\\10.192.32.107'}
    WAP01559        {$Share = '\\WAP01559'}
	10.192.32.18    {$Share = '\\10.192.32.18'}
    WAP01761        {$Share = '\\WAP01761\ldmain'}
	10.192.32.20    {$Share = '\\10.192.32.20\ldmain'}
    WAP03492        {$Share = '\\WAP03492'}
	10.194.11.14    {$Share = '\\10.194.11.14'}
    MDT             {$Share = '\\WAP03987\MDTMain'}
}

$MDT_LINE = """C:\Program Files\Microsoft Deployment Toolkit\Bin\DeploymentWorkbench.msc"""
$AVANTI_LINE = """C:\Program Files (x86)\RemotePackages\Console (1).rdp"""

# Check it share name has "\\" if not execute program rather than map a drive
If ($Share -Notmatch '\\'){
    switch ($Share){
        EXPLORER        {Start-Process  Explorer.exe -Verb RunAs}
		HYPER-V         {Start-Process -WindowStyle Hidden $Env:COMSPEC -workingdirectory $PSHOME -Credential $Env:USERDOMAIN\$User -ArgumentList "/c mmc.exe C:\Windows\System32\virtmgmt.msc"}
        MDT_MSC         {Start-Process -WindowStyle Hidden $Env:COMSPEC -workingdirectory $PSHOME -Credential $Env:USERDOMAIN\$User -ArgumentList "/c mmc.exe $MDT_LINE"}
        POWERSHELL      {Start-Process -WindowStyle Hidden $Env:COMSPEC -workingdirectory $PSHOME -Credential $Env:USERDOMAIN\$User -ArgumentList "/c Powershell.exe"}
        POWERSHELL_ISE  {Start-Process -WindowStyle Hidden $Env:COMSPEC -workingdirectory $PSHOME -Credential $Env:USERDOMAIN\$User -ArgumentList "/c Powershell_ise.exe"}
        CMD_PROMPT      {Start-Process $Env:COMSPEC -Verb RunAs}
		IVANTI			{Start-Process -WindowStyle Hidden $Env:COMSPEC -workingdirectory $PSHOME -ArgumentList "/c $AVANTI_LINE"}
        }
    
} else {

# Prompt for password and execute Explorer.exe
	If ($User -Match 'zz'){
        Remove-PSDrive $Share -ErrorAction SilentlyContinue
        New-PSDrive -Name "*" -PSProvider "FileSystem" -Root $Share -Credential $Env:USERDOMAIN\$User -ErrorAction SilentlyContinue
        Explorer $Share
		}
}