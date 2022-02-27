<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.172
	 Created on:   	3/5/2021 6:39 PM
	 Created by:   	Gary Brunner
	 Organization: 	
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		Display BIOS Information
#>

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

$CMPSSuppressFastNotUsedCheck = $true
   
$Form = New-Object System.Windows.Forms.Form
$Label = New-Object System.Windows.Forms.Label
$Cancel = New-Object System.Windows.Forms.Button
$Ok = New-Object System.Windows.Forms.Button
$TextBox = New-Object System.Windows.Forms.TextBox
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 108
$System_Drawing_Size.Width = 320
$Form.StartPosition = 'CenterScreen'
$Form.ClientSize = $System_Drawing_Size
$Form.DataBindings.DefaultDataSourceUpdateMode = 0
$Form.Name = "Form"
$Form.ShowIcon = $False
$Form.Text = "BIOS Information"

$Label.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 75
$System_Drawing_Point.Y = 10
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 180
$Label.Location = $System_Drawing_Point
$Label.Name = "label0"
$Label.Size = $System_Drawing_Size
$Label.Text = "Please enter a computer name"
$Form.Controls.Add($Label)

$TextBox.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 125
$System_Drawing_Point.Y = 35
$TextBox.Location = $System_Drawing_Point
$TextBox.Name = "TextBox"
$TextBox.Text = ""
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 80
$TextBox.Size = $System_Drawing_Size
$TextBox.TabIndex = 1
$Form.Controls.Add($TextBox)
    
$Cancel.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 45
$System_Drawing_Point.Y = 65
$Cancel.Location = $System_Drawing_Point
$Cancel.Name = "Cancel"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 75
$Cancel.Size = $System_Drawing_Size
$Cancel.TabIndex = 3
$Cancel.Text = "Cancel"
$Cancel.UseVisualStyleBackColor = $True
$Cancel.add_Click($handler_cancel_Click)
$Cancel.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
$Form.Controls.Add($Cancel)

$Ok.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 195
$System_Drawing_Point.Y = 65
$Ok.Location = $System_Drawing_Point
$Ok.Name = "Ok"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 75
$Ok.Size = $System_Drawing_Size
$Ok.TabIndex = 2
$Ok.Text = "Ok"
$Ok.UseVisualStyleBackColor = $True
$Ok.DialogResult=[System.Windows.Forms.DialogResult]::Ok
$Ok.add_Click($Ok_OnClick)
$Form.Controls.Add($Ok)

$InitialFormWindowState = $Form.WindowState
$Form.add_Load($OnLoadForm_StateCorrection)
$DialogResult = $Form.ShowDialog()
$DeviceName = $TextBox.Text

Get-WmiObject -Class Win32_BIOS -ErrorAction SilentlyContinue -ComputerName $DeviceName | Select SMBIOSBIOSVersion, SMBIOSMajorVersion, SMBIOSMinorVersion, BIOSVersion, Manufacturer, SerialNumber | Out-GridView -Title "$DeviceName BIOS Information"