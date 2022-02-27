<#           
.NOTES
===========================================================================
Created with:    SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.173
Created on:      3/17/2021 3:04 PM
Modified on:     4/8/2021 12:08 PM, 4/9/2021 5:26 PM, 10/5/2021 4:42 PM
Created by:      Gary Brunner
Organization:    Gary Brunner
Filename:        AppsOnAComputer.ps1 
===========================================================================
.DESCRIPTION
Provides listing for application installs on a computer
#>

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

Function Create-Form {

    # Create Form
    $Form = New-Object System.Windows.Forms.Form
    $Form.ShowIcon=$False
    $Form.Font = 'Microsoft Sans Serif,11' 
    $Form.Text = "App installs on a computer"
    $Form.Size = New-Object System.Drawing.Size(450, 166) 
    $Form.StartPosition = "CenterScreen"

    # Input Box Label
    $ComputerID = New-Object System.Windows.Forms.Label 
    $ComputerID.Location = New-Object System.Drawing.Size(8, 8) 
    $ComputerID.Size = New-Object System.Drawing.Size(424, 26)
    $ComputerID.Text = "Enter computer name and press Enter"
    $ComputerID.Font = 'Microsoft Sans Serif,11'
    $Form.Controls.Add($ComputerID)
    
    # Input Box
    $ComputerIDTextBox = New-Object System.Windows.Forms.TextBox 
    $ComputerIDTextBox.Location = New-Object System.Drawing.Size(10, 48) 
    $ComputerIDTextBox.Size = New-Object System.Drawing.Size(280, 28)
    $ComputerIDTextBox.Font = 'Microsoft Sans Serif,11'
    $Form.Controls.Add($ComputerIDTextBox)
    $ComputerIDTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter"){
        Script:GetID $ComputerIDTextBox.Text
    }})
    
    # Comment Box
    $Comment = New-Object System.Windows.Forms.Label 
    $Comment.Location = New-Object System.Drawing.Size(8, 84) 
    $Comment.Size = New-Object System.Drawing.Size(250, 26)
    $Comment.Text = "(Should take less than 60 seconds)"
    $Comment.Font = 'Microsoft Sans Serif,9'
    $Form.Controls.Add($Comment)

    # Check Box
    $CheckBox = New-Object System.Windows.Forms.CheckBox
    $CheckBox.location = New-Object System.Drawing.Point(262,82)
    $CheckBox.Text = "Send to .csv file"
    $CheckBox.Size = New-Object System.Drawing.Size(200, 26)
    $CheckBox.Font = 'Microsoft Sans Serif,10'
    $Form.Controls.Add($Checkbox)

    # Close Button
    $CloseButton = New-Object System.Windows.Forms.Button
    $CloseButton.Location = New-Object System.Drawing.Point(300, 44)
    $CloseButton.Size = New-Object System.Drawing.Size(120, 30)
    $CloseButton.Font = 'Microsoft Sans Serif,11'
    $CloseButton.Text = "Close"
    $CloseButton.Add_Click({$Form.Close()
    })
    $Form.Controls.Add($CloseButton)
    $Form.Add_Shown({$Form.Activate()
    })
    [void] $Form.ShowDialog()
}

Function GetID([String]$ReportName) {

Invoke-Sqlcmd -QueryTimeout 500 -ConnectionTimeout 500 -Query "use CM_LW2
Select distinct SYS.Netbios_Name0 as 'System ID',
SYS.User_Name0 as 'Username',
ISW.ARPDisplayName0 as 'Add Remove Programs',
ISW.InstallDate0 as 'Date of Install',
ISW.InstallSource0 as 'Source Path',
ISW.ProductName0 as 'Name',
ISW.ProductVersion0 as 'Version',
ISW.UninstallString0 as 'Uninstall String'
FROM v_GS_INSTALLED_SOFTWARE ISW
JOIN v_R_System SYS on ISW.ResourceID = SYS.ResourceID
WHERE SYS.Netbios_Name0 like $ReportName
Order by SYS.Netbios_Name0,
ISW.ProductName0,
ISW.ProductVersion0" | Export-CSV -Path ".\Applications$ReportName.csv" -NoClobber -NoTypeInformation -Append
# | Out-GridView -Title "Applications on $ReportName"

}

Create-Form