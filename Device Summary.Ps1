<#
.Synopsis
  This script displays Device Summary
.DESCRIPTION
  This script displays Device Summary
.PARAMETER List 
   None
.INPUTS
   User
.OUTPUTS
   Screen
.NOTES
   Version:        1.0
   Author:         GB
   Creation Date:  7.24.21
   Purpose/Change: Initial script development
   Modification Date: 8.30.21, 10.21.21, 2.2.22
   Purpose/Change: Modified for Ingalls Shipbuilding
.EXAMPLE
   ./Device Summary.ps1
.EXAMPLE
   ./Device Summary.ps1
#>

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

Function Update-Log ($Text) {
    $LogTextBox.AppendText("$Text")
    $LogTextBox.Update()
    $LogTextBox.ScrollToCaret()
}

Function Create-Form {
    $Form = New-Object System.Windows.Forms.Form
    $Form.ShowIcon=$False 
    $Form.Text = "Device Summary"
    $Form.Size = New-Object System.Drawing.Size(450, 890) 
    $Form.StartPosition = "CenterScreen"

    # Create form label
    $Label = New-Object System.Windows.Forms.label
    $Label.Location = New-Object System.Drawing.Size(6, 8) 
    $Label.Size = New-Object System.Drawing.Size(360, 28)
    $Label.Text = "Type computer name below and press {Enter}"
    $Form.Controls.Add($Label)

    # Creates output log
    $LogTextBox = New-Object System.Windows.Forms.RichTextBox
    $LogTextBox.Location = New-Object System.Drawing.Size(6, 76) 
    $LogTextBox.Size = New-Object System.Drawing.Size(422, 768)
    $LogTextBox.ReadOnly = 'True'
    $LogTextBox.BackColor = 'Black'
    $LogTextBox.ForeColor = 'White'
    $Font = New-Object System.Drawing.Font("",9,[System.Drawing.FontStyle]::Regular)
    $Form.Font = $Font
    $Form.Controls.Add($LogTextBox)

    # Device Input Box
    $DeviceTextBox = New-Object System.Windows.Forms.TextBox 
    $DeviceTextBox.Location = New-Object System.Drawing.Size(10, 42) 
    $DeviceTextBox.Size = New-Object System.Drawing.Size(200, 28)
    $DeviceTextBox.Text = ""
    $DeviceTextBox.TabIndex = 0
    $Form.Controls.Add($DeviceTextBox)
    $DeviceTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Script:CheckConnection $DeviceTextBox.Text

    }
    })
    
    # Clear Log Button
    $ClearButton = New-Object System.Windows.Forms.Button
    $ClearButton.Location = New-Object System.Drawing.Size(230, 38)
    $ClearButton.Size = New-Object System.Drawing.Size(120, 28)
    $ClearButton.Text = "Clear Screen"
    $ClearButton.Add_Click(
    {$LogTextBox.Clear()})
    $Form.Controls.Add($ClearButton)

    $Form.Add_Shown({$Form.Activate()})
    [void] $Form.ShowDialog()
}

Function CheckConnection {
[CmdletBinding()]
Param ([Parameter(Mandatory = $False, ValueFromPipeline = $False)]
$DeviceList = @()
)

       ForEach ($Device in $DeviceList) {
        If (!$Device) { 
            Update-Log "Entry is NULL `n"
        } Else {
            
            If (Test-Connection -ComputerName $Device -Count 1 -Quiet) {
                Update-Log "Working ... `n"
                Script:GetInfo0 $DeviceTextBox.Text
                Script:GetInfo1 $DeviceTextBox.Text
                Script:GetInfo2 $DeviceTextBox.Text
                Script:GetInfo3 $DeviceTextBox.Text
                Script:GetInfo4 $DeviceTextBox.Text
                Script:GetInfo5 $DeviceTextBox.Text
                           
            } Else {
                Update-Log "$Device ; Offline `n"
                 
            }

        }

    }

}

Function GetInfo0 {
    [CmdletBinding()]
    Param ([Parameter(Mandatory = $False, ValueFromPipeline = $False)]
    $DeviceList = @()
       )

       ForEach ($Device in $DeviceList) {
        $ScriptOutput = ""
        $ScriptOutput = (Get-WmiObject Win32_ComputerSystem -ErrorAction SilentlyContinue -ComputerName $Device).Manufacturer
        Update-Log "Manufacturer= $ScriptOutput `n"
        $ScriptOutput = ""
        $ScriptOutput = (Get-WmiObject Win32_ComputerSystem -ErrorAction SilentlyContinue -ComputerName $Device).Model
        Update-Log "Model= $ScriptOutput `n`n"
        $ScriptOutput = ""
        $ScriptOutput = Get-WmiObject -Class Win32_BIOS -ErrorAction SilentlyContinue -ComputerName $Device |`
        Select @{Label="BIOS Version";Expression={($_.SMBIOSBIOSVersion)}},`
        @{Label="BIOS Major Version";Expression={($_.SMBIOSMajorVersion)}},`
        @{Label="BIOS Minor Version";Expression={($_.SMBIOSMinorVersion)}},`
        Manufacturer, SerialNumber
        $ScriptOutput = $ScriptOutput -Replace '; ', "`n"
        $ScriptOutput = $ScriptOutput -Replace '@{', ''
        $ScriptOutput = $ScriptOutput -Replace '}', "`n`n"
        Update-Log $ScriptOutput

       }      

}

Function GetInfo1 {
    [CmdletBinding()]
    Param ([Parameter(Mandatory = $False, ValueFromPipeline = $False)]
    $DeviceList = @()
       )

       ForEach ($Device in $DeviceList) {
        $ScriptOutput = ""
        $ScriptOutput = Get-WmiObject -Class Win32_VideoController -ErrorAction SilentlyContinue -ComputerName $Device |`
        Select @{Label="Video Card"; Expression={($_.Caption)}},`
        @{Label="Driver Version"; Expression={($_.DriverVersion)}},`
        @{Label="Driver Date"; Expression={($_.DriverDate)}},`
        @{Label="Processor"; Expression={($_.VideoProcessor)}},`
        @{Label="Mode"; Expression={($_.VideoModeDescription)}}
        $ScriptOutput = $ScriptOutput -Replace '; ', "`n"
        $ScriptOutput = $ScriptOutput -Replace '@{', ''
        $ScriptOutput = $ScriptOutput -Replace '}', "`n`n"
        Update-Log $ScriptOutput

       }      

}

Function GetInfo2 {
    [CmdletBinding()]
    Param ([Parameter(Mandatory = $False, ValueFromPipeline = $False)]
    $DeviceList = @()
       )

       ForEach ($Device in $DeviceList) {
        $ScriptOutput = ""
        $ScriptOutput = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Device |`
        Select @{Label="Build Number"; Expression={($_.BuildNumber)}}
        $ScriptOutput = $ScriptOutput -Replace '; ', "`n"
        $ScriptOutput = $ScriptOutput -Replace '@{', ''
        $ScriptOutput = $ScriptOutput -Replace '}', "`n`n"
        Update-Log $ScriptOutput

       }      

}

Function GetInfo3 {
    [CmdletBinding()]
    Param ([Parameter(Mandatory = $False, ValueFromPipeline = $False)]
    $DeviceList = @()
       )

       ForEach ($Device in $DeviceList) {
        $ScriptOutput = ""
        $ScriptOutput = Get-WmiObject -Class Win32_LogicalDisk -ErrorAction SilentlyContinue -ComputerName $Device |`
        Select @{ Label = "Drive"; Expression = { ($_.DeviceId) } },`
        @{Label="Size(GB)"; Expression={ [math]::Round($_.Size/1GB, 2)}},`
        @{Label="Free Space(GB)"; Expression={[math]::Round($_.FreeSpace/1GB, 2)}}
        $ScriptOutput = $ScriptOutput -Replace '; ', "`n"
        $ScriptOutput = $ScriptOutput -Replace '@{', ''
        $ScriptOutput = $ScriptOutput -Replace '}', "`n`n"
        Update-Log $ScriptOutput

       }      

}

Function GetInfo4 {
    [CmdletBinding()]
    Param ([Parameter(Mandatory = $False, ValueFromPipeline = $False)]
    $DeviceList = @()
       )

       ForEach ($Device in $DeviceList) {
        $ScriptOutput = ""
        $ScriptOutput = Get-WmiObject -Class Win32_SystemEnclosure -ErrorAction SilentlyContinue -ComputerName $Device |`
        Select @{ Label = "Chassis type `n1 Other `n2 Unknown `n3 Desktop `n6 Mini Tower `n8 Portable `n9 Laptop `n10 Notebook `n11 Handheld `n30 Tablet `nThis system"; Expression={($_.ChassisTypes)}}
        $ScriptOutput = $ScriptOutput -Replace '; ', "`n"
        $ScriptOutput = $ScriptOutput -Replace '@{', ''
        $ScriptOutput = $ScriptOutput -Replace '}', "`n`n"
        Update-Log $ScriptOutput

       }      

}

Function GetInfo5 {
    [CmdletBinding()]
    Param ([Parameter(Mandatory = $False, ValueFromPipeline = $False)]
    $DeviceList = @()
       )

       ForEach ($Device in $DeviceList) {
        $ScriptOutput = ""
        $ScriptOutput = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Device |`
        Select @{Label="Last BootUp Time"; Expression={$_.ConverttoDateTime($_.lastbootuptime)}}
        $ScriptOutput = $ScriptOutput -Replace '; ', "`n"
        $ScriptOutput = $ScriptOutput -Replace '@{', ''
        $ScriptOutput = $ScriptOutput -Replace '}', "`n"
        Update-Log $ScriptOutput

       }      

}

Create-Form 
