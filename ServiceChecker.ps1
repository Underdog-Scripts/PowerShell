<#
.Synopsis
   This script performs a Service Check
.DESCRIPTION
   This script performs a Service Check
.PARAMETER List
   None
.INPUTS
   Text file of computer names or input
.OUTPUTS
   Screen display
.NOTES
   Version:        1.4
   Author:         GB
   Creation Date:  5.22.23
   Purpose/Change: Initial script development
   Modification Date: 5.23.23, 6.5.23, 6.7.23, 6.14.23, 6.15.23, 7.19.23, 7.20.23, 8.28.23,
                      9.1.23, 9.5.23, 9.21.23, 10.18.23, 10.19.23, 10.23.23, 1.6.24, 1.13.24,
                      10.18.24, 2.11.25
   Purpose/Change: Modified for USPS
.EXAMPLE
   ./ServiceChecker.ps1
.EXAMPLE
   ./ServiceChecker.ps1
#>

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Paths to files containing computer names and test results
$Loc = Get-Location

If (Test-Path ".\ServiceChecker.txt") {
    $ComputerList = ".\ServiceChecker.txt"
    # Read files with computer names
    $DeviceList = Get-Content -Path $ComputerList
}

# Display File Read Box #########################################################################################################
Function Dialog {
    $Form = New-Object System.Windows.Forms.Form
    $Form.ShowIcon=$False
    $Form.Text = "Service Checker"
    $Form.Size = New-Object System.Drawing.Size(700, 1020)
    $Form.StartPosition = "CenterScreen"

    # Creates output log
    $LogTextBox = New-Object System.Windows.Forms.RichTextBox
    $LogTextBox.Location = New-Object System.Drawing.Size(4, 4)
    $LogTextBox.Size = New-Object System.Drawing.Size(676, 968)
    $LogTextBox.ReadOnly = 'True'
    $LogTextBox.BackColor = 'Black'
    $LogTextBox.ForeColor = 'White'
    $LogTextBox.TabIndex = 1
    $LogTextBox.Text = "An input file was found here`n $Loc\ServiceChecker.txt`n`n Press Enter to Begin`n"
    $Font = [System.Drawing.Font]::New("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Regular)
    $Form.Font = $Font
    $Form.Controls.Add($LogTextBox)
    $LogTextBox.Add_KeyDown({
    If ($_.KeyCode -eq "Enter") {
        Script:CheckConnection $SystemName
    }

    })

    $Form.Add_Shown({$Form.Activate()})
    [void] $Form.ShowDialog()
}

# Display Name Prompt Box #######################################################################################################
Function Create-Form {
    $Form = New-Object System.Windows.Forms.Form
    $Form.ShowIcon=$False 
    $Form.Text = "Service Checker"
    $Form.Size = New-Object System.Drawing.Size(700, 1020) 
    $Form.StartPosition = "CenterScreen"

    # Create form label
    $Label = New-Object System.Windows.Forms.label
    $Label.Location = New-Object System.Drawing.Size(6, 8) 
    $Label.Size = New-Object System.Drawing.Size(400, 26)
    $Label.Text = "Type computer name below and press {Enter}"
    $Form.Controls.Add($Label)

    # Creates output log
    $LogTextBox = New-Object System.Windows.Forms.RichTextBox
    $LogTextBox.Location = New-Object System.Drawing.Size(4, 72) 
    $LogTextBox.Size = New-Object System.Drawing.Size(676, 900)
    $LogTextBox.ReadOnly = 'True'
    $LogTextBox.BackColor = 'Black'
    $LogTextBox.ForeColor = 'White'
    $LogTextBox.TabIndex = 2
    $Font = [System.Drawing.Font]::New("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Regular)
    $Form.Font = $Font
    $Form.Controls.Add($LogTextBox)

    # Device Input Box
    $DeviceTextBox = New-Object System.Windows.Forms.TextBox 
    $DeviceTextBox.Location = New-Object System.Drawing.Size(10, 37) 
    $DeviceTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $DeviceTextBox.Text = ""
    $DeviceTextBox.TabIndex = 1
    $Form.Controls.Add($DeviceTextBox)
    $DeviceTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Script:Connection $DeviceTextBox.Text
    }
    })

    # Clear output log Button
    $ClearButton = New-Object System.Windows.Forms.Button
    $ClearButton.Location = New-Object System.Drawing.Size(250, 34)
    $ClearButton.Size = New-Object System.Drawing.Size(150, 28)
    $ClearButton.Text = "Clear Screen"
    $ClearButton.Add_Click(
        {$LogTextBox.Clear()})
    $Form.Controls.Add($ClearButton)
    
    $Form.Add_Shown({$Form.Activate()})
    [void] $Form.ShowDialog()
}

# Update Screen Output ##########################################################################################################
Function Update-Log ($Text) {
    $LogTextBox.AppendText("$Text")
    $LogTextBox.Update()
    $LogTextBox.ScrollToCaret()
}

# Update Error Screen Output ####################################################################################################
Function Error-Log ($Text) {
    $LogTextBox.SelectionColor = 'Red'
    $LogTextBox.AppendText("$Text")
    $LogTextBox.Update()
    $LogTextBox.ScrollToCaret()
}

# Update Warning Screen Output ##################################################################################################
Function Warning-Log ($Text) {
    $LogTextBox.SelectionColor = 'Orange'
    $LogTextBox.AppendText("$Text")
    $LogTextBox.Update()
    $LogTextBox.ScrollToCaret()
}

# Update Header Screen Output ###################################################################################################
Function Header-Log ($Text) {
    $LogTextBox.SelectionColor = 'Magenta'
    $LogTextBox.Font.Bold
    $LogTextBox.AppendText("$Text")
    $LogTextBox.Update()
    $LogTextBox.ScrollToCaret()
}

# Perform Remote Session Work ###################################################################################################
Function CheckConnection {
    $LogTextBox.Clear()
    ForEach ($SystemName in $DeviceList) {
        $SystemName = $SystemName.ToUpper()
        Header-Log "$SystemName`n"
        Update-Log "Name - Status - StartType`n"
        $Session = New-PSSession -ComputerName $SystemName

        If (Test-Connection -ComputerName $SystemName -ErrorAction SilentlyContinue -Count 1 -Quiet) {
        # Open remote powershell session on remote computer
            $Session = New-PSSession -ComputerName $SystemName -ErrorAction SilentlyContinue

            If($Session -ne $Null) {
                # Check status of services ##########################################################################################
                $ServiceArray = "AdaptivaClient","BFE","BITS","CcmExec","LanManServer","LanManWorkstation","Riposte","WinMgmt"

                ForEach ($Service in $ServiceArray) {
                    $ScriptBlock1 = {param($Service)
                    (Get-Service -Name $Service -ErrorAction SilentlyContinue).Status}
                    $ServiceStatus = Invoke-Command -Session $Session -ScriptBlock $ScriptBlock1 -ArgumentList $Service

                    $ScriptBlock2 = {param($Service)
                    (Get-Service -Name $Service -ErrorAction SilentlyContinue).StartType}
                    $ServiceStartType = Invoke-Command -Session $Session -ScriptBlock $ScriptBlock2 -ArgumentList $Service
                    If($ServiceStatus -eq "1") {$ServiceStatus = "Stopped"}
                    If($ServiceStatus -eq "4") {$ServiceStatus = "Running"}
                    If($ServiceStartType -eq "2") {$ServiceStartType = "Automatic"}
                    If($ServiceStartType -eq "3") {$ServiceStartType = "Manual"}
                    If($ServiceStartType -eq "4") {$ServiceStartType = "Disabled"}
                    If($ServiceStatus -eq $Null) {$ServiceStatus = "Null"}
                    If($ServiceStartType -eq $Null) {$ServiceStartType = "Null"}

                    # Color code output for ease of reading
                    If ($ServiceStatus -eq "Null" -or $ServiceStartType -eq "Null"){
                        Error-Log "$Service - $ServiceStatus - $ServiceStartType`n"
                    } ElseIf ($ServiceStartType -eq "Disabled" -or $ServiceStartType -eq "Manual" -or $ServiceStatus -eq "Stopped") {

                        Warning-Log "$Service - $ServiceStatus - $ServiceStartType`n"
                    } Else {

                        Update-Log "$Service - $ServiceStatus - $ServiceStartType`n"
                    }

                }

            } Else {

                Error-Log "Session Connection Error`n`n"
            }

        } Else {

            Error-Log "Offline`n`n"
        }


    }
         
    Update-Log "`n"
    Update-Log "Script Complete`n"
}

Function Connection {
Param(
[ValidateNotNullOrEmpty()]
[string]$SystemName
)
    $LogTextBox.Clear()
    $SystemName = $SystemName.ToUpper()
    Header-Log "$SystemName`n"
    Update-Log "Name - Status - StartType`n"
    $Session = New-PSSession -ComputerName $SystemName

    If (Test-Connection -ComputerName $SystemName -ErrorAction SilentlyContinue -Count 1 -Quiet) {
    # Open remote powershell session on remote computer
        $Session = New-PSSession -ComputerName $SystemName -ErrorAction SilentlyContinue

        If($Session -ne $Null) {
            # Check status of services ##########################################################################################
            $ServiceArray = "AdaptivaClient","BFE","BITS","CcmExec","LanManServer","LanManWorkstation","Riposte","WinMgmt"

            ForEach ($Service in $ServiceArray) {
                $ScriptBlock1 = {param($Service)
                (Get-Service -Name $Service -ErrorAction SilentlyContinue).Status}
                $ServiceStatus = Invoke-Command -Session $Session -ScriptBlock $ScriptBlock1 -ArgumentList $Service

                $ScriptBlock2 = {param($Service)
                (Get-Service -Name $Service -ErrorAction SilentlyContinue).StartType}
                $ServiceStartType = Invoke-Command -Session $Session -ScriptBlock $ScriptBlock2 -ArgumentList $Service
                If($ServiceStatus -eq "1") {$ServiceStatus = "Stopped"}
                If($ServiceStatus -eq "4") {$ServiceStatus = "Running"}
                If($ServiceStartType -eq "2") {$ServiceStartType = "Automatic"}
                If($ServiceStartType -eq "3") {$ServiceStartType = "Manual"}
                If($ServiceStartType -eq "4") {$ServiceStartType = "Disabled"}
                If($ServiceStatus -eq $Null) {$ServiceStatus = "Null"}
                If($ServiceStartType -eq $Null) {$ServiceStartType = "Null"}

                # Color code output for ease of reading
                If ($ServiceStatus -eq "Null" -or $ServiceStartType -eq "Null"){
                    Error-Log "$Service - $ServiceStatus - $ServiceStartType`n"
                } ElseIf ($ServiceStartType -eq "Disabled" -or $ServiceStartType -eq "Manual" -or $ServiceStatus -eq "Stopped") {

                    Warning-Log "$Service - $ServiceStatus - $ServiceStartType`n"
                } Else {

                    Update-Log "$Service - $ServiceStatus - $ServiceStartType`n"
                }

            }

        } Else {

            Error-Log "Session Connection Error`n`n"
        }

    } Else {

        Error-Log "Offline`n`n"
    }

Update-Log "`n"
Update-Log "Script Complete`n"
}

If (Test-Path ".\ServiceChecker.txt") {
    Dialog
} Else {
    Create-Form
}

 