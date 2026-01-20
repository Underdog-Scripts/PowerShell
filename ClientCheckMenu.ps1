<#
.Synopsis
  This script performs a Service Check
.DESCRIPTION
  This script performs a Service Check
.PARAMETER List 
   None
.INPUTS
   Computer Name
.OUTPUTS
   Dialog display
.NOTES
   Version:        1.0
   Author:         GB
   Creation Date:  5.22.23
   Purpose/Change: Initial script development
   Modification Date: 5.23.23, 6.5.23, 6.7.23, 6.14.23, 6.15.23
   Purpose/Change: Modified for USPS
.EXAMPLE
   ./ServiceCheckMenu.ps1
.EXAMPLE
   ./ServiceCheckMenu.ps1
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
    $Form.Text = "Client Check"
    $Form.Size = New-Object System.Drawing.Size(470, 800) 
    $Form.StartPosition = "CenterScreen"

    # Create form label
    $Label = New-Object System.Windows.Forms.label
    $Label.Location = New-Object System.Drawing.Size(6, 8) 
    $Label.Size = New-Object System.Drawing.Size(400, 20)
    $Label.Text = "Type computer name below and press {Enter}"
    $Form.Controls.Add($Label)

    # Creates output log
    $LogTextBox = New-Object System.Windows.Forms.RichTextBox
    $LogTextBox.Location = New-Object System.Drawing.Size(4, 62) 
    $LogTextBox.Size = New-Object System.Drawing.Size(446, 694)
    $LogTextBox.ReadOnly = 'True'
    $LogTextBox.BackColor = 'Black'
    $LogTextBox.ForeColor = 'White'
    $LogTextBox.TabIndex = 2
    $Font = New-Object System.Drawing.Font("",10,[System.Drawing.FontStyle]::Regular)
    $Form.Font = $Font
    $Form.Controls.Add($LogTextBox)

    # Device Input Box
    $DeviceTextBox = New-Object System.Windows.Forms.TextBox 
    $DeviceTextBox.Location = New-Object System.Drawing.Size(10, 32) 
    $DeviceTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $DeviceTextBox.Text = ""
    $DeviceTextBox.TabIndex = 1
    $Form.Controls.Add($DeviceTextBox)
    $DeviceTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Script:CheckConnection "$DeviceTextBox.Text"

    }
    })

	# Clear output log Button
	$ClearButton = New-Object System.Windows.Forms.Button
	$ClearButton.Location = New-Object System.Drawing.Size(250, 28)
	$ClearButton.Size = New-Object System.Drawing.Size(120, 28)
	$ClearButton.Text = "Clear Screen"
	$ClearButton.Add_Click(
		{ $LogTextBox.Clear() })
	$Form.Controls.Add($ClearButton)
    
    $Form.Add_Shown({$Form.Activate()})
    [void] $Form.ShowDialog()
}

Function CheckConnection {
$System = $DeviceTextBox.Text

    If (Test-Connection -ComputerName $System -Count 1 -Quiet) {
        Update-Log "Computer Name = $System `n`n"

        # Open PowerShell remote session close at line 182
        New-PSSession -ComputerName $System
           
            # Check status of services
            $AdaptivaStatus = (Get-Service -Name "AdaptivaClient" -ErrorAction SilentlyContinue).Status
            $AdaptivaStartType = (Get-Service -Name "AdaptivaClient" -ErrorAction SilentlyContinue).StartType
            $BFEStatus = (Get-Service -Name "BFE" -ErrorAction SilentlyContinue).Status
            $BFEStartType = (Get-Service -Name "BFE" -ErrorAction SilentlyContinue).StartType
            $BITSStatus = (Get-Service -Name "BITS" -ErrorAction SilentlyContinue).Status
            $BITSStartType = (Get-Service -Name "BITS" -ErrorAction SilentlyContinue).StartType
            $CcmExecStatus = (Get-Service -Name "CcmExec" -ErrorAction SilentlyContinue).Status
            $CcmExecStartType = (Get-Service -Name "CcmExec" -ErrorAction SilentlyContinue).StartType
            $LanManServerStatus = (Get-Service -Name "LanManServer" -ErrorAction SilentlyContinue).Status
            $LanManServerStartType = (Get-Service -Name "LanManServer" -ErrorAction SilentlyContinue).StartType
            $LanManWorkstationStatus = (Get-Service -Name "LanManWorkstation" -ErrorAction SilentlyContinue).Status
            $LanManWorkstationStartType = (Get-Service -Name "LanManWorkstation" -ErrorAction SilentlyContinue).StartType
            $MSIServerStatus = (Get-Service -Name "MSIServer" -ErrorAction SilentlyContinue).Status
            $MSIServerStartType = (Get-Service -Name "MSIServer" -ErrorAction SilentlyContinue).StartType
            $RemoteRegistryStatus = (Get-Service -Name "RemoteRegistry" -ErrorAction SilentlyContinue).Status
            $RemoteRegistryStartType = (Get-Service -Name "RemoteRegistry" -ErrorAction SilentlyContinue).StartType
            $WinMgmtStatus = (Get-Service -Name "WinMgmt" -ErrorAction SilentlyContinue).Status
            $WinMgmtStartType = (Get-Service -Name "WinMgmt" -ErrorAction SilentlyContinue).StartType

            $HotfixData = (Get-WmiObject -Class Win32_QuickfixEngineering).HotFixID | Sort-Object -Descending | Select-Object -First 6
            ForEach ($HotfixItem in $HotfixData) {
                $UpdateList += "$HotfixItem `n"
            }

            # Check reboot status
            $RebootPending = "False"
            If (Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -ErrorAction SilentlyContinue) {
        
                $RebootPending = "True" 
            } ElseIf (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue) {
    
                $RebootPending = "True" 
            } ElseIf (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue) {
    
                $RebootPending = "True"
            } Else {   
    
                $RebootPending = "False"
            }
            
            # Check drive free space
            $FreeSpace = (Get-WmiObject -Class Win32_LogicalDisk | Where-Object {$_.DeviceID -eq "C:"}).FreeSpace
            $FreeC = [math]::Round($FreeSpace/1GB,2)

            # Check status of registry
            $AdaptivaServer = Invoke-command -Computer $System -ScriptBlock {(Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Adaptiva\Client\' -ErrorAction SilentlyContinue)."server_locator.server_name"}
            $AdaptivaSetup = Invoke-command -Computer $System -ScriptBlock {(Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Adaptiva\Client\' -ErrorAction SilentlyContinue)."setup.server_host_name"}
            $MPList = Invoke-command -Computer $System -ScriptBlock {(Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\CCM\' -ErrorAction SilentlyContinue)."LookUpMPList"}

            # Check PowerShell version
            $PSMajor = ($PSVersionTable).PSVersion.Major
            $PSMinor = ($PSVersionTable).PSVersion.Minor
            $PSBuild = ($PSVersionTable).PSVersion.Build
            $PSRevision = ($PSVersionTable).PSVersion.Revision       

            # Check status of WMI
            $VerifyWMI = Invoke-command -Computer $System -ScriptBlock {WinMgmt /VerifyRepository}

            # Check status of files
            $LastCache = Invoke-command -Computer $System -ScriptBlock {(Get-ChildItem -Path "C:\Windows\ccmcache" -ErrorAction SilentlyContinue).LastWriteTime | Sort-Object -Descending | Select-Object -First 1}
            $TaskLog = Invoke-command -Computer $System -ScriptBlock {(Get-ChildItem -Path "C:\Windows\CCM\Logs\CITaskMgr.log" -ErrorAction SilentlyContinue).LastWriteTime}
            $TransferLog = Invoke-command -Computer $System -ScriptBlock {(Get-ChildItem -Path "C:\Windows\CCM\Logs\DataTransferService.log" -ErrorAction SilentlyContinue).LastWriteTime}
        
        # Close PowerShell remote session
        Get-PSSession | Remove-PSSession
        
        # Display status of services
        Update-Log "[Services] `n"
        Update-Log "Adaptiva Client = $AdaptivaStatus - $AdaptivaStartType `n"
        Update-Log "BFE = $BFEStatus - $BFEStartType `n"
        Update-Log "BITS = $BITSStatus - $BITSStartType `n"
        Update-Log "CcmExec = $CcmExecStatus - $CcmExecStartType `n"
        Update-Log "LanManServer = $LanManServerStatus - $LanManServerStartType `n"
        Update-Log "LanManWorkstation = $LanManWorkstationStatus - $LanManWorkstationStartType `n"
        Update-Log "MSI Server = $MSIServerStatus - $MSIServerStartType `n"
        Update-Log "Remote Registry = $RemoteRegistryStatus - $RemoteRegistryStartType `n"

        # Display status of WMI
        Update-Log "$VerifyWMI `n`n"

        # Display status of registry
        Update-Log "[Registry] `n"
        Update-Log "Adaptiva Server = $AdaptivaServer `n"
        Update-Log "Adaptiva Setup = $AdaptivaSetup `n"
        Update-Log "MP List = $MPList `n"
        Update-Log "PowerShell Version = $PSMajor.$PSMinor.$PSBuild.$PSRevision  `n`n"

        # Display status of files
        Update-Log "[Files] `n"
        Update-Log "Last ccmcache Write = $LastCache `n"
        Update-Log "Last CITaskMgr Write = $TaskLog `n"
        Update-Log "Last DataTransferService Write = $TransferLog  `n"

        # Display status of drives
        # Update-Log "`n"
        Update-Log "Free Space = $FreeC GBs `n"

        # Display reboot requirements
        # Update-Log "`n"
        Update-Log "Reboot Pending = $RebootPending `n"

        # Display list of updates
        Update-Log "`n"
        Update-Log "[Last 5 Updates] `n"
        Update-Log "$UpdateList `n"

        Update-Log "Script Complete `n`n"
  
    } Else {
    
    # Display computer is Offline
    Update-Log "$System = Offline `n"
    Update-Log "Script Complete `n`n"
    }

}

Create-Form 


