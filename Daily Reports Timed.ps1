<#           
.NOTES
===========================================================================
Created with:    SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.173
Created on:      3/17/2021
Modified on:     4/8/2021, 2/21/2022
Created by:      Gary Brunner
Organization:    Gary Brunner
Filename:        Daily Reports Timed.ps1 
===========================================================================
.DESCRIPTION
Schedules the time a program starts
#>

#__________________________________________________________________________________________________

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Log file location and name
$ScriptName = "Daily Reports Timed"
$LogFile = ".\$ScriptName.txt"
Start-Transcript -Path $Logfile

$TimeIs = Get-Date -UFormat "%Y%m%d%H%M%S"
Write-Output "TimeIs = $TimeIs"

Function GetTime([String]$TimeTo) {

    $TimeIs = Get-Date -UFormat "%m%d%H%M"
    Write-Host "TimeIs = $TimeIs"
    Write-Host "TimeTo = $TimeTo"
    Write-Host ""

        While($TimeIs -le $TimeTo){
		    Start-Sleep -Seconds 1800
            $TimeIs = Get-Date -UFormat "%m%d%H%M"
            Write-Host "TimeIs = $TimeIs"
            Write-Host "TimeTo = $TimeTo"
        }

    $TimeIs = Get-Date -UFormat "%m%d%H%M"
    Write-Host "TimeIsDo = $TimeIs"
    Write-Host "TimeToDo = $TimeTo"
    Write-Host "Start-Process -FilePath .\Daily Reports.exe"
    Start-Process -FilePath ".\Daily Reports.exe"

}

Function Create-Form {

    Write-Host "Create Form"
    $TimeIs = Get-Date -UFormat "%m%d%H%M"
    Write-Host "TimeIs = $TimeIs"

    # Create Form
    $Form = New-Object System.Windows.Forms.Form
    $Form.ShowIcon=$False
    $Form.Font = 'Microsoft Sans Serif,11' 
    $Form.Text = "Schedule Program"
    $Form.Size = New-Object System.Drawing.Size(380, 146) 
    $Form.StartPosition = "CenterScreen"

    # Input Box Label
    $TimeTo = New-Object System.Windows.Forms.Label 
    $TimeTo.Location = New-Object System.Drawing.Size(8, 8) 
    $TimeTo.Size = New-Object System.Drawing.Size(424, 26)
    $TimeTo.Text = "$TimeIs Press Enter"
    $TimeTo.Font = 'Microsoft Sans Serif,11'
    $Form.Controls.Add($TimeTo)
    
    # Input Box
    $TimeToTextBox = New-Object System.Windows.Forms.TextBox 
    $TimeToTextBox.Location = New-Object System.Drawing.Size(10, 48) 
    $TimeToTextBox.Size = New-Object System.Drawing.Size(200, 28)
    $TimeToTextBox.Font = 'Microsoft Sans Serif,11'
    $Form.Controls.Add($TimeToTextBox)
    $TimeToTextBox.Add_KeyDown({
    If ($_.KeyCode -eq "Enter"){
        Script:GetTime $TimeToTextBox.Text
        $Form.Close()
    }})
    
    # Close Button
    $CloseButton = New-Object System.Windows.Forms.Button
    $CloseButton.Location = New-Object System.Drawing.Point(220, 44)
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

Create-Form

#__________________________________________________________________________________________________

# Add file ending to log file
$TimeIs = Get-Date -UFormat "%Y%m%d%H%M%S"
Write-Output "$TimeIs : $ScriptName complete"

Stop-Transcript