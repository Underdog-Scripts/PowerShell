<#
.Synopsis
    This script reads a text file and searches the registry
.DESCRIPTION
    This script reads a text file and searches the registry
.PARAMETER List 
    Computer IP address or computer name
.INPUTS
    Computer IP address or computer name
.OUTPUTS
    Output file of each action result
.NOTES
    Version:        1.0
    Author:         GTB
    Creation Date:  4.1.25
    Purpose/Change: Initial script development
    Update Date:  4.1.25
    Purpose/Change: Modified to use files for input and output
.EXAMPLE
    ./Reg Check File.ps1
.EXAMPLE
    ./Reg Check File.ps1
#>

# Paths to files containing computer names and test results
$ComputerList = ".\RegCheckFile.txt"
$ResultsList = ".\RegCheckFileOut.txt"

# Read files with computer names
$DeviceList = Get-Content -Path $ComputerList
$PackageName = Read-Host "Please enter package name"

ForEach ($SystemName in $DeviceList) {
    If (Test-Connection -ComputerName $SystemName -Count 1 -Quiet) {
        $Session = New-PSSession -ComputerName $SystemName
        $ScriptBlock = {
        Get-ItemProperty HKLM:\SOFTWARE\WOW6432Node\USPS\RSS_$PackageName | Select-Object DelayReason, FastRetries, PkgDesc, PkgName, PkgStatus, RollBack}
        $Results = Invoke-Command -Session $Session -ScriptBlock $ScriptBlock
        $Results | Out-File -Append $ResultsList
    } Else {
    
    "$SystemName = Offline" | Out-File -Append $ResultsList
    }
    
}



