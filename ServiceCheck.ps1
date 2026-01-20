<#
.Synopsis
  This script performs a Service Check
.DESCRIPTION
  This script performs a Service Check
.PARAMETER List 
   None
.INPUTS
   List of computers
.OUTPUTS
   Text file ServicesCheck.txt
.NOTES
   Version:        1.0
   Author:         GB
   Creation Date:  5.22.23
   Purpose/Change: Initial script development
   Modification Date: 5.23.23, 6.5.23, 6.7.23
   Purpose/Change: Modified for USPS
.EXAMPLE
   ./ServiceCheck.ps1
.EXAMPLE
   ./ServiceCheck.ps1
#>

# Read list of computers from text file
$SystemList = Get-Content -Path ".\SystemList.txt"

# Add columns names to file
Write-Output "Service Name;Status;StartUp Type" | Out-File -FilePath ".\ServicesCheck.txt" -NoClobber -Append 

# Check if computer is online, then start gathering information
ForEach ($System in $SystemList) {
    If (Test-Connection -ComputerName $System -Count 1 -Quiet) {

    # Send information to text file
    Write-Output "Computer Name;$System" | Out-File -FilePath ".\ServicesCheck.txt" -NoClobber -Append

    # List of services to check
    $ServiceArray = "BFE","BITS","CcmExec","LanManServer","LanManWorkstation","MSIServer","RemoteRegistry","WinMgmt"
        ForEach ($ServiceName in $ServiceArray) {
            $ServiceNameStatus = (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue -ComputerName $System).Status
            $ServiceNameStartType = (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue -ComputerName $System).StartType
            Write-Output "$ServiceName;$ServiceNameStatus;$ServiceNameStartType" | Out-File -FilePath ".\ServicesCheck.txt" -NoClobber -Append
        }

        # Gather PowerShell version information
        $PSMajor = ($PSVersionTable).PSVersion.Major
        $PSMinor = ($PSVersionTable).PSVersion.Minor
        $PSBuild = ($PSVersionTable).PSVersion.Build
        $PSRevision = ($PSVersionTable).PSVersion.Revision
        Write-Output "PowerShell Version;$PSMajor.$PSMinor.$PSBuild.$PSRevision" | Out-File -FilePath ".\ServicesCheck.txt" -NoClobber -Append
        
        # Check status of WMI
        $VerifyWMI = WinMgmt /VerifyRepository
        $VerifyWMI = $VerifyWMI -replace "WMI repository ", ""
        Write-Output "WMI repository;$VerifyWMI" | Out-File -FilePath ".\ServicesCheck.txt" -NoClobber -Append
        Write-Output "" | Out-File -FilePath ".\ServicesCheck.txt" -NoClobber -Append
    } Else {
       
        Write-Output "$System;Offline" | Out-File -FilePath ".\ServicesCheck.txt" -NoClobber -Append
        Write-Output "Completed" | Out-File -FilePath ".\ServicesCheck.txt" -NoClobber -Append
        Write-Output "" | Out-File -FilePath ".\ServicesCheck.txt" -NoClobber -Append
    }
 
}
