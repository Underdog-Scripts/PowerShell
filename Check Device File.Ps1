<#
.Synopsis
  This script displays System Summary
.DESCRIPTION
  This script displays System Summary
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
   Modification Date: 8.30.21
   Purpose/Change: Modified for Ingalls Shipbuilding
.EXAMPLE
   ./System Summary.ps1
.EXAMPLE
   ./System Summary.ps1
#>

# Paths to files containing computer names and test results
$DeviceList = Get-Content -Path ".\Check.txt"

ForEach ($Device in $DeviceList) {
    If (!$Device) { 
        Write-Host "Entry is NULL `n"
    } Else {
            
        If (Test-Connection -ComputerName $Device -Count 1 -Quiet) {
        Write-Output $Device | Out-File -FilePath "C:\Users\zzbrunnga\Documents\Toys\ClientHealth.txt" -NoClobber -Append
        
        $ScriptOutput = ""
        $ScriptOutput = Get-Service -Name 'CcmExec' -ComputerName $Device -ErrorAction SilentlyContinue
        If ($ScriptOutput) {
            Write-Output "CcmExec ; Detected" | Out-File -FilePath "C:\Users\zzbrunnga\Documents\Toys\ClientHealth.txt" -NoClobber -Append
        } Else {
            
            Write-Output "CcmExec ; Not Detected" | Out-File -FilePath "C:\Users\zzbrunnga\Documents\Toys\ClientHealth.txt" -NoClobber -Append
        }
                
        $ScriptOutput = ""
        $ScriptOutput = Get-Service -Name 'Bits' -ComputerName $Device -ErrorAction SilentlyContinue
        If ($ScriptOutput) {
            Write-Output "Bits ; Detected" | Out-File -FilePath "C:\Users\zzbrunnga\Documents\Toys\ClientHealth.txt" -NoClobber -Append
        } Else {
            
            Write-Output "Bits ; Not Detected" | Out-File -FilePath "C:\Users\zzbrunnga\Documents\Toys\ClientHealth.txt" -NoClobber -Append
        }
                
        $ScriptOutput = ""
        $ScriptOutput = Get-Service -Name 'WuauServ' -ComputerName $Device -ErrorAction SilentlyContinue
        If ($ScriptOutput) {
            Write-Output "WuauServ ; Detected" | Out-File -FilePath "C:\Users\zzbrunnga\Documents\Toys\ClientHealth.txt" -NoClobber -Append
        } Else {
            
            Write-Output "WuauServ ; Not Detected" | Out-File -FilePath "C:\Users\zzbrunnga\Documents\Toys\ClientHealth.txt" -NoClobber -Append
        }
               
        $ScriptOutput = ""
        $ScriptOutput = Get-Service -Name 'WinMgmt' -ComputerName $Device -ErrorAction SilentlyContinue
        If ($ScriptOutput) {
            Write-Output "WinMgmt ; Detected" | Out-File -FilePath "C:\Users\zzbrunnga\Documents\Toys\ClientHealth.txt" -NoClobber -Append
        } Else {
            
            Write-Output "WinMgmt ; Not Detected" | Out-File -FilePath "C:\Users\zzbrunnga\Documents\Toys\ClientHealth.txt" -NoClobber -Append
        }

        $ScriptOutput = ""
	    $ScriptOutput = Get-WmiObject -Class Win32_BIOS -ErrorAction SilentlyContinue -ComputerName $Device |`
		Select @{Label="BIOS Version";Expression={($_.SMBIOSBIOSVersion)}},`
        Manufacturer, SerialNumber
        $ScriptOutput = $ScriptOutput -Replace '; ', "`n"
        $ScriptOutput = $ScriptOutput -Replace '=', " ; "
        $ScriptOutput = $ScriptOutput -Replace '@{', ''
        $ScriptOutput = $ScriptOutput -Replace '}', ""
        Write-Output $ScriptOutput | Out-File -FilePath "C:\Users\zzbrunnga\Documents\Toys\ClientHealth.txt" -NoClobber -Append

        $ScriptOutput = ""
		$ScriptOutput = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Device |`
        Select @{Label="Build Number"; Expression={($_.BuildNumber)}}
        $ScriptOutput = $ScriptOutput -Replace '; ', "`n"
        $ScriptOutput = $ScriptOutput -Replace '=', " ; "
        $ScriptOutput = $ScriptOutput -Replace '@{', ''
        $ScriptOutput = $ScriptOutput -Replace '}', ""
        Write-Output $ScriptOutput | Out-File -FilePath "C:\Users\zzbrunnga\Documents\Toys\ClientHealth.txt" -NoClobber -Append
        
        $ScriptOutput = ""
		$ScriptOutput = Get-WmiObject -Class Win32_LogicalDisk -ErrorAction SilentlyContinue -ComputerName $Device |`
        Select @{ Label = "Drive"; Expression = { ($_.DeviceId) } },`
        @{Label="Size(GB)"; Expression={ [math]::Round($_.Size/1GB, 2)}},`
        @{Label="Free Space(GB)"; Expression={[math]::Round($_.FreeSpace/1GB, 2)}}
        $ScriptOutput = $ScriptOutput -Replace '; ', "`n"
        $ScriptOutput = $ScriptOutput -Replace '=', " ; "
        $ScriptOutput = $ScriptOutput -Replace '@{', ''
        $ScriptOutput = $ScriptOutput -Replace '}', ""
        Write-Output $ScriptOutput | Out-File -FilePath "C:\Users\zzbrunnga\Documents\Toys\ClientHealth.txt" -NoClobber -Append

        $ScriptOutput = ""
    	$ScriptOutput = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Device |`
        Select @{Label="Last BootUp Time"; Expression={$_.ConverttoDateTime($_.lastbootuptime)}}
        $ScriptOutput = $ScriptOutput -Replace '; ', "`n"
        $ScriptOutput = $ScriptOutput -Replace '=', " ; "
        $ScriptOutput = $ScriptOutput -Replace '@{', ''
        $ScriptOutput = $ScriptOutput -Replace '}', "`n`n"
        Write-Output $ScriptOutput | Out-File -FilePath "C:\Users\zzbrunnga\Documents\Toys\ClientHealth.txt" -NoClobber -Append

        } Else {

            Write-Output "$Device ; Offline `n`n" | Out-File -FilePath "C:\Users\zzbrunnga\Documents\Toys\ClientHealth.txt" -NoClobber -Append
        }

    }

}

