# Read text file of computer name and write report file with BIOS version

Write-Host "HP Pilot Systems"
$InputList = ".\HP Pilot Systems.txt"

# Read names from input file
$DeviceList = Get-Content -Path $InputList

# Gathers details about each name and writes to report file
ForEach ($Device in $DeviceList) {
    $Model = "NULL"
    $BIOSVer = "NULL"
    If (Test-Connection -ComputerName $Device -Count 1 -Quiet) {
        $Model = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction SilentlyContinue -ComputerName $Device | Select Model
        $Model = $Model -Replace '@{Model=', ''
        $Model = $Model -Replace '}', ''
        $BIOSVer = Get-WmiObject -Class Win32_BIOS -ErrorAction SilentlyContinue -ComputerName $Device | Select Name
        $BIOSVer = $BIOSVer -Replace '@{Name=', ''
        $BIOSVer = $BIOSVer -Replace '}', ''
        
		Write-Host "$Device ; $Model"

		        
    } Else {

        Write-Host "$Device ; $Model ; Offline"
		
    }

}

