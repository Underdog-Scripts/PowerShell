$DeviceList = Get-Content -Path ".\All.txt"

ForEach ($Device in $DeviceList) {
    Get-WmiObject -Class Win32_VideoController -ErrorAction SilentlyContinue -ComputerName $Device | Select SystemName, DriverVersion | Out-File ".\Video Versions.txt" -NoClobber -Append
}

