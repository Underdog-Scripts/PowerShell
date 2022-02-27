# $DeviceList = Get-Content -Path "C:\Users\zzbrunnga\Documents\Win7.txt"

# ForEach ($Device in $DeviceList) {
#    $Device
    Get-WmiObject -Class Win32_OperatingSystem -ErrorAction SilentlyContinue -ComputerName "ID7346-053" | Select * | Out-File "C:\Users\zzbrunnga\Documents\SystemStatus.txt" -NoClobber -Append
    # Get-WmiObject -Class Win32_OperatingSystem -ErrorAction SilentlyContinue -ComputerName $Device | Select Csname, Caption, Version, BuildNumber, @{Label="LastBootUpTime";Expression={$_.ConverttoDateTime($_.lastbootuptime)}}, FreePhysicalMemory, RegisteredUser | Out-File "./SystemStatus.txt" -NoClobber -Append
    # Get-WmiObject -Class Win32_BIOS -ErrorAction SilentlyContinue -ComputerName $Device | Select SMBIOSBIOSVersion, SMBIOSMajorVersion, SMBIOSMinorVersion, BIOSVersion, Manufacturer, SerialNumber | Out-File "./$Device.txt" -NoClobber -Append
    # Get-WmiObject -Class Win32_ComputerSystem -ErrorAction SilentlyContinue -ComputerName $Device | Select Name, Caption, ChassisSKUNumber, Manufacturer, Model, OEMStringArray, PrimaryOwnerName, TotalPhysicalMemory | Out-File "./$Device.txt" -NoClobber -Append
    # Get-WmiObject -Class Win32_DiskDrive -ErrorAction SilentlyContinue -ComputerName $Device | Select Caption, Size | Out-File "./$Device.txt" -NoClobber -Append
    # Get-WmiObject -Class Win32_VideoController -ErrorAction SilentlyContinue -ComputerName $Device | Select PSComputerName, VideoModeDescription, Caption, Description, Name, VideoProcessor | Out-File "./$Device.txt" -NoClobber -Append
# }

