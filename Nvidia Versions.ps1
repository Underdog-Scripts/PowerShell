$Computers = Get-Content ".\AllLast.txt"

ForEach ($Computer in $Computers) {
    $Version = Get-WmiObject -Class Win32_VideoController -ComputerName $Computer -ErrorAction SilentlyContinue | Where-Object Name -like "*NVIDIA*" | Select Name, DriverVersion
    $Version = $Version -Replace "@{Name=", "$Computer; "
    $Version = $Version -Replace "DriverVersion=", ""
    $Version = $Version -Replace "}", ""
    $Version | Out-File ".\Nvidia Versions.txt" -NoClobber -Append
}