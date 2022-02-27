$Computers = Get-Content ".\All1.txt"
$Service = "*"

ForEach ($Computer in $Computers) {
    $Computer 
    $Servicestatus = Get-Service -Name $Service -ComputerName $Computer
    $ServiceStatus | Select-Object Name,Status,MachineName | Export-Csv -LiteralPath "C:\Users\zzbrunnga\Documents\ServiceList.csv" -NoClobber -Append -NoTypeInformation
    
}

