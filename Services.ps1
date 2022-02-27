$Computers = "Learn"
$Service = "*"

ForEach ($Computer in $Computers) {
    $Computer 
    $Servicestatus = Get-Service -Name $Service -ComputerName $Computer
    
}

$ServiceStatus | Select-Object Name,Status,MachineName | Export-Csv -LiteralPath "D:\Ingalls\SccmServer\Toys\ServiceList.csv" -NoClobber -Append -NoTypeInformation