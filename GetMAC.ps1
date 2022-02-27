$IP = "10.194.9.157"
$Client = "NG00244191"

arp -a $IP

(gwmi -Class Win32_NetworkAdapterConfiguration | where { $_.IpAddress -eq $IP }).MACAddress

# Solution 1
Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" -ComputerName $Client Select-Object -Property MACAddress, Description

# Solution 2
Get-WmiObject -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" -ComputerName $Client Select-Object -Property MACAddress, Description

# Solution 3
getmac.exe /s $Client
getmac /s $IP

nbtstat -a $Client
nbtstat -A $IP
nbtstat -n 
