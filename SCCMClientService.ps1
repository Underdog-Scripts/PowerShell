<#
(Get-WMIObject -ComputerName LPF00EF85 -Namespace root\ccm -Class SMS_Client).ClientVersion
(Get-WMIObject -ComputerName LPF007EW7 -Namespace root\ccm -Class SMS_Client).ClientVersion
(Get-WMIObject -ComputerName LPF007F6Q -Namespace root\ccm -Class SMS_Client).ClientVersion
(Get-WMIObject -ComputerName LPF00E8UA -Namespace root\ccm -Class SMS_Client).ClientVersion
(Get-WMIObject -ComputerName LPF00EHAP -Namespace root\ccm -Class SMS_Client).ClientVersion
(Get-WMIObject -ComputerName VCTXW764P2186 -Namespace root\ccm -Class SMS_Client).ClientVersion
(Get-WMIObject -ComputerName WMJ0048UG -Namespace root\ccm -Class SMS_Client).ClientVersion
Write-Host "Access issues"
(Get-WMIObject -ComputerName LPC04NTZV -Namespace root\ccm -Class SMS_Client).ClientVersion
(Get-WMIObject -ComputerName LPC0PU44S -Namespace root\ccm -Class SMS_Client).ClientVersion
(Get-WMIObject -ComputerName LPC0WBLVH -Namespace root\ccm -Class SMS_Client).ClientVersion
(Get-WMIObject -ComputerName LPF00CQNW -Namespace root\ccm -Class SMS_Client).ClientVersion
(Get-WMIObject -ComputerName LPF00E88U -Namespace root\ccm -Class SMS_Client).ClientVersion
(Get-WMIObject -ComputerName LPF00EHFS -Namespace root\ccm -Class SMS_Client).ClientVersion
(Get-WMIObject -ComputerName LR9TRZYY -Namespace root\ccm -Class SMS_Client).ClientVersion
(Get-WMIObject -ComputerName WMJ004H2L -Namespace root\ccm -Class SMS_Client).ClientVersion
(Get-WMIObject -ComputerName 10.65.111.84 -Namespace root\ccm -Class SMS_Client).ClientVersion
#>

Get-Service -ComputerName LPF007EW7 -Name 'Dhcp' | Select-Object -Property Machinename, Name, Status
Get-Service -ComputerName LPF007F6Q -Name 'Dhcp' | Select-Object -Property Machinename, Name, Status
Get-Service -ComputerName LPF00E8UA -Name 'Dhcp' | Select-Object -Property Machinename, Name, Status
Get-Service -ComputerName LPF00EHAP -Name 'Dhcp' | Select-Object -Property Machinename, Name, Status
Get-Service -ComputerName VCTXW764P2186 -Name 'Dhcp' | Select-Object -Property Machinename, Name, Status
Get-Service -ComputerName WMJ0048UG -Name 'Dhcp' | Select-Object -Property Machinename, Name, Status
Get-Service -ComputerName LPC04NTZV -Name 'Dhcp' | Select-Object -Property Machinename, Name, Status
Get-Service -ComputerName LPC0PU44S -Name 'Dhcp' | Select-Object -Property Machinename, Name, Status
Get-Service -ComputerName LPF00CQNW -Name 'Dhcp' | Select-Object -Property Machinename, Name, Status
Get-Service -ComputerName LPF00E88U -Name 'Dhcp' | Select-Object -Property Machinename, Name, Status
Get-Service -ComputerName LPF00EHFS -Name 'Dhcp' | Select-Object -Property Machinename, Name, Status
Get-Service -ComputerName LR9TRZYY -Name 'Dhcp' | Select-Object -Property Machinename, Name, Status
Get-Service -ComputerName WMJ004H2L -Name 'Dhcp' | Select-Object -Property Machinename, Name, Status
Get-Service -ComputerName 10.65.111.84 -Name 'Dhcp' | Select-Object -Property Machinename, Name, Status
Get-Service -ComputerName LPC0WBLVH -Name 'Dhcp' | Select-Object -Property Machinename, Name, Status
