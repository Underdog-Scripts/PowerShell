$DeviceCollection = "AppFac-SW Citrix Receiver 14.3.0.5014 ENT Computers"
Write-Host $DeviceCollection
Get-CMCollection -Name $DeviceCollection | Get-CMDeviceCollectionDirectMembershipRule | Select RuleName
