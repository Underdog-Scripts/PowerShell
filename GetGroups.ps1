Param(
	[Parameter(Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
	[string]$User
)

Import-Module ActiveDirectory

Get-WmiObject -Class Win32_MappedLogicalDisk | select Name, ProviderName
(Get-ADUser $User –Properties MemberOf | Select-Object MemberOf).MemberOf
