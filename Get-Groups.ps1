Param(
	[Parameter(Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
	[string]$User
)

Import-Module ActiveDirectory

(Get-ADUser $User –Properties MemberOf | Select-Object MemberOf).MemberOf
