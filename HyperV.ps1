<#
.Synopsis
   This script will run Hyper-V with different credentials
.DESCRIPTION
   This script will run Hyper-V with different credentials
.EXAMPLE
   ./HyperV.ps1 UserName
.EXAMPLE
   ./HyperV.ps1 UserName
#>

Param(
	[Parameter(Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
	[string]$User
)

Start-Process -WindowStyle Hidden $Env:COMSPEC -workingdirectory $PSHOME -Credential $Env:USERDOMAIN\$User -ArgumentList “/c mmc.exe C:\Windows\System32\virtmgmt.msc"