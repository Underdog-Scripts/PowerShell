<#
.Synopsis
   This script will search AD for partial usernames
.DESCRIPTION
   This script will search AD for partial usernames
.EXAMPLE
   ./SearchAD zz*
.EXAMPLE
   ./SearchAD smit*
#>

Param(
	[Parameter(Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
	[string]$ADName
)

get-aduser -f {SamAccountName -like $ADName} | Format-Table SamAccountName | Out-File "C:\Temp\SearchAD.txt"

If (Get-Content C:\Temp\SearchAD.txt | %{$_ -match $ADName})
{
    echo 'Contains String'
}
else
{
    echo 'Not Contains String'
}