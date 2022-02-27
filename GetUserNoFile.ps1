# Accepts partial user ID or name and searches active directory for information
param (
  [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
  [string] $CmdLine
)

Get-ADUser -Filter {GivenName -like $CmdLine -or Surname -like $CmdLine -or SamAccountName -like $CmdLine} | 
  Format-List GivenName,Surname,SamAccountName,DistinguishedName