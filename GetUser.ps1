<#
.Synopsis
Retrieves the first name, fast name userID and distinguished name
.Description
Retrieves the first name, fast name userID and distinguished name from partial input of each
.Parameter CmdLine
Complete or partial first name, last name or UserID
.Example
GetUser Manni*
.Example
GetUser clmann*
#>

# Prompt for SamAccountName or network ID
param (
  [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
  [string] $CmdLine
)

# Search AD for SamAccountName and return Given Name, Surname etc
Get-ADUser -Filter {GivenName -like $CmdLine -or Surname -like $CmdLine -or SamAccountName -like $CmdLine} | 
  Format-List GivenName,Surname,SamAccountName,DistinguishedName