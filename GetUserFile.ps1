# Accepts partial user ID or name and searches active directory for information
param (
  [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
  [string] $CmdLine
)
$OutputFile = $env:Temp + '\GetUser.txt'

 # Verify and delete previous report file
 if (Test-Path $OutputFile) {
    Remove-Item $OutputFile
}

Get-ADUser -Filter {GivenName -like $CmdLine -or Surname -like $CmdLine -or SamAccountName -like $CmdLine} | 
    Format-List GivenName,Surname,SamAccountName,DistinguishedName |
    Out-File $OutputFile
