<#
.Synopsis
Retrieve the security descriptor outputs to format list
.Description
Retrieve the security descriptor of a file, folder or registry key outputs to format list
.Parameter CmdLine
Name of file, folder or registry key
.Example
Get security descriptor for C: drive
GetSecDesc C:\
.Example
Get security descriptor for C: drive
GetSecDesc HKLM:\Software\Microsoft\Windows
#>
Param(
  [Parameter(Mandatory=$True)]
  [string]$CmdLine
)

Get-Acl $CmdLine | Format-List
    

