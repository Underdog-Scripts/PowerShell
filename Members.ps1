Import-Module ActiveDirectory
Get-ADPrincipalGroupMembership "cov/kck39993" | Select-Object -Property Name, GroupScope, GroupCategory | Sort-Object -Property Name | FT -A