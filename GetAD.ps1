Import-Module ActiveDirectory
$users = Get-ADUser -SearchBase "OU=Employee,OU=FacultyStaff,OU=People,DC=campus,DC=wm,DC=edu" -Filter * 
$logfile = "D:\GetAD.txt"

foreach($user in $users){
  add-content $logfile $user.Name
  $groups = Get-ADPrincipalGroupMembership $user.SamAccountName 
  foreach($group in $groups){
    add-content $logfile $group.name
  }
}
