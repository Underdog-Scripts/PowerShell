
$RemoteComputers = Get-Content D:\Downloads\CompAdm.txt
$Results = "NULL"

If (Test-Connection -ComputerName $RemoteComputers -Quiet){
    $Results = Invoke-Command -ComputerName $RemoteComputers -ScriptBlock {Get-LocalGroupMember -Group "Administrators"}
    write-host $Results
    If ($Results -match 'adminuser'){
                
        $RemoteComputers | Out-File D:\Downloads\Admins.txt
    }

}


#If (Test-Connection -ComputerName $RemoteComputers -Quiet){
#    Invoke-Command -ComputerName $RemoteComputers -ScriptBlock {Remove-LocalGroupMember -Group "Administrators" -Member "campus\adminuser"}
#}



