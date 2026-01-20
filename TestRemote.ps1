Test-NetConnection -ComputerName RemotePCName -Port 5985

Test-WSMan -ComputerName RemotePCName

$sess = New-PSSession -ComputerName RemotePCName -Credential (Get-Credential)
Invoke-Command -Session $sess -ScriptBlock { Get-Date }
Remove-PSSession $sess
