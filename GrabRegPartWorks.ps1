$System = "GTS"
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate"
$ValueName = "LastTaskOperationHandle"

$PSession = New-PSSession -ComputerName $System
Invoke-Command -Session $PSession -ArgumentList $RegPath, $ValueName -ScriptBlock {
    param($RegPath, $ValueName)

    $GrabReg = (Get-ItemProperty $RegPath).$ValueName
    $GrabReg
}

Get-PSSession | Remove-PSSession