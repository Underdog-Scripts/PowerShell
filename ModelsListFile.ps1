$DeviceList = Get-Content -Path ".\WhatModel.txt"

ForEach ($Device in $DeviceList) {
    If (!$Device) { 
        Write-Host "Entry is NULL `n"
    } Else {

        Invoke-Sqlcmd -QueryTimeout 500 -ConnectionTimeout 500 -Query "use CM_LW2
        select V_R_SYSTEM.Name0, V_GS_COMPUTER_SYSTEM.Model0 
        from V_R_System inner join V_GS_COMPUTER_SYSTEM on V_GS_COMPUTER_SYSTEM.ResourceID = V_R_System.ResourceId where V_R_SYSTEM.Name0 Like '%$Device%'" | Export-Csv -LiteralPath ".\WhatModel.csv" -NoClobber -Append -NoTypeInformation
    }
}