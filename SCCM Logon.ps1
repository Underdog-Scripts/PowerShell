[CmdletBinding()]
$SiteServer = "ALCRMCWSSP01.pmusa.net"
$SiteCode = "P01"
 
try {
    [System.Reflection.Assembly]::LoadFrom((Join-Path (Get-Item $env:SMS_ADMIN_UI_PATH).Parent.FullName "Microsoft.ConfigurationManagement.ApplicationManagement.dll")) | Out-Null
    [System.Reflection.Assembly]::LoadFrom((Join-Path (Get-Item $env:SMS_ADMIN_UI_PATH).Parent.FullName "Microsoft.ConfigurationManagement.ApplicationManagement.Extender.dll")) | Out-Null
    [System.Reflection.Assembly]::LoadFrom((Join-Path (Get-Item $env:SMS_ADMIN_UI_PATH).Parent.FullName "Microsoft.ConfigurationManagement.ApplicationManagement.MsiInstaller.dll")) | Out-Null
}
catch {
    Write-Error $_.Exception.Message
}
 
$Applications = Get-WmiObject -Namespace "root\SMS\site_$($SiteCode)" -Class "SMS_Application" -ComputerName $SiteServer -Filter 'IsLatest = "True"'
$Applications | ForEach-Object {
    $LocalizedDisplayName = $_.LocalizedDisplayName
    $CurrentApplication = [wmi]$_.__PATH
    $ApplicationXML = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($CurrentApplication.SDMPackageXML,$True)
    foreach ($DeploymentType in $ApplicationXML.DeploymentTypes) {
        $Object = New-Object -TypeName PSObject
        $RequiresLogon = {
            $GetRequiresLogon = $DeploymentType.Installer.RequiresLogOn
            $GetRequiresLogon
            if ($GetRequiresLogon -eq "True") {
                return "True"
            }
            else {
                return "False"
            }
        }
        if ($RequiresLogon.Invoke() -eq "True") {
            $Object | Add-Member -MemberType NoteProperty -Name ApplicationName -Value $LocalizedDisplayName
            $Object | Add-Member -MemberType NoteProperty -Name RequiresLogon -Value "True"
            $Object | Out-File C:\Temp\UserLogon.txt -Width 120 -Append  
        }
        else {
           # Write-Output "`nNo applications was found that requires a user to be logged on for the installation to start`n"
        }
    }
}
