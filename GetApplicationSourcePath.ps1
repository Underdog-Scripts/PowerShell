Write-Host "Getting all applications, please wait..."
Write-Host ""

$Applications = Get-CMApplication -Name "Premiere Elements"

Write-Host "Got all applications, count all applications, please wait..."
Write-Host ""

$AppCount = $Applications.Count

Write-Host "$AppCount applications counted."
Write-Host ""

$i = 1

ForEach($Element in $Applications)
{
    $LocDispName = $Element.LocalizedDisplayName
    
    Write-Host "Application: $LocDispName ($i/$AppCount)"
    
    $XMLList = Get-WmiObject -ComputerName "ALCRMCWSSP01" -Namespace "root\sms\site_P01" -Class SMS_ConfigurationItemBaseClass | Where {$_.LocalizedDisplayName -eq $LocDispName}

     foreach ($Element in $XMLList)
     {
         # Using the __PATH property, obtain a direct reference to the instance
         $Element = [wmi]"$($Element.__PATH)"

         foreach ($Data in $Element.SDMpackageXML)
         {
            $xml = $null
            $xml = [xml]$Data
            $AppLocation = $xml.AppMgmtDigest.DeploymentType.Installer.Contents.Content.Location
            Write-Host "Content Location: $AppLocation"

         }
    }
    Write-Host ""
    $i++
}