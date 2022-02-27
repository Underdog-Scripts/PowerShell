$Dashes = '------------------------------------------------------------------------------'
$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-Location -Path "$($SiteCode.Name):\"

ForEach ($TSName in Get-Content "C:\Temp\TSList.txt")
   {
   $TS = Get-CMTaskSequence -Name $TSName
   '{Name}', $TS.Name, '{Description}', $TS.Description, '{DependentProgram}', $TS.DependentProgram, '{ActionInProgress}', $TS.ActionInProgress | Out-File -filepath C:\Temp\TSDetails.txt -Append
   $TS.References | select Package, Program | Out-File -filepath C:\Temp\TSDetails.txt -Append -Width 200
   $Dashes | Out-File -filepath C:\Temp\TSDetails.txt -Append -Width 200
   }

