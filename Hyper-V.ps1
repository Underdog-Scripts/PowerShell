# Runs a program with different credentials

Start-Process -WindowStyle Hidden $Env:COMSPEC -workingdirectory $PSHOME -Credential $Env:USERDOMAIN\zzgary.brunner -ArgumentList “/c mmc.exe C:\Windows\System32\virtmgmt.msc"