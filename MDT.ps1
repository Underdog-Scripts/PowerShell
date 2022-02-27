# Runs a program with different credentials

Start-Process -WindowStyle Hidden $Env:COMSPEC -workingdirectory $PSHOME -Credential $Env:USERDOMAIN\adm-$Env:USERNAME -ArgumentList “/c mmc C:\Progra~1\Micros~3\Bin\Deploy~1.msc”

