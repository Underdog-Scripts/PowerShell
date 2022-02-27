# Runs a program with different credentials

Start-Process -WindowStyle Hidden $Env:COMSPEC -workingdirectory $PSHOME -Credential $Env:USERDOMAIN\adm-$Env:USERNAME -ArgumentList “/c gpmc.msc”
