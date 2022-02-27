# Runs a program with different credentials

Start-Process -WindowStyle Hidden $Env:COMSPEC -workingdirectory $PSHOME -Credential $Env:USERDOMAIN\Username -ArgumentList “/c dsa.msc”

