Get-ChildItem -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall -Recurse -ErrorAction SilentlyContinue | 
% {
If((Get-ItemProperty -Path $_.PsPath) -Match "Gary")
    {
        $_.PsPath
    }
}

Get-ChildItem -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall -Recurse -ErrorAction SilentlyContinue | 
% {
If((Get-ItemProperty -Path $_.PsPath) -Match "Gary")
    {
        $_.PsPath
    }
}