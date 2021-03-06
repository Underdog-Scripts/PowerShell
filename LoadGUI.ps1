[CmdletBinding()]

Param(
    [Parameter(Mandatory=$True,Position=1)]
    [string]$XamlPath
)

$Global:xmlWPF = Get-Content -Path $XamlPath

#[XML]$Global:xmlWPF = (Get-Content -Path $XamlPath)
#[XML]$Global:xmlWPF = (Get-Content -Path $XamlPath) -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'


#Add WPF and Windows Forms assemblies
try{
    Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,system.windows.forms
} catch {
    Throw "Failed to load Windows Presentation Framework assemblies."
}


#Create the XAML reader using a new XML node reader
$Global:xamGUI = [Windows.Markup.XamlReader]::Load((new-object System.Xml.XmlNodeReader $xmlWPF))


#Create hooks to each named object in the XAML
$xmlWPF.SelectNodes("//*[@Name]") | %{
    Set-Variable -Name ($_.Name) -Value $xamGUI.FindName($_.Name) -Scope Global
}

