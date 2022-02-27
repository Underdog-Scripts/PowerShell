param(
#    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [System.IO.FileInfo]$Path,
 
#    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("ProductCode", "ProductVersion", "ProductName", "Manufacturer", "ProductLanguage", "FullVersion")]
    [string]$Property
)

Function GUID {
   Param ([System.IO.FileInfo]$Path, [string]$Property)
    try {
        # Read property from MSI database
        $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
        $MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($Path.FullName, 0))
        $Query = "SELECT Value FROM Property WHERE Property = '$($Property)'"
        $View = $MSIDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDatabase, ($Query))
        $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
        $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
        $Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
        
        # Commit database and close view
        $MSIDatabase.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDatabase, $null)
        $View.GetType().InvokeMember("Close", "InvokeMethod", $null, $View, $null)           
        $MSIDatabase = $null
        $View = $null
 
        # Return the value
        return $Value
    } catch {

        Write-Warning -Message $_.Exception.Message ; break
    }
}


    {
    # Run garbage collection and release ComObject
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
    [System.GC]::Collect()
}

GUID "c:\adobe\Adobe Illustrator CS5.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\Adobe Reader XI.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeALMAnchorService2-mul-x64.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeALMAnchorService2-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeAmericanEnglishSpeechAnalysisModels1All.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeAssetServices4All.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeAUM6.0All.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeBridge3All.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeCameraRaw5.0All-x64.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeCameraRaw5.0All.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeCMaps2-mul-x64.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeCMaps2-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeColorCommonSetCMYK2-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeColorCommonSetRGB2-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeColorEU_ExtraSettings2-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeColorEU_Recommended2-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeColorJA_ExtraSettings2-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeColorJA_Recommended2-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeColorNA_ExtraSettings2-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeColorNA_Recommended2-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeColorPhotoshop2-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeConnect-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeContribute5-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeCSIAll.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeCSIx64All.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeDefaultLanguage2-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeDeviceCentral2-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeDreamweaver10-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeDriveAll.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeDrivex64All.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeDynamiclinkSupport1All.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeExtendScriptToolkit3.0.0All.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeExtensionManager2All.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeFireworks10All.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeFlash10-en-ExtensionFL30.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeFlash10-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeFlash10-others-ExtensionFL30.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeFlash10-STI-en.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeFlash10-STI-other.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeFlashPlayer10_axDbg_mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeFlashPlayer10_plDbg_mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeFontsAll.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeFontsAllx64.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeIllustrator14mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeLinguisticsAll.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeLinguisticsAll_x64.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeOutputModuleAll.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobePDFL9-mul-x64.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobePDFL9-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobePDFSettings10-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobePDFSettings11-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobePDFSettings9-ja_JP.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobePDFSettings9-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobePhotoshop11-Core.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobePhotoshop11-Core_x64.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobePhotoshop11-Support.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobePhotoShopElements.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeSearchforHelp-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeServiceManager-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeSoundbooth2All.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeSoundbooth2CodecsAll.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeSuiteSharedConfiguration-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeSys_AdobeInDesign_CS6.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeSys_Captivate5_1_1_0000_EN_00.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeTypeSupport9-mul-x64.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeTypeSupport9-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeVersionCue4All.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeVideoProfilesCS2-mul.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeWebSuitePremium4-mul-trial.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeWinSoftLinguisticsPluginAll.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeWinSoftLinguisticsPluginAll_x64.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\AdobeXMPPanelsAll.msi" ProductCode | Out-File -Append ".\Adobe.txt"
Start-Sleep -s 1
GUID "c:\adobe\Adobe_DreamweaverCS5_11_5_EN_00.msi" ProductCode | Out-File -Append ".\Adobe.txt"
