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

GUID "C:\Adobe\AcrobatProfessional\v10\AcroPro.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\AcrobatProfessional\v2015\Adobe Acrobat\AcroPro.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\AcrobatReader\v11.0.06\AcroRead.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\Adobe Illustrator CC 2018\22.1\Completed\Build\Illustrator.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\Adobe Illustrator CC 2018\22.1\Vendor Source\Build\Illustrator.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\Adobe Photoshop CC 2018\19.1.5\Completed\Build\Photoshop.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\Adobe Photoshop CC 2018\19.1.5\Vendor Source\Build\Photoshop.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\AuthorwareWebPlayer\v1.0\AuthorWare.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\Captivate\v5.0\Build\AdobeSys_Captivate5_1_1_0000_EN_00.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\WinBootstrapper.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeAcrobat9-es_ES\AcroPro.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeAIR1.0\setup.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeAIR1.0\Adobe AIR\Versions\1.0\Resources\template.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeALMAnchorService2-mul\AdobeALMAnchorService2-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeALMAnchorService2-mul-x64\AdobeALMAnchorService2-mul-x64.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeAmericanEnglishSpeechAnalysisModels1All\AdobeAmericanEnglishSpeechAnalysisModels1All.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeAMP-mul\setup.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeAMP-mul\Adobe AIR\Versions\1.0\Resources\template.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeAssetServices4All\AdobeAssetServices4All.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeAUM6.0All\AdobeAUM6.0All.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeBridge3All\AdobeBridge3All.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeCameraRaw5.0All\AdobeCameraRaw5.0All.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeCameraRaw5.0All-x64\AdobeCameraRaw5.0All-x64.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeCMaps2-mul\AdobeCMaps2-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeCMaps2-mul-x64\AdobeCMaps2-mul-x64.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeColorCommonSetCMYK2-mul\AdobeColorCommonSetCMYK2-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeColorCommonSetRGB2-mul\AdobeColorCommonSetRGB2-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeColorEU_ExtraSettings2-mul\AdobeColorEU_ExtraSettings2-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeColorEU_Recommended2-mul\AdobeColorEU_Recommended2-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeColorJA_ExtraSettings2-mul\AdobeColorJA_ExtraSettings2-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeColorJA_Recommended2-mul\AdobeColorJA_Recommended2-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeColorNA_ExtraSettings2-mul\AdobeColorNA_ExtraSettings2-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeColorNA_Recommended2-mul\AdobeColorNA_Recommended2-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeColorPhotoshop2-mul\AdobeColorPhotoshop2-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeConnect-mul\AdobeConnect-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeContribute5-mul\AdobeContribute5-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeCSIAll\AdobeCSIAll.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeCSIx64All\AdobeCSIx64All.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeDefaultLanguage2-mul\AdobeDefaultLanguage2-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeDeviceCentral2-mul\AdobeDeviceCentral2-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeDreamweaver10-mul\AdobeDreamweaver10-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeDriveAll\AdobeDriveAll.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeDrivex64All\AdobeDrivex64All.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeDynamicLinkSupport1All\AdobeDynamiclinkSupport1All.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeExtendScriptToolKit3.0.0All\AdobeExtendScriptToolkit3.0.0All.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeExtensionManager2All\AdobeExtensionManager2All.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeFireworks10All\AdobeFireworks10All.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeFlash10-en-ExtensionFL30\AdobeFlash10-en-ExtensionFL30.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeFlash10-mul\AdobeFlash10-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeFlash10-others-ExtensionFL30\AdobeFlash10-others-ExtensionFL30.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeFlash10-STI-es\AdobeFlash10-STI-es.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeFlash10-STI-other\AdobeFlash10-STI-other.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeFlashPlayer10_axDbg_mul\AdobeFlashPlayer10_axDbg_mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeFlashPlayer10_plDbg_mul\AdobeFlashPlayer10_plDbg_mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeFontsAll\AdobeFontsAll.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeFontsAllx64\AdobeFontsAllx64.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeIllustrator14mul\AdobeIllustrator14mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeLinguisticsAll\AdobeLinguisticsAll.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeLinguisticsAll_x64\AdobeLinguisticsAll_x64.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeOutputModuleAll\AdobeOutputModuleAll.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobePDFL9-mul\AdobePDFL9-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobePDFL9-mul-x64\AdobePDFL9-mul-x64.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobePDFSettings9-ja_JP\AdobePDFSettings9-ja_JP.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobePDFSettings9-mul\AdobePDFSettings9-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobePhotoshop11-Core\AdobePhotoshop11-Core.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobePhotoshop11-Core_x64\AdobePhotoshop11-Core_x64.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobePhotoshop11-Support\AdobePhotoshop11-Support.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeSearchforHelp-mul\AdobeSearchforHelp-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeServiceManager-mul\AdobeServiceManager-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeSoundbooth2All\AdobeSoundbooth2All.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeSoundbooth2CodecsAll\AdobeSoundbooth2CodecsAll.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeSuiteSharedConfiguration-mul\AdobeSuiteSharedConfiguration-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeTypeSupport9-mul\AdobeTypeSupport9-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeTypeSupport9-mul-x64\AdobeTypeSupport9-mul-x64.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeVersionCue4All\AdobeVersionCue4All.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeVideoProfilesCS2-mul\AdobeVideoProfilesCS2-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeWebSuitePremium4-mul-trial\AdobeWebSuitePremium4-mul-trial.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeWinSoftLinguisticsPluginAll\AdobeWinSoftLinguisticsPluginAll.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeWinSoftLinguisticsPluginAll_x64\AdobeWinSoftLinguisticsPluginAll_x64.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AdobeXMPPanelsAll\AdobeXMPPanelsAll.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\aifsdk-win\aifsdk-win.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AMECore1All\AMECore1All.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\AMEImporter1All\AMEImporter1All.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\CreativeSuiteWebPremium\CS4\Media\payloads\kuler2.0-mul\kuler2.0-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\Dreamweaver\CS5\Build\Adobe_DreamweaverCS5_11_5_EN_00.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v20.0.0.270\activex\install_flash_player_20_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v20.0.0.270\plugin\install_flash_player_20_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v20.0.0.286\activex\install_flash_player_20_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v20.0.0.286\plugin\install_flash_player_20_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v20.0.0.306\activex\install_flash_player_20_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v20.0.0.306\plugin\install_flash_player_20_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v21.0.0.197\activex\install_flash_player_21_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v21.0.0.197\plugin\install_flash_player_21_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v21.0.0.213\Flash\Completed\install_flash_player_21_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v21.0.0.213\Flash\Completed\install_flash_player_21_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v22.0.0.192\Completed\install_flash_player_22_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v22.0.0.192\Completed\install_flash_player_22_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v22.0.0.210\Completed\install_flash_player_22_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v22.0.0.210\Completed\install_flash_player_22_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v23.0.0.162\Completed\install_flash_player_23_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v23.0.0.162\Completed\install_flash_player_23_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v23.0.0.205\Completed\install_flash_player_23_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v23.0.0.205\Completed\install_flash_player_23_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v23.0.0.207\Completed\install_flash_player_23_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v23.0.0.207\Completed\install_flash_player_23_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v24.0.0.186\Completed\install_flash_player_24_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v24.0.0.186\Completed\install_flash_player_24_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v24.0.0.194\Completed\install_flash_player_24_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v24.0.0.194\Completed\install_flash_player_24_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v24.0.0.221\Completed\install_flash_player_24_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v24.0.0.221\Completed\install_flash_player_24_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v25.0.0.127\Completed\install_flash_player_25_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v25.0.0.127\Completed\install_flash_player_25_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v25.0.0.148\Completed\install_flash_player_25_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v25.0.0.148\Completed\install_flash_player_25_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v25.0.0.171\Completed\install_flash_player_25_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v25.0.0.171\Completed\install_flash_player_25_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v26.0.0.131\Completed\install_flash_player_26_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v26.0.0.131\Completed\install_flash_player_26_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v26.0.0.137\Completed\install_flash_player_26_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v26.0.0.137\Completed\install_flash_player_26_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v26.0.0.151\Completed\install_flash_player_26_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v26.0.0.151\Completed\install_flash_player_26_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v27.0.0.130\Completed\install_flash_player_27_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v27.0.0.130\Completed\install_flash_player_27_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v27.0.0.170\Completed\install_flash_player_27_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v27.0.0.170\Completed\install_flash_player_27_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v27.0.0.183\Completed\install_flash_player_27_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v27.0.0.183\Completed\install_flash_player_27_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v27.0.0.187\Completed\install_flash_player_27_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v27.0.0.187\Completed\install_flash_player_27_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v28.0.0.126\Completed\install_flash_player_28_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v28.0.0.126\Completed\install_flash_player_28_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v28.0.0.161\Completed\install_flash_player_28_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v28.0.0.161\Completed\install_flash_player_28_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v29.0.0.171\Completed\install_flash_player_29_active_x.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\FlashPlayer\v29.0.0.171\Completed\install_flash_player_29_plugin.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\Illustrator\CS5\Build\Adobe Illustrator CS5.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\Illustrator\CS5\Exceptions\AdobePDFSettings10-mul\AdobePDFSettings10-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\InDesign\CS6\Build\AdobeSys_AdobeInDesign_CS6.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\InDesign\CS6\Exceptions\AdobePDFSettings11-mul\AdobePDFSettings11-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\LiveCycle\11.0\Completed\Designer.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\Photoshop\CS6\Build\Photoshop_CS6.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\Photoshop\CS6\Exceptions\AdobePDFSettings11-mul\AdobePDFSettings11-mul.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\Photoshop\Elements\AdobePhotoShopElements.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\Reader DC\Completed\AcroRead.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\Robohelp\2015\x86\Completed\Robohelp 2015.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\Robohelp\2017\Completed\Robohelp 2017.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\Robohelp\2019\Completed\Source\Adobe RoboHelp 2019.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\Robohelp 2015\x64\Robohelp 2015.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\Robohelp 2015\x86\Robohelp 2015.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\ShockwavePlayer\v12.1.9.160\sw_lic_full_installer.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\ShockwavePlayer\v12.2.0.162\sw_lic_full_installer.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\ShockwavePlayer\v12.2.3.183\sw_lic_full_installer.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
GUID "C:\Adobe\ShockwavePlayer\v12.2.4.194\sw_lic_full_installer.msi" ProductCode | Out-File -Append ".\Adobe2.txt"
