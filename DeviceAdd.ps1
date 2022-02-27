 #Inputfile information
 $AddDevFilePref = "C:\Users\brunnegt\Documents\Scripts\PowerShell\Add-Devices"
 $UserAndComputers = Import-CSV -LiteralPath $AddDevFilePref".csv"
 $option = [System.StringSplitOptions]::RemoveEmptyEntries

 #Logfile information
 $logfile = $AddDevFilePref+"_"+(Get-Date -Format MM.dd.yyyy_hh.mm.ss)+".log"

Function AddDeviceToCollection {
Param ([string]$UserA, [string]$ColID)
#   Param ([string]$DeviceA, [string]$ColID)
   Trap { Return "error" }
      Try {
         
         $CMCMUser = Get-CMDevice -Name UserA
            
            If ($CMCMUser.ResourceID) {
               $CMCollection = Get-CMDevicerCollection -Id $ColID
                  If ($CMCollection.Name) {
                     Try {
                        Add-CMDeviceCollectionDirectMembershipRule -CollectionId $ColID -ResourceId $CMUser.ResourceID

                        Write-Host "Device $UserA was added to Collection $ColID"
                        CD C:
                        $LogText =  "Device $UserA was added to Collection $ColID"
                        log -text $LogText
                        CD $SiteCode":"
                     } Catch {
                     
                        Write-Host "Device $UserA not in Collection $ColID"
                        CD C:
                        $LogText =  "Device $UserA not in Collection $ColID"
                        Log -text $LogText
                        CD $SiteCode":"
                     }
                  } Else {
                  
                     Write-Host "Device $UserA Collection was not found: $ColID"
                     CD C:
                     $LogText =  "Device $UserA Collection was not found: $ColID"
                     log -text $LogText              
                     CD $SiteCode":"
                     #Return $False
                  }
            } Else {
               
               Write-Host "Device not found: $UserA"
               CD C:
               $LogText =  "Device not found: "
               log -text $LogText
               CD $SiteCode":"
               #Return $False
            }
      } Catch {
        
         Write-Host "ERROR: Device not found: "
         CD C:
         $LogText =  "ERROR: Device not found: "
         Log -text $LogText
         CD $SiteCode":"
         #Return $False
         
      }

}

# Customizations
         $initParams = @{}
            if((Get-Module ConfigurationManager) -eq $null) {
               # Import the ConfigurationManager.psd1 module
               Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
            }
            if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
               # Connect to the site's drive if it is not already present
               New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
            }

         # Set the current location to be the site code.
         Set-Location "$($SiteCode):\" @initParams

            Try{
               # Set the current location to be the site code.
               Set-Location "$($SiteCode):\" @initParams
               CD C:
               $LogText =  "Connected to the site: $SiteCode"
               log -text $LogText
            } Catch {
            
               CD C:
               $LogText =  "ERROR connecting to the site: $SiteCode"
               log -text $LogText
            }

         Foreach ($UCObject in $UserAndComputers) {
            
            Try {
               CD $SiteCode":"
               #AddDeviceToCollection -DeviceA $UCObject.Name -ColID $UCObject.CollectionID

               AddDeviceToCollection -UserA $UCObject.Name -ColID $UCObject.CollectionID
               

            } Catch {
               CD C:
               Write-host "Unable to find device: $UCObject.Name"
               $LogText =  "ERROR Finding device: $UCObject.Name"
               log -text $LogText
               CD $SiteCode":"
            }

         }
         
      CD C:
      $LogText =  "-----$(Get-Date) Running as - $($env:USERNAME) on $($env:COMPUTERNAME) Finished adding users to AD and updating the SCCM Schedule update-----"
      log -text $LogText 