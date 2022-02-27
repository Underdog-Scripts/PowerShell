 #Inputfile information
 $AddDevFilePref = "C:\Users\username\Documents\Scripts\PowerShell\Add-Devices"
 $UserAndComputers = Import-CSV -LiteralPath $AddDevFilePref".csv"
 $option = [System.StringSplitOptions]::RemoveEmptyEntries

 #Logfile information
 $logfile = $AddDevFilePref+"_"+(Get-Date -Format MM.dd.yyyy_hh.mm.ss)+".log"

# Customizations
$initParams = @{}

# Do not change anything below this line

# Import the ConfigurationManager.psd1 module 
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}

# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams

Try{
    # Set the current location to be the site code.
    Set-Location "$($SiteCode):\" @initParams
   
    

} Catch {
    
    
}

Function AddDeviceToCollection
{
    Param ([string]$UserA, [string]$ColID)
    CD $SiteCode":"

    Trap { Return "error" }
        
    Try {
        $CMUser = Get-CMDevice -Name $UserA
        If ($CMUser.ResourceID)
        {
            $CMCollection = Get-CMDeviceCollection -Id $ColID
            
            If ($CMCollection.Name)
            {
                Try
                {
                    
                    Add-CMDeviceCollectionDirectMembershipRule -CollectionId $ColID -ResourceId $CMUser.ResourceID
                    Write-Host "Device $UserA was added collection $ColID"
                    
                    CD $SiteCode":"
                } Catch
                {
                    Write-Host "Device $UserA already exists in Collecion $ColID"
                   
                    CD $SiteCode":"
                }
                                
            } Else
            {
                Write-Host "Device Collecion was not found: $ColID"
                        
                CD $SiteCode":"
                #Return $False
            }
             

        } Else
        {
            Write-Host "Device not found: $UserA"
           
            CD $SiteCode":"
            #Return $False
        }
    

    } Catch {
        Write-Host "ERROR: Device not found: $UserA"
        
        CD $SiteCode":"
        #Return $False
    }


}

Foreach ($UCObject in $UserAndComputers)
{

    Try {
       AddDeviceToCollection -UserA $UCObject.Name -ColID $UCObject.CollectionID
       Param ([string]$UserA, [string]$ColID)

    } Catch {
       
       CD $SiteCode":"
    }
                     
}

