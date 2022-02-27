#Configuration Manager Variables
$SiteCode = "P01" #ConfigMgr 3 character site code
$SiteServer = "ALCRMCWSSP01.pmusa.net" #ConfigMgr site server

$UserAndComputers = Import-CSV -LiteralPath "\\pmusa.net\dslpmu\Projects\Windows10\Scripts\Tools\SCCM_Scripts\BulkCollectionUpdates\Add-Users-Computers.csv"

$option = [System.StringSplitOptions]::RemoveEmptyEntries

#Logfile information
$logfile = ("\\pmusa.net\dslpmu\Projects\Windows10\Scripts\Tools\SCCM_Scripts\BulkCollectionUpdates\Logs\SCCM-User&ComputerCollectionAdds.log") 





Function Log
{
    param 
    (
        [Switch]$fout,
        [String]$text
        
    )
    
    ac $logfile $text
    if($showConsoleOutput)
    {
        if($fout)
        {
            Write-Host $text -ForegroundColor Red
        }
        else
        {
            Write-Host $text -ForegroundColor Green
        }
    }
}


Function AddUserToCollection
{
    Param ([string]$UserA, [string]$ColID)

    Trap { Return "error" }
        
    Try {
        $CMUser = Get-CMUser -Name $UserA
        If ($CMUser.ResourceID)
        {
            $CMCollection = Get-CMUserCollection -Id $ColID
            
            If ($CMCollection.Name)
            {
                Try
                {
                    
                    Add-CMUserCollectionDirectMembershipRule -CollectionId $ColID -ResourceId $CMUser.ResourceID
                    Write-Host "User $UserA was added collection $ColID"
                    CD C:
                    $LogText =  "User $UserA was added to Collecion $ColID"
                    log -text $LogText
                    CD $SiteCode":"
                } Catch
                {
                    Write-Host "User $UserA is already in Collecion $ColID"
                    CD C:
                    $LogText =  "User $UserA is already in Collecion $ColID"
                    log -text $LogText
                    CD $SiteCode":"
                }

                                
            } Else
            {
                Write-Host "User Collecion was not found: $ColID"
                CD C:
                $LogText =  "User Collecion was not found: $ColID"
                log -text $LogText              
                CD $SiteCode":"
                #Return $False
            }
             

        } Else
        {
            Write-Host "User not found: $UserA"
            CD C:
            $LogText =  "User not found: $UserA"
            log -text $LogText
            CD $SiteCode":"
            #Return $False
        }
    

    } Catch {
        Write-Host "ERROR: User not found: $UserA"
        CD C:
        $LogText =  "ERROR: User not found: $UserA"
        log -text $LogText
        CD $SiteCode":"
        #Return $False
    }


}

Function RemoveUserFromCollection
{
    Param ([string]$UserA, [string]$ColID)

    Trap { Return "error" }
        
    Try {
        $CMUser = Get-CMUser -Name $UserA
        If ($CMUser.ResourceID)
        {
            $CMCollection = Get-CMUserCollection -Id $ColID
            
            If ($CMCollection.Name)
            {
                
                Remove-CMUserCollectionDirectMembershipRule -CollectionId $ColID -ResourceId $CMUser.ResourceID -Force
                Write-Host "User $UserA was removed from collection $ColID"
                CD C:
                $LogText =  "User $UserA was removed from Collecion $ColID"
                log -text $LogText
                CD $SiteCode":"
                                
            } Else
            {
                Write-Host "User Collecion was not found: $ColID"
                CD C:
                $LogText =  "User Collecion was not found: $ColID"
                log -text $LogText              
                CD $SiteCode":"
                #Return $False
            }
             

        } Else
        {
            Write-Host "User not found: $UserA"
            CD C:
            $LogText =  "User not found: $UserA"
            log -text $LogText
            CD $SiteCode":"
            #Return $False
        }
    

    } Catch {
        Write-Host "ERROR: User not found: $UserA"
        CD C:
        $LogText =  "ERROR: User not found: $UserA"
        log -text $LogText
        CD $SiteCode":"
        #Return $False
    }


}

Function AddDeviceToCollection
{
    Param ([string]$UserA, [string]$ColID)

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
                    CD C:
                    $LogText =  "Device $UserA was added to Collecion $ColID"
                    log -text $LogText
                    CD $SiteCode":"
                } Catch
                {
                    Write-Host "Device $UserA already exists in Collecion $ColID"
                    CD C:
                    $LogText =  "Device $UserA already exists in Collecion $ColID"
                    log -text $LogText
                    CD $SiteCode":"
                }
                                
            } Else
            {
                Write-Host "Device Collecion was not found: $ColID"
                CD C:
                $LogText =  "Device Collecion was not found: $ColID"
                log -text $LogText              
                CD $SiteCode":"
                #Return $False
            }
             

        } Else
        {
            Write-Host "Device not found: $UserA"
            CD C:
            $LogText =  "Device not found: $UserA"
            log -text $LogText
            CD $SiteCode":"
            #Return $False
        }
    

    } Catch {
        Write-Host "ERROR: Device not found: $UserA"
        CD C:
        $LogText =  "ERROR: Device not found: $UserA"
        log -text $LogText
        CD $SiteCode":"
        #Return $False
    }


}

Function RemoveDeviceFromCollection
{
    Param ([string]$UserA, [string]$ColID)

    Trap { Return "error" }
        
    Try {
        $CMUser = Get-CMDevice -Name $UserA
        If ($CMUser.ResourceID)
        {
            $CMCollection = Get-CMDeviceCollection -Id $ColID
            
            If ($CMCollection.Name)
            {
                
                Remove-CMDeviceCollectionDirectMembershipRule -CollectionId $ColID -ResourceId $CMUser.ResourceID -Force
                Write-Host "Device $UserA was removed from collection $ColID"
                CD C:
                $LogText =  "Device $UserA was removed from Collecion $ColID"
                log -text $LogText
                CD $SiteCode":"
                                
            } Else
            {
                Write-Host "Device Collecion was not found: $ColID"
                CD C:
                $LogText =  "Device Collecion was not found: $ColID"
                log -text $LogText              
                CD $SiteCode":"
                #Return $False
            }
             

        } Else
        {
            Write-Host "Device not found: $UserA"
            CD C:
            $LogText =  "Device not found: $UserA"
            log -text $LogText
            CD $SiteCode":"
            #Return $False
        }
    

    } Catch {
        Write-Host "ERROR: Device not found: $UserA"
        CD C:
        $LogText =  "ERROR: Device not found: $UserA"
        log -text $LogText
        CD $SiteCode":"
        #Return $False
    }


}

Function Reset-Log 
{ 
    #function checks to see if file in question is larger than the paramater specified if it is it will roll a log and delete the oldes log if there are more than x logs. 
    param([string]$fileName, [int64]$filesize = 1mb , [int] $logcount = 5) 
     
    $logRollStatus = $true 
    if(test-path $filename) 
    { 
        $file = Get-ChildItem $filename 
        if((($file).length) -ige $filesize) #this starts the log roll 
        { 
            $fileDir = $file.Directory 
            $fn = $file.name #this gets the name of the file we started with 
            $files = Get-ChildItem $filedir | ?{$_.name -like "$fn*"} | Sort-Object lastwritetime 
            $filefullname = $file.fullname #this gets the fullname of the file we started with 
            #$logcount +=1 #add one to the count as the base file is one more than the count 
            for ($i = ($files.count); $i -gt 0; $i--) 
            {  
                #[int]$fileNumber = ($f).name.Trim($file.name) #gets the current number of the file we are on 
                $files = Get-ChildItem $filedir | ?{$_.name -like "$fn*"} | Sort-Object lastwritetime 
                $operatingFile = $files | ?{($_.name).trim($fn) -eq $i} 
                if ($operatingfile) 
                 {$operatingFilenumber = ($files | ?{($_.name).trim($fn) -eq $i}).name.trim($fn)} 
                else 
                {$operatingFilenumber = $null} 
 
                if(($operatingFilenumber -eq $null) -and ($i -ne 1) -and ($i -lt $logcount)) 
                { 
                    $operatingFilenumber = $i 
                    $newfilename = "$filefullname.$operatingFilenumber" 
                    $operatingFile = $files | ?{($_.name).trim($fn) -eq ($i-1)} 
                    write-host "moving to $newfilename" 
                    move-item ($operatingFile.FullName) -Destination $newfilename -Force 
                } 
                elseif($i -ge $logcount) 
                { 
                    if($operatingFilenumber -eq $null) 
                    {  
                        $operatingFilenumber = $i - 1 
                        $operatingFile = $files | ?{($_.name).trim($fn) -eq $operatingFilenumber} 
                        
                    } 
                    write-host "deleting " ($operatingFile.FullName) 
                    remove-item ($operatingFile.FullName) -Force 
                } 
                elseif($i -eq 1) 
                { 
                    $operatingFilenumber = 1 
                    $newfilename = "$filefullname.$operatingFilenumber" 
                    write-host "moving to $newfilename" 
                    move-item $filefullname -Destination $newfilename -Force 
                } 
                else 
                { 
                    $operatingFilenumber = $i +1  
                    $newfilename = "$filefullname.$operatingFilenumber" 
                    $operatingFile = $files | ?{($_.name).trim($fn) -eq ($i-1)} 
                    write-host "moving to $newfilename" 
                    move-item ($operatingFile.FullName) -Destination $newfilename -Force    
                } 
                     
            } 
 
                     
          } 
         else 
         { $logRollStatus = $false} 
    } 
    else 
    { 
        $logrollStatus = $false 
    } 
    $LogRollStatus 
} 




Reset-Log -fileName $logfile -filesize 50mb -logcount 20

$LogText =  "-----$(Get-Date) Running as - $($env:USERNAME) on $($env:COMPUTERNAME) Begin processing the addition of users and computers to SCCM Collections-----"
log -text $LogText



#Update SCCM Deployment Schedule for the Migration Wrapper

#Connection information for SCCM

# Site configuration
$SiteCode = "P01" # Site code 
$ProviderMachineName = "ALCRMCWSSP01.pmusa.net" # SMS Provider machine name

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
    CD C:
    $LogText =  "Connected to the site: $SiteCode"
    log -text $LogText
    

} Catch {
    CD C:
    $LogText =  "ERROR connecting to the site: $SiteCode"
    log -text $LogText
    
}

CD $SiteCode":"

Foreach ($UCObject in $UserAndComputers)
{

    Switch ($UCObject.TYPE)
    {
        "USER"
            {
                Switch ($UCObject.Action)
                {
                    "ADD"
                    {
                        Try {
                            AddUserToCollection -UserA $UCObject.Name -ColID $UCObject.CollectionID

                        } Catch {
                            CD C:
                            Write-host "Unable to find user: $UCObject.Name"
                            $LogText =  "ERROR Finding User: $UCObject.Name"
                            log -text $LogText
                            CD $SiteCode":"
                        }
                    }
<#
                    "REMOVE"
                    {
                        Try {
                            RemoveUserFromCollection -UserA $UCObject.Name -ColID $UCObject.CollectionID
                        } Catch {
                            CD C:
                            Write-host "Unable to find user: $UCObject.Name"
                            $LogText =  "ERROR Finding User: $UCObject.Name"
                            log -text $LogText
                            CD $SiteCode":"
                        }
                    }
#>
                }
            }
        "DEVICE"
        {
                Switch ($UCObject.Action)
                {
                    "ADD"
                    {
                         Try {
                            AddDeviceToCollection -UserA $UCObject.Name -ColID $UCObject.CollectionID

                        } Catch {
                            CD C:
                            Write-host "Unable to find Device: $UCObject.Name"
                            $LogText =  "ERROR Finding Device: $UCObject.Name"
                            log -text $LogText
                            CD $SiteCode":"
                        }
                    }
<#                    
                    "REMOVE"
                    {
                            Try {
                                RemoveDeviceFromCollection -UserA $UCObject.Name -ColID $UCObject.CollectionID

                            } Catch {
                                CD C:
                                Write-host "Unable to find Device: $UCObject.Name"
                                $LogText =  "ERROR Finding Device: $UCObject.Name"
                                log -text $LogText
                                CD $SiteCode":"
                            }
                     } 
#>
                }
        }
    }

}

CD C:
$LogText =  "-----$(Get-Date) Running as - $($env:USERNAME) on $($env:COMPUTERNAME) Finished adding users to AD and updating the SCCM Schedule update-----"
log -text $LogText
