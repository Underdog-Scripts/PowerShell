<#
.Synopsis
   This script allows for changes to SCCM collections
.DESCRIPTION
   This script allows for changes to SCCM collections by adding and removing from user and device collections
.PARAMETER List of device or user names
   Reads text file containing a list of devices or user names to be added or removed from specific collection
.INPUTS
   Path to text file to be read and container name
.OUTPUTS
   Log file of each actions result
.NOTES
   Version:        1.0
   Author:         GTB
   Creation Date:  12.5.18
   Purpose/Change: Initial script development
.EXAMPLE
   ./UpdateCollection.ps1
.EXAMPLE
   ./UpdateCollection.ps1
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ConfigMgr 3 character site code and site server
$SiteCode = "P01" 
$SiteServer = "ALCRMCWSSP01.PMUSA.NET"

# Default input file and log file paths
$DefaultFilePref = "C:\Users\brunnegt\Documents\Scripts\PowerShell\Add-Users"
$UserAndComputers = Import-Csv -LiteralPath $DefaultFilePref".csv"
$option = [System.StringSplitOptions]::RemoveEmptyEntries

#Logfile information
$logfile = $DefaultFilePref+"_"+(Get-Date -Format MM.dd.yyyy_hh.mm.ss)+".log"

#######################################################################################################################

Function Log {
   param([Switch]$fout, [String]$text)
   AC $logfile $text
      if($showConsoleOutput) {
         
         if($fout) {
           
            Write-Host $text -ForegroundColor Red
         } else {
            
            Write-Host $text -ForegroundColor Green
         }
         
      }

   }

#######################################################################################################################

# AddUserToCollection function - (from MassAdd-UsersandComputersToCollections.ps1)
Function AddUserToCollection {
   Param ([string]$UserA, [string]$ColID)
   Trap { Return "Input File Value Error" }
      Try {
         $CMUser = Get-CMUser -Name $UserA
		    If ($CMUser.ResourceID) {
               $CMCollection = Get-CMUserCollection -Id $ColID
                  If ($CMCollection.Name) {
                     Try {
                        Add-CMUserCollectionDirectMembershipRule -CollectionId $ColID -ResourceId $CMUser.ResourceID
                        Write-Host "User $UserA was added Collection $ColID"
                        CD C:
                        $LogText =  "User $UserA was added to Collection $ColID"
                        log -text $LogText
                        CD $SiteCode":"
                     } Catch {
                
                        Write-Host "User $UserA is already in Collection $ColID"
                        CD C:
                        $LogText =  "User $UserA is already in Collection $ColID"
                        log -text $LogText
                        CD $SiteCode":"
                     }
                  } Else {
               
                     Write-Host "User Collection was not found: $ColID"
                     CD C:
                     $LogText =  "User Collection was not found: $ColID"
                     log -text $LogText              
                     CD $SiteCode":"
                     #Return $False
                  }
            } Else {
         
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
         $LogText = "ERROR: User not found: $UserA"
         log -text $LogText
         CD $SiteCode":"
         #Return $False
         
      }

}

#######################################################################################################################
             
Function RemoveUserFromCollection {
   Param ([string]$UserA, [string]$ColID)
   Trap { Return "error" }
      Try {
         $CMUser = Get-CMUser -Name $UserA
            If ($CMUser.ResourceID) {
               $CMCollection = Get-CMUserCollection -Id $ColID
                  If ($CMCollection.Name) {
                   
                        Remove-CMUserCollectionDirectMembershipRule -CollectionId $ColID -ResourceId $CMUser.ResourceID -Force
                        Write-Host "User $UserA was removed from Collection $ColID"
                        CD C:
                        $LogText =  "User $UserA was removed from Collection $ColID"
                        log -text $LogText
                        CD $SiteCode":"
                  } Else {
               
               Write-Host "User not found: $UserA"
               CD C:
               $LogText = "User not found: $UserA"
               log -text $LogText
               CD $SiteCode":"
               #Return $False
                  }
            } Else {

            Write-Host "User not found: $UserA"
            CD C:
            $LogText = "User not found: $UserA"
            log -text $LogText
            CD $SiteCode":"
            #Return $False
            }
      } Catch {
        
         Write-Host "ERROR: User not found: $UserA"
         CD C:
         $LogText = "ERROR: User not found: $UserA"
         Log -text $LogText
         CD $SiteCode":"
         #Return $False
         
      }

}

#######################################################################################################################

Function AddDeviceToCollection {
   Param ([string]$DeviceA, [string]$ColID)
   Trap { Return "error" }
      Try {
         $CMDevice = Get-CMDevice -Name $DeviceA
            If ($CMDevice.ResourceID) {
               $CMCollection = Get-CMDeviceCollection -Id $ColID
                  If ($CMCollection.Name) {
                     Try {
                        Add-CMDeviceCollectionDirectMembershipRule -CollectionId $ColID -ResourceId $CMDevice.ResourceID
                        Write-Host "Device $DeviceA was added to Collection $ColID"
                        CD C:
                        $LogText = "Device $DeviceA was added to Collection $ColID"
                        log -text $LogText
                        CD $SiteCode":"
                     } Catch {
                     
                        Write-Host "Device $DeviceA not in Collection $ColID"
                        CD C:
                        $LogText = "Device $DeviceA not in Collection $ColID"
                        Log -text $LogText
                        CD $SiteCode":"
                     }
                  } Else {
                  
                     Write-Host "Device Collection was not found: $ColID"
                     CD C:
                     $LogText = "Device Collection was not found: $ColID"
                     log -text $LogText              
                     CD $SiteCode":"
                     #Return $False
                  }
            } Else {
               
               Write-Host "Device not found: $DeviceA"
               CD C:
               $LogText = "Device not found: $DeviceA"
               log -text $LogText
               CD $SiteCode":"
               #Return $False
            }
      } Catch {
        
         Write-Host "ERROR: Device not found: $DeviceA"
         CD C:
         $LogText = "ERROR: Device not found: $DeviceA"
         Log -text $LogText
         CD $SiteCode":"
         #Return $False
         
      }

}

#######################################################################################################################

Function RemoveDeviceFromCollection {
   Param ([string]$DeviceA, [string]$ColID)
   Trap { Return "error" }
      Try {
         $CMDevice = Get-CMDevice -Name $DeviceA
            If ($CMDevice.ResourceID) {
               $CMCollection = Get-CMDeviceCollection -Id $ColID
                  If ($CMCollection.Name) {
                        Remove-CMDeviceCollectionDirectMembershipRule -CollectionId $ColID -ResourceId $CMDevice.ResourceID -Force
                        Write-Host "Device $DeviceA was removed from Collection $ColID"
                        CD C:
                        $LogText = "Device $DeviceA was removed from Collection $ColID"
                        log -text $LogText
                        CD $SiteCode":"
                  } Else {
                  
                     Write-Host "Device Collection was not found: $ColID"
                     CD C:
                     $LogText = "Device Collection was not found: $ColID"
                     log -text $LogText              
                     CD $SiteCode":"
                     #Return $False
                  }
            } Else {
               
               Write-Host "Device not found: $DeviceA"
               CD C:
               $LogText = "Device not found: $DeviceA"
               log -text $LogText
               CD $SiteCode":"
               #Return $False
            }
      } Catch {
        
         Write-Host "ERROR: Device not found: $DeviceA"
         CD C:
         $LogText = "ERROR: Device not found: $DeviceA"
         Log -text $LogText
         CD $SiteCode":"
         #Return $False
         
      }

}

#######################################################################################################################

# A function to create the form 
function Update_Collection{
   [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
   [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    
   # Set the size of your form
   $Form = New-Object System.Windows.Forms.Form
   $Form.width = 840
   $Form.height = 360
   $Form.Text = 'Please select a task'

   # Create label to display SCCM site code
   $label0 = New-Object System.Windows.Forms.Label
   $label0.Location = New-Object System.Drawing.Point(140,14)
   $label0.Size = New-Object System.Drawing.Size(120,20)
   $label0.Text = 'Site Code = P01'
   $form.Controls.Add($label0)

   # Create label to display SCCM site server
   $label1 = New-Object System.Windows.Forms.Label
   $label1.Location = New-Object System.Drawing.Point(340,14)
   $label1.Size = New-Object System.Drawing.Size(240,20)
   $label1.Text = 'Site Server = ALCRMCWSSP01.PMUSA.NET'
   $form.Controls.Add($label1)
 
   # Create text for second button group
   $label3 = New-Object System.Windows.Forms.Label
   $label3.Location = New-Object System.Drawing.Point(40,48)
   $label3.Size = New-Object System.Drawing.Size(300,14)
   $label3.Text = 'Using (Name, Type, Action, CollectionID) CVS file format'
   $form.Controls.Add($label3)
   
   # Create underline for group name   
   $label3a = New-Object System.Windows.Forms.Label
   $label3a.Location = New-Object System.Drawing.Point(40,51)
   $label3a.Size = New-Object System.Drawing.Size(300,20)
   $label3a.Text = '_______________________________________________________'
   $form.Controls.Add($label3a)
   
   # Create label for text file path
   $label4 = New-Object System.Windows.Forms.Label
   $label4.Location = New-Object System.Drawing.Point(40,70)
   $label4.Size = New-Object System.Drawing.Size(280,20)
   $label4.Text = 'Please enter/verify the path to the CSV file below:'
   $form.Controls.Add($label4)

   # Create input box for input file path
   $textBox0 = New-Object System.Windows.Forms.TextBox
   $textBox0.Location = New-Object System.Drawing.Point(40,90)
   $textBox0.Size = New-Object System.Drawing.Size(780,20)
   $textBox0.Text = $DefaultFilePref+".csv"
   $form.Controls.Add($textBox0)
   
   # Create the collection of radio buttons      
   # Create text group box
   $textBox = New-Object System.Windows.Forms.TextBox
   $textBox.Location = New-Object System.Drawing.Point(40,90) 
   $MyGroupBox = New-Object System.Windows.Forms.GroupBox
   $MyGroupBox.Location = '40,140'
   $MyGroupBox.size = '740,120'
   $MyGroupBox.text = 'Available Tasks'
   
   # Create text for first button group
   $label2 = New-Object System.Windows.Forms.Label
   $label2.Location = New-Object System.Drawing.Point(60,160)
   $label2.Size = New-Object System.Drawing.Size(300,14)
   $label2.Text = 'Using (Name, CollectionID) CVS file format'
   $form.Controls.Add($label2)
   
   # Create underline for group name
   $label2a = New-Object System.Windows.Forms.Label
   $label2a.Location = New-Object System.Drawing.Point(60,163)
   $label2a.Size = New-Object System.Drawing.Size(300,20)
   $label2a.Text = '____________________________________'
   $form.Controls.Add($label2a)
      
   # Add Users to a Collection - Radio Button
   $RadioButton1 = New-Object System.Windows.Forms.RadioButton
   $RadioButton1.Location = '20,40'
   $RadioButton1.size = '240,20'
   $RadioButton1.Checked = $false 
   $RadioButton1.Text = 'Add Users to a Collection'
    
   # Remove Users from a Collection - Radio Button
   $RadioButton2 = New-Object System.Windows.Forms.RadioButton
   $RadioButton2.Location = '370,40'
   $RadioButton2.size = '240,20'
   $RadioButton2.Checked = $false 
   $RadioButton2.Text = 'Remove Users from a Collection'
    
   # Add Devices to a Collection - Radio Button
   $RadioButton3 = New-Object System.Windows.Forms.RadioButton
   $RadioButton3.Location = '20,60'
   $RadioButton3.size = '220,20'
   $RadioButton3.Checked = $false
   $RadioButton3.Text = 'Add Devices to a Collection'
    
   # Remove Device from a Collection - Radio Button
   $RadioButton4 = New-Object System.Windows.Forms.RadioButton
   $RadioButton4.Location = '370,60'
   $RadioButton4.size = '220,20'
   $RadioButton4.Checked = $false
   $RadioButton4.Text = 'Remove Device from a Collection'
     
   # Placeholder 1 - Radio Button
   $RadioButton5 = New-Object System.Windows.Forms.RadioButton
   $RadioButton5.Location = '20,80'
   $RadioButton5.size = '320,20'
   $RadioButton5.Checked = $false
   $RadioButton5.Text = 'Placeholder 1'
    
   # Placeholder 2 - Radio Button
   $RadioButton6 = New-Object System.Windows.Forms.RadioButton
   $RadioButton6.Location = '370,80'
   $RadioButton6.size = '320,20'
   $RadioButton6.Checked = $false
   $RadioButton6.Text = 'Placeholder 2'

   # Add an OK button
   $OKButton = new-object System.Windows.Forms.Button
   $OKButton.Location = '295,280'
   $OKButton.Size = '75,23'
   $OKButton.Text = 'OK'
   $OKButton.DialogResult=[System.Windows.Forms.DialogResult]::OK
 
   # Add a cancel button
   $CancelButton = new-object System.Windows.Forms.Button
   $CancelButton.Location = '450,280'
   $CancelButton.Size = '75,23'
   $CancelButton.Text = 'Cancel'
   $CancelButton.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
 
   # Add all the Form controls on one line 
   $form.Controls.AddRange(@($MyGroupBox,$OKButton,$CancelButton))
 
   # Add all the GroupBox controls on one line
   $MyGroupBox.Controls.AddRange(@($Radiobutton1,$RadioButton2,$RadioButton3,$RadioButton4,$RadioButton5,$RadioButton6))
    
   # Assign the Accept and Cancel options in the form to the corresponding buttons
   $form.AcceptButton = $OKButton
   $form.CancelButton = $CancelButton
 
   # Activate the form
   $form.Add_Shown({$form.Activate()})    
    
   # Get the results from the button click
   $dialogResult = $form.ShowDialog()
 
   # If the OK button is selected 
   if ($dialogResult -eq "OK") {
      if ($RadioButton1.Checked) { 
         # If the button 1 is selected 11111111111111111111111111111111111111111111111111111111111111111111111111111111111111111
         # Customizations
         $initParams = @{}
         $ProviderMachineName = "ALCRMCWSSP01.PMUSA.NET"
            if((Get-Module ConfigurationManager) -eq $null) {
               # Import the ConfigurationManager.psd1 module
               Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
            }

            if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
               # Connect to the site's drive if it is not already present
               New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
            }
            
            #Inputfile information
            $AddUserFilePref = "C:\Users\brunnegt\Documents\Scripts\PowerShell\Add-Users"
            $UserAndComputers = Import-Csv -LiteralPath $AddUserFilePref".csv"
            $option = [System.StringSplitOptions]::RemoveEmptyEntries

            #Logfile information
            $logfile = $AddUserFilePref+"_"+(Get-Date -Format MM.dd.yyyy_hh.mm.ss)+".log"
         # Set the current location to be the site code.
         Set-Location "$($SiteCode):\" @initParams

            Try{
               # Set the current location to be the site code.
               Set-Location "$($SiteCode):\" @initParams
               CD C:
               $LogText = "Connected to the site: $SiteCode"
               log -text $LogText
            } Catch {
            
               CD C:
               $LogText = "ERROR connecting to the site: $SiteCode"
               log -text $LogText
            }

         CD $SiteCode":"
         Foreach ($UCObject in $UserAndComputers) {
            Try {
               AddUserToCollection -UserA $UCObject.Name -ColID $UCObject.CollectionID
            } Catch {
               CD C:
               Write-host "Unable to find user: $UCObject.Name"
               $LogText = "ERROR Finding User: $UCObject.Name"
               log -text $LogText
               CD $SiteCode":"
            }

         }    

      CD C:
      $LogText =  "-----$(Get-Date) Running as - $($env:USERNAME) on $($env:COMPUTERNAME) Finished adding users to AD and updating the SCCM Schedule update-----"
      log -text $LogText

      # 11111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111

      } elseif ($RadioButton2.Checked) { # If the button 2 is selected 22222222222222222222222222222222222222222222222222222222222222222222222222222222222
         # Customizations
         $initParams = @{}
         $ProviderMachineName = "ALCRMCWSSP01.PMUSA.NET"
            if((Get-Module ConfigurationManager) -eq $null) {
               # Import the ConfigurationManager.psd1 module
               Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
            }
            if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
               # Connect to the site's drive if it is not already present
               New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
            }
            
            #Inputfile information
            $RemUserFilePref = "C:\Users\brunnegt\Documents\Scripts\PowerShell\Remove-Users"
            $UserAndComputers = Import-CSV -LiteralPath $RemUserFilePref".csv"
            $option = [System.StringSplitOptions]::RemoveEmptyEntries
            
            #Logfile information
            $logfile = $RemUserFilePref+"_"+(Get-Date -Format MM.dd.yyyy_hh.mm.ss)+".log"

         # Set the current location to be the site code.
         Set-Location "$($SiteCode):\" @initParams
            Try{
               # Set the current location to be the site code.
               Set-Location "$($SiteCode):\" @initParams
               CD C:
               $LogText = "Connected to the site: $SiteCode"
               log -text $LogText
            } Catch {
            
               CD C:
               $LogText = "ERROR connecting to the site: $SiteCode"
               log -text $LogText
            }

         CD $SiteCode":"
         Foreach ($UCObject in $UserAndComputers) {
            Try {
               RemoveUserFromCollection -UserA $UCObject.Name -ColID $UCObject.CollectionID

            } Catch {
               CD C:
               Write-host "Unable to find user: $UCObject.Name"
               $LogText = "ERROR Finding User: $UCObject.Name"
               log -text $LogText
               CD $SiteCode":"
            }

         }
         
      CD C:
      $LogText =  "-----$(Get-Date) Running as - $($env:USERNAME) on $($env:COMPUTERNAME) Finished adding users to AD and updating the SCCM Schedule update-----"
      log -text $LogText
      
      # 22222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222      

      } elseif ($RadioButton3.Checked) { # If the button 3 is selected 33333333333333333333333333333333333333333333333333333333333333333333333333333333333
         # Customizations
         $initParams = @{}
         $ProviderMachineName = "ALCRMCWSSP01.PMUSA.NET"
            if((Get-Module ConfigurationManager) -eq $null) {
               # Import the ConfigurationManager.psd1 module
               Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
            }
            if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
               # Connect to the site's drive if it is not already present
               New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
            }
            
            #Inputfile information
            $AddDevFilePref = "C:\Users\brunnegt\Documents\Scripts\PowerShell\Add-Devices"
            $UserAndComputers = Import-CSV -LiteralPath $AddDevFilePref".csv"
            $option = [System.StringSplitOptions]::RemoveEmptyEntries

            #Logfile information
            $logfile = $AddDevFilePref+"_"+(Get-Date -Format MM.dd.yyyy_hh.mm.ss)+".log"

         # Set the current location to be the site code.
         Set-Location "$($SiteCode):\" @initParams
            Try{
               # Set the current location to be the site code.
               Set-Location "$($SiteCode):\" @initParams
               CD C:
               $LogText = "Connected to the site: $SiteCode"
               log -text $LogText
            } Catch {
            
               CD C:
               $LogText = "ERROR connecting to the site: $SiteCode"
               log -text $LogText
            }

         CD $SiteCode":"
         Foreach ($UCObject in $UserAndComputers) {
            Try {
               # AddDeviceToCollection -DeviceA $UCObject.Name -ColID $UCObject.CollectionID
               AddDeviceToCollection -DeviceA $UCObject.Name -ColID $UCObject.CollectionID
            } Catch {

               CD C:
               Write-host "Unable to find device: $UCObject.Name"
               $LogText = "ERROR Finding device: $UCObject.Name"
               log -text $LogText
               CD $SiteCode":"
            }

         }
         
      CD C:
      $LogText =  "-----$(Get-Date) Running as - $($env:USERNAME) on $($env:COMPUTERNAME) Finished adding users to AD and updating the SCCM Schedule update-----"
      log -text $LogText   
               
      # 33333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333

      } elseif ($RadioButton4.Checked) { # If the button 4 is selected 44444444444444444444444444444444444444444444444444444444444444444444444444444444444
         # Customizations
         $initParams = @{}
         $ProviderMachineName = "ALCRMCWSSP01.PMUSA.NET"
            if((Get-Module ConfigurationManager) -eq $null) {
               # Import the ConfigurationManager.psd1 module
               Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
            }
            if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
               # Connect to the site's drive if it is not already present
               New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
            }
            
            #Inputfile information
            $RemDevFilePref = "C:\Users\brunnegt\Documents\Scripts\PowerShell\Remove-Devices"
            $UserAndComputers = Import-CSV -LiteralPath $RemDevFilePref".csv"
            $option = [System.StringSplitOptions]::RemoveEmptyEntries

            #Logfile information
            $logfile = $RemDevFilePref+"_"+(Get-Date -Format MM.dd.yyyy_hh.mm.ss)+".log"

         # Set the current location to be the site code.
         Set-Location "$($SiteCode):\" @initParams
            Try{
               # Set the current location to be the site code.
               Set-Location "$($SiteCode):\" @initParams
               CD C:
               $LogText = "Connected to the site: $SiteCode"
               log -text $LogText
            } Catch {
            
               CD C:
               $LogText = "ERROR connecting to the site: $SiteCode"
               log -text $LogText
            }

         CD $SiteCode":"
         Foreach ($UCObject in $UserAndComputers) {
            Try {
               RemoveDeviceFromCollection -DeviceA $UCObject.Name -ColID $UCObject.CollectionID

            } Catch {
               CD C:
               Write-host "Unable to find device: $UCObject.Name"
               $LogText = "ERROR Finding device: $UCObject.Name"
               log -text $LogText
               CD $SiteCode":"
            }

         }
         
      CD C:
      $LogText =  "-----$(Get-Date) Running as - $($env:USERNAME) on $($env:COMPUTERNAME) Finished adding users to AD and updating the SCCM Schedule update-----"
      log -text $LogText
  
      # 44444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444  
               
      } elseif ($RadioButton5.Checked) { # If the button 5 is selected 55555555555555555555555555555555555555555555555555555555555555555555555555555555555
         # Customizations
         $initParams = @{}
         $ProviderMachineName = "ALCRMCWSSP01.PMUSA.NET"
            if((Get-Module ConfigurationManager) -eq $null) {
               # Import the ConfigurationManager.psd1 module
               Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
            }
            if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
               # Connect to the site's drive if it is not already present
               New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
            }
            
            #Inputfile information
            $PH1FilePref = "C:\Users\brunnegt\Documents\Scripts\PowerShell\PH1"
            $UserAndComputers = Import-CSV -LiteralPath $PH1FilePref".csv"
            $option = [System.StringSplitOptions]::RemoveEmptyEntries

            #Logfile information
            $logfile = $PH1FilePref+"_"+(Get-Date -Format MM.dd.yyyy_hh.mm.ss)+".log"

         # Set the current location to be the site code.
         Set-Location "$($SiteCode):\" @initParams
            Try{
               # Set the current location to be the site code.
               Set-Location "$($SiteCode):\" @initParams
               CD C:
               $LogText = "Connected to the site: $SiteCode"
               log -text $LogText
            } Catch {
            
               CD C:
               $LogText = "ERROR connecting to the site: $SiteCode"
               log -text $LogText
            }

         CD $SiteCode":"
         Foreach ($UCObject in $UserAndComputers) {
            Try {
               RemoveUserFromCollection -UserA $UCObject.Name -ColID $UCObject.CollectionID

             } Catch {
                CD C:
                Write-host "Unable to find Device: $UCObject.Name"
                $LogText = "ERROR Finding Device: $UCObject.Name"
                log -text $LogText
                CD $SiteCode":"
             }

         }
         
      CD C:
      $LogText =  "-----$(Get-Date) Running as - $($env:USERNAME) on $($env:COMPUTERNAME) Finished adding users to AD and updating the SCCM Schedule update-----"
      log -text $LogText
      
      # 55555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555
          
      } elseif ($RadioButton6.Checked) { # If the button 6 is selected 66666666666666666666666666666666666666666666666666666666666666666666666666666666666
      # Customizations
         $initParams = @{}
         $ProviderMachineName = "ALCRMCWSSP01.PMUSA.NET"
            if((Get-Module ConfigurationManager) -eq $null) {
               # Import the ConfigurationManager.psd1 module
               Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
            }
            if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
               # Connect to the site's drive if it is not already present
               New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
            }
            
            #Inputfile information
            $PH2FilePref = "C:\Users\brunnegt\Documents\Scripts\PowerShell\PH2"
            $UserAndComputers = Import-CSV -LiteralPath $PH2FilePref".csv"
            $option = [System.StringSplitOptions]::RemoveEmptyEntries

            #Logfile information
            $logfile = $PH2FilePref+"_"+(Get-Date -Format MM.dd.yyyy_hh.mm.ss)+".log"

         # Set the current location to be the site code.
         Set-Location "$($SiteCode):\" @initParams
            Try{
               # Set the current location to be the site code.
               Set-Location "$($SiteCode):\" @initParams
               CD C:
               $LogText = "Connected to the site: $SiteCode"
               log -text $LogText
            } Catch {
            
               CD C:
               $LogText = "ERROR connecting to the site: $SiteCode"
               log -text $LogText
            }

         CD $SiteCode":"
         Foreach ($UCObject in $UserAndComputers) {
            Try {
               RemoveUserFromCollection -UserA $UCObject.Name -ColID $UCObject.CollectionID

             } Catch {
                CD C:
                Write-host "Unable to find Device: $UCObject.Name"
                $LogText =  "ERROR Finding Device: $UCObject.Name"
                log -text $LogText
                CD $SiteCode":"
             }

         }
         
      CD C:
      $LogText =  "-----$(Get-Date) Running as - $($env:USERNAME) on $($env:COMPUTERNAME) Finished adding users to AD and updating the SCCM Schedule update-----"
      log -text $LogText
      
      # 66666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666         
           
      } else {
      
         [System.Windows.Forms.MessageBox]::Show("No Radio Button Selected")
      
      }
      
   }

}

#######################################################################################################################

Function Reset-Log 
{ 
   #function checks to see if file in question is larger than the paramater specified if it is it will roll a log and delete the oldes log if there are more than x logs. 
   param([string]$fileName, [int64]$filesize = 1mb , [int] $logcount = 5) 
     
   $logRollStatus = $true 
   if(test-path $filename) { 
      $file = Get-ChildItem $filename 
         if((($file).length) -ige $filesize) { #this starts the log roll  
            $fileDir = $file.Directory 
            $fn = $file.name #this gets the name of the file we started with 
            $files = Get-ChildItem $filedir | ?{$_.name -like "$fn*"} | Sort-Object lastwritetime 
            $filefullname = $file.fullname #this gets the fullname of the file we started with 
            #$logcount +=1 #add one to the count as the base file is one more than the count 
               for ($i = ($files.count); $i -gt 0; $i--) {  
                  #[int]$fileNumber = ($f).name.Trim($file.name) #gets the current number of the file we are on 
                  $files = Get-ChildItem $filedir | ?{$_.name -like "$fn*"} | Sort-Object lastwritetime 
                  $operatingFile = $files | ?{($_.name).trim($fn) -eq $i} 
                     if ($operatingfile) {
                        $operatingFilenumber = ($files | ?{($_.name).trim($fn) -eq $i}).name.trim($fn)
                     } else {
                        $operatingFilenumber = $null
                     } 
 
                  if(($operatingFilenumber -eq $null) -and ($i -ne 1) -and ($i -lt $logcount)) { 
                     $operatingFilenumber = $i 
                     $newfilename = "$filefullname.$operatingFilenumber" 
                     $operatingFile = $files | ?{($_.name).trim($fn) -eq ($i-1)} 
                     write-host "moving to $newfilename" 
                     move-item ($operatingFile.FullName) -Destination $newfilename -Force 
                  } elseif($i -ge $logcount) {
                 
                     if($operatingFilenumber -eq $null) {  
                        $operatingFilenumber = $i - 1 
                        $operatingFile = $files | ?{($_.name).trim($fn) -eq $operatingFilenumber
                     } 
                        
                  } 
                     write-host "deleting " ($operatingFile.FullName) 
                     remove-item ($operatingFile.FullName) -Force 
                  } elseif($i -eq 1) {
                 
                     $operatingFilenumber = 1 
                     $newfilename = "$filefullname.$operatingFilenumber" 
                      write-host "moving to $newfilename" 
                      move-item $filefullname -Destination $newfilename -Force 
                  } else {
                 
                     $operatingFilenumber = $i +1  
                     $newfilename = "$filefullname.$operatingFilenumber" 
                     $operatingFile = $files | ?{($_.name).trim($fn) -eq ($i-1)} 
                     write-host "moving to $newfilename" 
                     move-item ($operatingFile.FullName) -Destination $newfilename -Force    
                  } 
                     
               } 
 
                     
         } else { 
            $logRollStatus = $false 
         }
          
    } else { 
       $logrollStatus = $false 
    }
     
    $LogRollStatus
     
} 

#######################################################################################################################

Update_Collection
