D:\Scripts\PS1\LoadGUI.PS1 -XamlPath 'D:\Scripts\PS1\TestMenu.xaml'

#Set default settings
$VMName = "NG00244194-VM8"
$VMStartupMemory = 4
$VMGeneration = 2
$VMBootDevice = "CD"
$VMHostComputerName = "NG00244194"
$VMProcessors = 2
$BootCDPath = "c:\mount\bootiso\IVANTI_WinPE_WAP01761_20171109.iso"
$VMVHDPath = "D:\VirtualPS\UEFIHDD\UEFIVM8\$VMName\Virtual Hard Disks\$VMName.vhdx"
$VMVHDSize = 32
$VMTPMEnable = $True
$VMSwitchName = "New Virtual Switch"
$VMMacAddress="000000000000"

#Set defaults into menu
$VMNameTextbox.Text = $VMName
$VMGenerationTextbox.Text = $VMGeneration
$VMStartMemoryTextbox.Text = $VMStartupMemory
$VMProcessorsTextBox.Text = $VMProcessors
$BootISOTextbox.Text = $BootCDPath
$VMVHDLocationTextBox.Text = $VMVHDPath
$VMVHDSizeTextBox.Text = $VMVHDSize
$VMVSwitchTextBox.Text = $VMSwitchName
$VMTPMEnableCheckbox.IsChecked = $VMTPMEnable

$SubmitButton.add_Click({
    $xamGUI.close()

})

$xamGUI.ShowDialog()|out-null

#Set values to user selected
$VMName = $VMNameTextbox.Text
$VMStartupMemory = $VMStartMemoryTextbox.Text
$VMGeneration = $VMGenerationTextbox.Text
$VMBootDevice = "CD"
$VMProcessors = $VMProcessorsTextBox.Text
$BootCDPath = $BootISOTextbox.Text
$VMVHDPath = $VMVHDLocationTextBox.Text
$VMVHDSize = $VMVHDSizeTextBox.Text
$VMTPMEnable = $VMTPMEnableCheckbox.IsChecked
$VMSwitchName = $VMVSwitchTextBox.Text


if (test-path $VMVHDPath){
    $VHD = get-vhd -Path $VMVHDPath
}else{
    new-VHD -path $VMVHDPath -SizeBytes ([int64]$VMVHDSize * 1gb) -ComputerName $VMHostComputerName |out-null
}

New-VM -name $VMName -MemoryStartupBytes ([int64]$VMStartupMemory * 1gb) -Generation $VMGeneration -BootDevice $VMBootDevice -Computername $VMHostComputerName -SwitchName $VMSwitchName
set-vmDvdDrive -vmname $VMName -controllernumber 0 -path $BootCDPath 
get-vm $VMName| Set-VMProcessor -count $VMProcessors
add-VMHardDiskDrive -VMName $VMName -path $VMVHDPath 

if($VMTPMEnable){
    set-VMKeyProtector -VMName $VMName -NewLocalKeyProtector
    enable-vmtpm -VMName $VMName
}

#start-vm -vmname $VMName

while ((get-vmnetworkadapter -vmname $VMName).Macaddress -like "000000000000"){
    write-host "Waiting 5 seconds"
    start-sleep -s 5
}

$VMMacAddress=(get-vmnetworkadapter -vmname $VMName).Macaddress
#get-vm $VMName |get-member
write-host "Mac : " $VMMacAddress