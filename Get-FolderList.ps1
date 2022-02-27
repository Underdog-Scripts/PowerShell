$FolderName = 'Retired' #Folder to scan
$ObjectType = '5000' #Collection ObjectType
$SiteServer = 'ALCRMCWSSP02'

$FolderIDQuery = Get-WmiObject -Namespace "Root\SMS\P01" -Class SMS_ObjectContainernode -Filter "Name='$FolderName' and ObjectType='$ObjectType'" -ComputerName $SiteServer
    
$ItemsInFolder = Get-WmiObject -Namespace "Root\SMS\P01" -Class SMS_ObjectContainerItem -Filter "ContainerNodeID='$($FolderIDQuery.ContainerNodeID)' and ObjectType='$ObjectType'" -ComputerName $SiteServer

foreach($item in $ItemsInFolder)
    {
     Write-Host $item
     }