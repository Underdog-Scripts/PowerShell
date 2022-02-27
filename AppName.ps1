# Export-CMApplication -Name ".NET Data Provider for Teradata v14.10" -Path "C:\Documents\Scripts\PowerShell\TestExport.zip" -Verbose
# Get-CMApplication -Name ".NET Data Provider for Teradata v14.10"
# Get-CMApplication -Name ".NET Data Provider for Teradata v14.10"
# Move-CMObject -FolderPath "P01:\Software Library\Overview\Application Management\Applications\WKS\4.  RETIRED" -ObjectId "65535"
Select[ContainerNodeID] 
,[Name]
,[ObjectType]
,[ParentContainerNodeID]
,[FolderGUID]
FROM [CM_P01].[dbo].[vSMS_Folders]
Where Name Like '4.  RETIRED'