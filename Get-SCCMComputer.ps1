function Get-SCCMComputer
{

	<#
	.Synopsis
	Queries SCCM for computers using various input fields.

	.Description
		Takes any combination of parameters specified at the commandline and builds a query 
		that is then run against the SCCM database and returns the results.
	        
	.Parameter ComputerName
		Search for computer(s) by name. Accepts wildcards.

	.Parameter OperatingSystem
        Applies a filter on the OperatingSystem field to limit returns that match the specified OS.
			Accepts wildcards.
		
	.Parameter ComputerManufacturer
		Applies a filter on the ComputerManufacturer field to limit returns that match the specified manufactuer.
			Accepts wildcards.
		
	.Parameter ComputerModel
		Applies a filter on the ComputerModel field to limit returns that match the specified model.
			Accepts wildcards.
		
	.Parameter ComputerSerial
		Applies a filter on the ComputerNumber field to limit returns that match the specified serial number.
			Accepts wildcards.
		
	.Parameter LastUser
		Returns all computers that were last logged into by the specified user. Accepts wildcards, useful for including
			ADM user accounts in searches.
		
	.Parameter TopUser
		Returns the most logged in user by time. This can be useful for determining the primary user of a machine, although
			it is not considered foolproof, especially in cases where users may us RunAs frequently.
        If used as the only paramater, to find machines logged into by account, a wildcard is automatically appended to the
        start of the input data in order to match

    .Paramater Domain
        The TopUser field in the SCCM database contains the domain of the logged in user by default. Use this parameter to 
        specify an alternate domain or a wildcard.

	.Switch IncludeServers
		Servers are filtered out of queries by default, unless this switch is included.
		
		
    .Example
		# Return information about a single computer.
		Get-SCCMComputer USSP-1MJ73R1-L
    
    .Example
        # Return a list of computers that match a specified model number.
        Get-SCCMComputer -ComputerModel "*Latitude E6520*"
		
    .Example
        # Return a list of computers last logged into with a specific account.
        Get-SCCMComputer -LastUser bauera01

    .Example
        # Use the -IncludeServers switch if the expected data may also include servers.
		Get-SCCMComputer -ComputerModel "*PowerEdge R630*" -IncludeServers
    
   
                         
	.NOTES
       ===========================================================================
       Created on:          10/28/2015
       Last Updated:        02/08/2017
       Created by:          Aaron Bauer
       Organization:        SJM
       Filename:            Get-SCCMComputer
       Revision No:         1.0
       Modifications:       1.0: Initial release
                            1.1: Added support for Lenovo machines that store their
                                  model information in a different field from every
                                  other manufacturer in the world
                            1.2: Added support for declaring Chassis Types
                            1.3: Added support for returning CPU, RAM, and Disk infos
                            1.4: Included filtering for docking station chassis types (12)
                                  by default
       ===========================================================================
	#>

    [CmdletBinding()]
    PARAM(
        [Parameter(ValueFromPipelineByPropertyName=$true,
                    ValueFromPipeline=$true)]
        [Alias("CompName")]
        [String]$ComputerName
        ,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                    ValueFromPipeline=$true)]
        [Alias("OS")]
        [String]$OperatingSystem
        ,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                    ValueFromPipeline=$true)]
        [Alias("CompMFG")]
        [String]$ComputerManufacturer
        ,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                    ValueFromPipeline=$true)]
        [Alias("CompModel")]
        [String]$ComputerModel
        ,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                    ValueFromPipeline=$true)]
        [Alias("CompSerial")]
        [String]$ComputerSerial
        ,
        
        [Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
        [ValidateSet("Laptop","Desktop","Virtual","Server","1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24")]
        #[ValidateRange(1,23)]
        [Alias("Chassis")]
        [String]$ChassisType
        ,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                    ValueFromPipeline=$true)]
        [String]$LastUser
        ,

        [Parameter(ValueFromPipelineByPropertyName=$true,
                    ValueFromPipeline=$true)]
        [String]$TopUser
        ,
        
        [Parameter(ValueFromPipelineByPropertyName=$true,
                    ValueFromPipeline=$true)]
        [String]$Domain
        ,

        [Switch]$IncludeMultipleDisks
        ,

        [Switch]$IncludeServers

        ,

        [Switch]$IncludeDocks
               
        )#PARAM

    Process
    {
        if ($Domain)
        {
            $Domain = $Domain.ToString().Replace("*","%")
            if (-not $TopUser)
            {
                Write-Warning "TopUser must be specified when using the Domain parameter."
                return
            }
        }
        else
        {
            $Domain = $env:USERDOMAIN.ToLower()
        }
        if ($ComputerName)
        {
            #Convert incoming cast type to string so that we can replace PowerShell wildcard syntax with SQL syntax
            $ComputerName = $ComputerName.ToString().Replace('*','%')
            $ComputerNameQuery = "RV.Netbios_Name0 like '$ComputerName'"
        }
        else
        {
            $ComputerNameQuery = "RV.Netbios_Name0 like '%'"
        }

        if ($OperatingSystem)
        {
            #Convert incoming cast type to string so that we can replace PowerShell wildcard syntax with SQL syntax
            $OperatingSystem = $OperatingSystem.ToString().Replace('*','%')
            $OperatingSystemQuery = "and OS.Caption0 like '$OperatingSystem'"
        }

        if ($ComputerManufacturer)
        {
            #Convert incoming cast type to string so that we can replace PowerShell wildcard syntax with SQL syntax
            $ComputerManufacturer = $ComputerManufacturer.ToString().Replace('*','%')
            $ComputerManufacturerQuery = "and CS.Manufacturer0 like '$ComputerManufacturer'"
        }

        if ($ComputerModel)
        {
            #Convert incoming cast type to string so that we can replace PowerShell wildcard syntax with SQL syntax
            $ComputerModel = $ComputerModel.ToString().Replace('*','%')
            $ComputerModelQuery = "and (CS.Model0 like '$ComputerModel' or CP.Version0 like '$ComputerModel')"
        }

        if ($ComputerSerial)
        {
            #Convert incoming cast type to string so that we can replace PowerShell wildcard syntax with SQL syntax
            $ComputerSerial = $ComputerSerial.ToString().Replace('*','%')
            $computerSerialQuery = "and BIOS.SerialNumber0 like '$ComputerSerial'"
        }

        if ($ChassisType)
        {
            switch ($ChassisType)
            {
                "Desktop" {$ChassisTypesQuery = "and SE.ChassisTypes0 IN ('3','4','6','7','15','16','24')"}
                "Laptop" {$ChassisTypesQuery = "and SE.ChassisTypes0 IN ('8','9','10','21')"}
                "Virtual" {$ChassisTypesQuery = "and SE.ChassisTypes0 = '1'"}
                "Server" {$ChassisTypesQuery = "and SE.ChassisTypes0 IN ('17','23')"}
                default {$ChassisTypesQuery = "and SE.ChassisTypes0 = '$ChassisType'"}
            }
        }

        if ($LastUser)
        {
            #Convert incoming cast type to string so that we can replace PowerShell wildcard syntax with SQL syntax
            $LastUser = $LastUser.ToString().Replace('*','%')
            $LastUserQuery = "and RV.User_Name0 like '$LastUser'"
        }

        if ($TopUser)
        {
            #Convert incoming cast type to string so that we can replace PowerShell wildcard syntax with SQL syntax
            $TopUser = $Domain + "\" + $TopUser.ToString().Replace('*','%')
            $TopUserQuery = "and CU.TopConsoleUser0 like '$TopUser'"
        }
       
        if ($IncludeServers.IsPresent -eq $False)
        {
            $ServersFilter1 = "and RV.Operating_System_Name_and0 not like '%server%'"
        }
        
        if ($IncludeMultipleDisks)
        {
            $IncludeMultipleDisksQuery = "and LD.DeviceID0 like '%'"
        }
        else
        {
            $IncludeMultipleDisksQuery = "and LD.DeviceID0 = 'C:'"
        }

        if ($IncludeDocks)
        {
            $IncludeDocksQuery = "and SE.ChassisTypes0 like '%'"
        }
        else
        {
            $IncludeDocksQuery = "and SE.ChassisTypes0 <> '12'"
        }

                    
   
   $SqlQuery = "
                Select distinct
                RV.Netbios_Name0 ComputerName
	            ,OS.Caption0 OperatingSystem
                ,CS.SystemType0 OsBits
		        ,CS.Manufacturer0 Manufacturer
		        ,CASE
                    WHEN CS.Manufacturer0 = 'LENOVO' THEN [CP].[Version0]
                    WHEN CS.Model0 != 'LENOVO' THEN [CS].[Model0]
                    END AS Model
                ,CASE
                    WHEN SE.ChassisTypes0 = '1' THEN 'Virtual'
				    WHEN SE.ChassisTypes0 IN ('3','4','6','7','15','16','24') THEN 'Desktop'
				    WHEN SE.ChassisTypes0 IN ('8','9','10','21') THEN 'Laptop'
                    WHEN SE.ChassisTypes0 = '12' THEN 'Dock'
                    WHEN SE.ChassisTypes0 IN ('17','23') THEN 'Server'
                    ELSE SE.ChassisTypes0
				    END AS ChassisType
                ,BIOS.SerialNumber0 SerialNumber
                ,CPU.Name0 AS Processor
                ,PARSENAME(CONVERT(varchar, CAST(RAM.TotalPhysicalMemory0 /1024 AS money),1),2) + ' MB' AS InstalledMemory
                ,BIOS.SMBIOSBIOSVersion0 AS BiosVersion
                ,LD.DeviceID0 DiskID
				,LD.Description0 DiskType
				,CONVERT(varchar, CAST(LD.Size0 /1024.00 AS money), 1) + ' GB' AS DiskCapacity
				,CONVERT(varchar, CAST(LD.FreeSpace0 /1024.00 AS money), 1) + ' GB' AS DiskSpaceFree
                ,RV.User_Domain0 Domain
	            ,RV.User_Name0 LastUser	    
		        ,CU.TopConsoleUser0 TopUser
	            ,DATEDIFF(d, OS.LastBootUpTime0, GETDATE()) AS UpTimeDays
                ,WS.LastHWScan ScanTime                                

                FROM 
                v_R_System_Valid RV
	            left join v_GS_WORKSTATION_STATUS WS on RV.ResourceID=WS.ResourceID
		        left join v_GS_COMPUTER_SYSTEM CS on RV.ResourceID=CS.ResourceID
		        left join v_GS_PC_BIOS BIOS on RV.ResourceID=BIOS.ResourceID
		        left join v_GS_SYSTEM_CONSOLE_USAGE CU on RV.ResourceID=CU.ResourceID
                left join v_GS_OPERATING_SYSTEM OS on RV.ResourceID=OS.ResourceID
                left join v_GS_COMPUTER_SYSTEM_PRODUCT CP on RV.ResourceID=CP.ResourceID
                left join v_GS_SYSTEM_ENCLOSURE SE on RV.ResourceID=SE.ResourceID
                left join v_GS_LOGICAL_DISK LD on RV.ResourceID=LD.ResourceID
                left join v_GS_X86_PC_MEMORY RAM on RV.ResourceID=RAM.ResourceID
				left join v_GS_PROCESSOR CPU on RV.ResourceID=CPU.ResourceID

                WHERE
	            $ComputerNameQuery
                $OperatingSystemQuery
                $ComputerManufacturerQuery
                $ComputerModelQuery
                $ComputerSerialQuery
                $LastUserQuery
                $TopUserQuery
                $ChassisTypesQuery
                $ServersFilter1
                $IncludeMultipleDisksQuery
                $IncludeDocksQuery
                "
    $Global:Get_SCCMComputerFunctionQuery = $SqlQuery
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$SqlConnection.ConnectionString = 'Server="tcp:ALCRMCWSSP01.pmusa.net,1433"; Database="CM_C01"' #;Integrated Security=SSPI'
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
	$SqlCmd.CommandText = $SqlQuery
	$SqlCmd.Connection = $SqlConnection
    $SqlCmd.CommandTimeout = 0
	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$SqlAdapter.SelectCommand = $SqlCmd
	$DataSet = New-Object System.Data.DataSet
	$SqlAdapter.Fill($DataSet) | Out-Null
	$SqlConnection.Close()
    return $DataSet.Tables[0]

    }
}