<#
.Synopsis
Retrieves the list of members for provided AD group and opens notepad with results
.Description
Retrieves the list of members for provided Active Directory group and opens notepad with results
.Parameter Group
Group Name
.Example
GetMembersFile School of Education
.Example
GetMembersFile Counseling Center
.Example
GetMembersFile fs^files^facman$^Warehouse_w
#>
Param(
	[Parameter(Mandatory=$False,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
	[string]$Group,
    [boolean]$ListAffliliate = $true,
    [boolean]$ListAlumni = $true,
    [boolean]$ListEmployee = $true,
    [boolean]$ListExchange = $true,
    [boolean]$ListFacStaff = $true,
    [boolean]$ListGroups  = $true,
    [boolean]$ListRobots = $true,
    [boolean]$ListStudents = $true,
    [boolean]$ListPeople = $true,
    [boolean]$ListServerAccounts = $true,
    [boolean]$Recursive = $false
)


$title = 'Enter Group'
$msg   = 'Enter Group to List Members:'

Add-Type -AssemblyName Microsoft.VisualBasic 
$Group = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

if (Test-path $env:Temp\GetAllMembers.txt) { 
    Remove-Item $env:Temp\GetAllMembers.txt
}

Import-Module ActiveDirectory

Function GetMembers {
    Param(
	    [Parameter(Mandatory=$True)]
	    [string]$GroupInternal
    )

    Write-Output "`n"
    Write-Output "Membership for `"$GroupInternal`""


    $Members = Get-ADGroupMember -Identity $GroupInternal

    $MemberObjects = ForEach ($Member in $Members) { Get-ADObject -Identity $Member -Properties * }
    
    $Groups  =        $MemberObjects | Where-Object {$_.objectClass -eq "Group"}
    $Affliliates =    $MemberObjects | Where-Object {$_.distinguishedName -Like "*Affliliates*"}
    $Alumni =         $MemberObjects | Where-Object {$_.distinguishedName -Like "*Alumni*"}
    $Employee =       $MemberObjects | Where-Object {$_.distinguishedName -Like "*Employee*"}
    $Exchange =       $MemberObjects | Where-Object {$_.distinguishedName -Like "*Exchange*"} 
    $FacStaff =       $MemberObjects | Where-Object {$_.distinguishedName -Like "*FacultyStaff*"}
    $People =         $MemberObjects | Where-Object {$_.distinguishedName -Like "*People*"}
    $Robots =         $MemberObjects | Where-Object {$_.distinguishedName -Like "*Robots*"}
    $ServerAccounts = $MemberObjects | Where-Object {$_.distinguishedName -Like "*ServerAccounts*"}
    $Students =       $MemberObjects | Where-Object {$_.distinguishedName -Like "*Students*"}
    

    If ($ListGroups) {
        Write-Output ""
        Write-Output "GROUPS"

        ForEach($Member in $Groups) {
            $stringitem = -join("     ",$Member.SamAccountName)
            $stringitem = -join($stringitem,":",$Member.DisplayName)
            $stringitem
        }
    }
   
   If ($ListAffliliates ) {
        Write-Output ""
        Write-Output "AFFILIATES"

        ForEach($Member in $Affliliates ) {
            $stringitem = -join("     ",$Member.SamAccountName)
            $stringitem = -join($stringitem,":",$Member.DisplayName)
            $stringitem
        }
    }

    If ($ListAlumni) {
        Write-Output ""
        Write-Output "ALUMNI"

        ForEach($Member in $Alumni) {
            $stringitem = -join("     ",$Member.SamAccountName)
            $stringitem = -join($stringitem,":",$Member.DisplayName)
            $stringitem
        }
    }
       
    If ($ListEmployee) {
        Write-Output ""
        Write-Output "EMPLOYEE"

        ForEach($Member in $Employee) {
            $stringitem = -join("     ",$Member.SamAccountName)
            $stringitem = -join($stringitem,":",$Member.DisplayName)
            $stringitem
        }
    }

     If ($ListExchange) {
        Write-Output ""
        Write-Output "EXCHANGE"

        ForEach($Member in $Exchange) {
            $stringitem = -join("     ",$Member.SamAccountName)
            $stringitem = -join($stringitem,":",$Member.DisplayName)
            $stringitem
        }
    }
    
    If ($ListFacStaff) {
        Write-Output ""
        Write-Output "FACULTY/STAFF"

        ForEach($Member in $FacStaff) {
            $stringitem = -join("     ",$Member.SamAccountName)
            $stringitem = -join($stringitem,":",$Member.DisplayName)
            $stringitem
        }
    }

    If ($ListPeople) {
        Write-Output ""
        Write-Output "PEOPLE"

        ForEach($Member in $people) {
            $stringitem = -join("     ",$Member.SamAccountName)
            $stringitem = -join($stringitem,":",$Member.DisplayName)
            $stringitem
        }
    }

    If ($ListRobots) {
        Write-Output ""
        Write-Output "ROBOTS"

        ForEach($Member in $Robots) {
            $stringitem = -join("     ",$Member.SamAccountName)
            $stringitem = -join($stringitem,":",$Member.DisplayName)
            $stringitem
        }
    }

    If ($ListServerAccounts) {
        Write-Output ""
        Write-Output "SERVER ACCOUNTS"

        ForEach($Member in $ServerAccounts) {
            $stringitem = -join("     ",$Member.SamAccountName)
            $stringitem = -join($stringitem,":",$Member.DisplayName)
            $stringitem
        }
    }
    
    If ($ListStudents) {
        Write-Output ""
        Write-Output "STUDENTS"

        ForEach($Member in $Students) {
            $stringitem = -join("     ",$Member.SamAccountName)
            $stringitem = -join($stringitem,":",$Member.DisplayName)
            $stringitem
        }
    }

    If ($Recursive) {
        ForEach ($Group in $Groups) {
            Write-Output "`n"
            Write-Output "--------"
            $scratch = $Group.SamAccountName
            GetMembers -GroupInternal $scratch
        }
    }
}

# Section added to present output to notepad for copy and paste

$OutputFile = $env:Temp + '\GetAllMembers.txt'

GetMembers $Group | Format-List | Out-File $OutputFile

if (Test-Path $env:Windir\Notepad.exe){
Start-Process $env:Windir\notepad.exe $OutputFile -NoNewWindow -Wait
}

if (Test-Path $OutputFile) {
    Remove-Item $OutputFile
}