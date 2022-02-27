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
    [boolean]$ListGroups   = $true,
    [boolean]$ListFacStaff = $true,
    [boolean]$ListStudents = $true,
    [boolean]$Recursive = $false
)

$title = 'Enter Group'
$msg   = 'Enter Group to List Members:'

Add-Type -AssemblyName Microsoft.VisualBasic
$Group = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

if (Test-path $env:Temp\GetMembers.txt) { 
    Remove-Item $env:Temp\GetMembers.txt
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

    $Groups   = $MemberObjects | Where-Object {$_.objectClass -eq "group"}
    $Students = $MemberObjects | Where-Object {$_.distinguishedName -Like "*OU=Students*"}
    $FacStaff = $MemberObjects | Where-Object {$_.distinguishedName -Like "*OU=FacultyStaff*"}

    If ($ListGroups) {
        Write-Output ""
        Write-Output "GROUPS"

        ForEach($Member in $Groups) {
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

    If ($ListFacStaff) {
        Write-Output ""
        Write-Output "FACULTY/STAFF"

        ForEach($Member in $FacStaff) {
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

$OutputFile = $env:Temp + '\GetMembers.txt'

GetMembers $Group | Format-List | Out-File $OutputFile

if (Test-Path $env:Windir\Notepad.exe){
Start-Process $env:Windir\notepad.exe $OutputFile -NoNewWindow -Wait
}

if (Test-Path $OutputFile) {
    Remove-Item $OutputFile
}