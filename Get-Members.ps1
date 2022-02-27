Param(
	[Parameter(Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
	[string]$Group,
    [boolean]$ListGroups   = $true,
    [boolean]$ListUsers = $true,
    [boolean]$ListITP = $true,
    [boolean]$Recursive = $false
)

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

    $Groups = $MemberObjects | Where-Object {$_.objectClass -eq "user"}
    $Users = $MemberObjects | Where-Object {$_.distinguishedName -Like "*OU=Users*"}
    $ITP = $MemberObjects | Where-Object {$_.distinguishedName -Like "*OU=ITP*"}

    If ($ListGroups) {
        Write-Output ""
        Write-Output "GROUPS"

        ForEach($Member in $Groups) {
            $stringitem = -join("     ",$Member.SamAccountName)
            $stringitem = -join($stringitem,":",$Member.DisplayName)
            $stringitem
        }
    }

    If ($ListUsers) {
        Write-Output ""
        Write-Output "USERS"

        ForEach($Member in $Users) {
            $stringitem = -join("     ",$Member.SamAccountName)
            $stringitem = -join($stringitem,":",$Member.DisplayName)
            $stringitem
        }
    }

    If ($ListITP) {
        Write-Output ""
        Write-Output "ITP"

        ForEach($Member in $ITP) {
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

GetMembers $Group