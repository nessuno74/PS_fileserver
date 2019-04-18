<# 
    .SYNOPSIS   
        Get all the groups that a user is MemberOf.

    .DESCRIPTION
        This script retrieves all the groups that a user is MemberOf in a recursive way.

    .PARAMETER SamAccountName
        The name of the user you want to check #>

Param (
    [String]$SamAccountName = 'gforzoni',
    $DomainUsersGroup = ''
)


Function Get-ADMemberOf {
    Param (
        [Parameter(ValueFromPipeline)]
        [PSObject[]]$Group,
        [String]$DomainUsersGroup = 'CN=Domain Users,CN=Users,DC=facile55,DC=local'
    )
    Process {
        foreach ($G in $Group) {
            $G | Get-ADGroup | Select -ExpandProperty Name
            Get-ADGroup $G -Properties MemberOf| Select-Object Memberof | ForEach-Object {
                Get-ADMemberOf $_.Memberof
            }
        }
    }
}


$Groups = Get-ADUser $SamAccountName -Properties MemberOf | Select-Object -ExpandProperty MemberOf
$Groups += $DomainUsersGroup
$Groups | Get-ADMemberOf | Select -Unique | Sort-Object