$Groups = Get-ADGroup -Filter * -SearchBase 'OU=root,DC=facile55,DC=local'

$Results = foreach( $Group in $Groups ){

    Get-ADGroupMember -Identity $Group |  where {$_.objectClass -eq "user"} | foreach {

        [pscustomobject]@{

            GroupName = $Group.Name

            Name = $_.Name

            }

        }

    }

$Results| Export-Csv -Path c:\data\groups1.csv -NoTypeInformation﻿
