#Import-Module ActiveDirectory
#Get-ADUser -Filter * -SearchBase 'OU=IT,OU=Utenti,DC=facile55,DC=local' #-Properties emailAddress
#| `
 #   ForEach-Object { Set-ADUser -EmailAddress ($_.givenName + '.' + $_.surname + '@test.net') -Identity $_ }



 #Import-Module ActiveDirectory
#Get-ADUser -Filter * -SearchBase 'OU=IT,OU=Utenti,DC=facile55,DC=local'| `
 # ForEach-Object { Set-ADUser -EmailAddress ($_.givenName + '.' + $_.surname + '@facile.it').ToLower() -Identity $_ }


#Get-ADUser -Filter {$_.givenName -eq 'givenname'} -SearchBase 'OU=IT,OU=Utenti,DC=facile55,DC=local' | Select-Object GivenName


#$nomi= Get-ADUser -Filter * -SearchBase 'OU=IT,OU=Utenti,DC=facile55,DC=local' | 
#Select-Object -ExpandProperty givenName 
#$cognomi= Get-ADUser -Filter * -SearchBase 'OU=IT,OU=Utenti,DC=facile55,DC=local' | 
#Select-Object -ExpandProperty surname 

#$nomi="$nomi".ToLower()
#$cognomi="$cognomi".ToLower()

#ForEach-Object { write-host ($nomi + '.' + $cognomi + '@facile.it') }


#Get-ADUser -Filter * -SearchBase 'OU=IT,OU=Utenti,DC=facile55,DC=local'| `
#ForEach-Object { Set-ADUser -EmailAdd ($_.givenName + '.' + $_.surname + '@facile.it').ToLower() -Identity $_ }

Get-ADUser -Filter * -SearchBase 'OU=IT,OU=Utenti,DC=facile55,DC=local'| `
ForEach-Object {Set-ADUser -Identity $_.samaccountname -samaccountname $_.samaccountname.ToLower()}

