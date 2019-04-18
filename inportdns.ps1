$lines = Import-Csv C:\Users\gforzoni\Desktop\dns_record.csv



foreach ($data in $lines)
{
    $name = $data.'name'

    $domain = $data.'domain'

    $alias = $data.'alias'
    
    #write-host $name,$domain,$alias
    Add-DnsServerResourceRecordCName -ComputerName SRVVMI0076.facile55.local -HostNameAlias $alias -Name $name -ZoneName $domain

}



