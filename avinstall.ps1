$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop

Start-Sleep -Seconds 5

$hostname=hostname.exe
$ipadd= ipconfig.exe | findstr /i "ipv4"
$username= whoami
$produttore= Get-WMIObject -class Win32_ComputerSystem -Recurse | select Manufacturer 
$Modello= Get-WMIObject -class Win32_ComputerSystem -Recurse | select Model 
$serial= Get-WMIObject -class Win32_bios | select SerialNumber 
$bitloker= Get-BitLockerVolume  | select -ExpandProperty Protectionstatus

if (Get-Service AVP -ErrorAction SilentlyContinue){
    $av=1
}elseif (Get-Service KAVFS -ErrorAction SilentlyContinue){
    $av=1
}else{
    $av=0
    $wshell.Popup("Antivirus non presente verificare $hostname",0,"Antivirus")
}

START "https://dev.facile.it/info_pc.php?hostname=$hostname&ip=$ipadd&user=$username&brand=$produttore&model=$modello&serial=$serial&bitlocker=$bitloker&antivirus=$av"
