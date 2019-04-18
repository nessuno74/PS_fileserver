Add-pssnapin VMware.VimAutomation.Core -ErrorAction SilentlyContinue

$daysback = -30 

$filenamepath = "c:\vmIOPS.csv"

write-host "What vCenter Do you want to Connect To? - You may be prompted for a Username and Password"
write-host ""
Connect-VIServer -menu write-host ""

write-host ""
write-host "Gathering list of (powered on) Virtual Machines....."
write-host ""

$vms = Get-Folder | get-vm | where {$_.PowerState -eq "PoweredOn" }

write-host "Detected " $vms.Count " powered on virtual machines."
write-host ""
write-host "Gathering IOPS statistics for each virtual machine from the last 30 days......"
write-host ""

Add-Content $filenamepath “VM,Write IOPS Average,Read IOPS Average,Write IOPS Max, Read IOPS Max”;

$i=0;

$vms | sort | % {

    $wval = (((Get-Stat $_ -stat "datastore.numberWriteAveraged.average" -Start (Get-Date).adddays($daysback) -Finish (Get-Date) ) | select -expandproperty Value) | measure -average -max);

    $rval = (((Get-Stat $_ -stat "datastore.numberReadAveraged.average" -Start (Get-Date).adddays($daysback) -Finish (Get-Date) ) | select -expandproperty Value) | measure -average -max);

    $thisline = $_.Name + "," + $wval.average + "," + $rval.average + "," + $wval.maximum + "," + $rval.maximum;

    Add-Content c:\users\gforzoni\desktop\vm\vm.csv $thisline;

    $i++

    write-host $i $_.Name 

}

write-host "" write-host "Completed!"