# Script to unmount Veeam vPower NFS datastore
 
# Fill in the information below
# vCenter server address (FQDN or IP)
$vcenter = "192.168.1.175"
# vCenter Username
$user = "FOONET\Administrator"
# Password
$pass = "XXXX"

# DO NOT TOUCH BELOW!!
 
# Connect to vCenter
Connect-VIServer -Server $vcenter -Username $user -Password $pass | Out-Null

$hosts = Get-VMHost
foreach ($VMHost in $hosts) {
    $veeamshare = Get-Datastore | where {$_.type -eq "NFS" -and $_.name -Match "VeeamBackup_*"} 
    Remove-Datastore -VMHost $VMHost -Datastore $veeamshare -confirm:$false
}
 
# Disconnect from vCenter
Disconnect-VIServer -Server $vcenter -Confirm:$false | Out-Null