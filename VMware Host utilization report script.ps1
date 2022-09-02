<#

.Synopsis
 This Script helps user to get the ESXi hosts Utilization report.

.DESCRIPTION
 The Script will perform checks and provide the cluster report such as powered on and off of VM's and Snapshot count and Host Utilization report on each of the ESXi hosts. This report will contain vSAN storage capacity of cluster, 
 CPU and Memory capacity of ESXi Host.

.INPUT
    Please note: All the files to be updated whenever you make changes.
 1)Input_DisplayNames_HostPerformance.csv - CSV file containing below fields.
    Vcenter,Cluster,Organization
 2)credfile.csv- CSV file containing below fields
    Vcentername,user,pwd - Kindly update the vCenter, user, pwd field everytime you make changes to the other two input files.
 3)Input_DisplayNames_vCenterClusterStatus.csv - CSV file containing below fields
    vCenter, DisplayName, Cluster, Displayname, Datastore - Kindly note: Please update all the fields in order when you modify it.
        
.OUTPUT
 The report will be generated in report.html file stored in the powershell solution path. Also, it will send the report in email to the mentioned Owners.

.Work Instructions
 1) If the password of the vcenter is changed, then just add the details of the vCenter along with the new password directly in the credfile.csv.
 2) If another vCenter detail is to be added for execution, then add all the details of the required vCenters in the credfile.csv directly.
 3) If the recepients or CC recepients has to be altered or changed, then alter or change it in line 511 and 512.

.NOTES
 Version:        1.0.0
 Author :        Rajesh. R 
 Author email:   <rajesh.r@dxc.com>
 Creation Date:  11/08/2022
 Purpose/Change: Initial script development

 #>

#------------------------------------------------------Declarations-------------------------------------------------------------#
$path="$((Get-Location).Path)"
Start-Transcript -Path "$path\TranScript_$date.txt"
$date = "{0,10:dd.MM.yyyy}" -f $(Get-Date)
$Input_BackupLocation = Import-Csv "$path\Input_BackupLocation.csv"
$date_backup = "{0:yyyyMMdd}" -f (get-date).AddHours(-24)
$html_report="$path\Alstom_healthcheck.html"
$errlog="$path\ErrorLog.csv"
$CredFilePath = [string]$path + "\credfile.csv"
$enscript = [string]$path + "\enscript.txt"
$credentials = Get-Content $CredFilePath | select -Skip 1 | ConvertFrom-Csv -Delimiter "," -Header ("vCenter","User","pwd")
$path_VMsInfo = "$path\VMsInfo.csv"

#--------------------------------------------------------Extraction---------------------------------------------------------------#

# HTML Table 1 - vCenter, Cluster, DisplayNames and status
$Input_DisplayNamesTB1 = "$($path)\Input_DisplayNames_vCenterClusterStatus.csv" # InputFile defined here
$Input_Table1 = gc $Input_DisplayNamesTB1 | select -Skip 1 | ConvertFrom-Csv -Delimiter "," -Header ("vCenter","vCenterDisplayName","Cluster","ClusterDisplayName","Datastore")
$vcnameTB1 = $Input_Table1 | Group-Object -Property vCenter -AsHashTable -AsString
$vcnameTBProp = $Input_Table1 | Group-Object -Property vCenter
$vCenterNameTB1 = $Input_Table1 | Group-Object -Property vCenter -AsHashTable -AsString

# HTML Table 2 - Host Performance
$Input_DisplayNamesTB2 = "$($path)\Input_DisplayNames_HostPerformance.csv" # InputFile defined here
$Input_Table2 = gc $Input_DisplayNamesTB2 | select -Skip 1 | ConvertFrom-Csv -Delimiter "," -Header ("vCenter","Cluster","Hostname")
$vcnameTB2 = $Input_Table2 | Group-Object -Property vCenter
$vCenterNameTB2 = $Input_Table2 | Group-Object -Property vCenter -AsHashTable -AsString
$ClusterNameTB2 = $Input_Table2 | Group-Object -Property Cluster
$ClusterTB2 = $Input_Table2 | Group-Object -Property Cluster -AsHashTable -AsString

#-----------------------------------------------------Removal of old files---------------------------------------------------#

if(Test-Path $html_report)
{
Write-Host "`tPrevious log report exists" 
write-host "`tDeleting previous log report" -ForegroundColor Green
Remove-Item $html_report -Confirm:$false -ErrorAction SilentlyContinue
}

if(Test-Path $path_VMsInfo)
{
Write-Host "`tPrevious log report exists" 
write-host "`tDeleting previous log report" -ForegroundColor Green
Remove-Item $path_VMsInfo -Confirm:$false -ErrorAction SilentlyContinue
}

Write-Host "`t------------------------------------------------------------`n"
if(Test-Path $errlog)
{
Write-Host "`tPrevious error log report exists" 
write-host "`tDeleting previous log report" -ForegroundColor Green
Remove-Item $errlog -Confirm:$false -ErrorAction SilentlyContinue
}
Write-Host "`t------------------------------------------------------------`n"
if ($credentials -ne $null)
{
if(Test-path $enscript)
{
Write-Host "`t Removing previous enscripted file" -ForegroundColor Green
Remove-Item  $enscript -Confirm:$false -ErrorAction SilentlyContinue
}
}

#------------------------------------------------------Global variables-------------------------------------------------#

$htmlbody1 = @()
$htmlbody  = @()
$finalformat= @()
$hostname1 = @()
$hostcheck = @()
$CPUMax    = @()
$CPUAvg    = @()
$CPUMin    = @()
$MemoryMax = @()
$MemoryAvg = @()
$MemoryMin = @()
$hosts1    = @()
$result    = @()

#-------------------------------------------- vmHost performance function - Table 2---------------------------------------#

Function HostPerformance($cluster) {
$allhosts = @()
foreach ($onec in $cluster) {
$cluster = Get-Cluster -Name $onec.Name
$hosts = $cluster | Get-VMHost
foreach($vmHost in $hosts){
      
  $vmHost = $vmHost.Name
  Write-Host "`tCalculating the CPU and memory utilization of ESXi Host- $vmHost" -ForegroundColor Green
  $hoststat = "" | Select HostName,OverallStatus, Cluster, MemoryInstalled, MemoryAllocated, MemoryConsumed, MemoryAvailable, CPUMax, CPUAvg, CPUMin, MemoryMax, MemoryAvg, MemoryMin
  $hoststat.HostName = $vmHost
  $hoststat.Cluster = $onec.Name

  try{
  $vmcheck = Get-VMHost -name $vmHost -ErrorAction Stop
  $statcpu = Get-Stat -Entity ($vmHost)-start (get-date).AddHours(-24) -IntervalMins 5 -MaxSamples 288 -stat cpu.usage.average | Measure-Object -Property value -Average -Maximum -Minimum 
  $statmemconsumed = Get-Stat -Entity ($vmHost)-start (get-date).AddHours(-24) -IntervalMins 5 -MaxSamples 288 -stat mem.consumed.average
  $statmemusage = Get-Stat -Entity ($vmHost)-start (get-date).AddHours(-24) -IntervalMins 5 -MaxSamples 288 -stat mem.usage.average
  $statmemallocated = Get-VMhost $vmHost | Select @{N="allocated";E={$_ | Get-VM | %{$_.MemoryGB} | Measure-Object -Sum | Select -ExpandProperty Sum}}
  $statmeminstalled = Get-VMHost $vmHost | select MemoryTotalGB
  $statmeminstalled = $statmeminstalled.MemoryTotalGB
    
  $memconsumed = $statmemconsumed | Measure-Object -Property value -Average
  $vm = get-vmHost $vmHost 
  $memusage = $vm | Get-Stat -Stat mem.usage.average -Start (Get-Date).AddHours(-24)-IntervalMins 5 -MaxSamples (288) |Measure-Object Value -Average -Maximum -Minimum

  $CPUMax = [Math]::Round((($statcpu).Maximum),2)
  $CPUAvg = [Math]::Round((($statcpu).Average),2)
  $CPUMin = [Math]::Round((($statcpu).Minimum),2)
  $MEMMax = [Math]::Round((($memusage).Maximum),2)
  $MEMAvg = [Math]::Round((($memusage).Average),2)
  $MEMMin = [Math]::Round((($memusage).Minimum),2)
  $allocated = "{0:N0}" -f ($statmemallocated.allocated)
  $consumed1 = "{0:N0}" -f ($memconsumed.Average/1024/1024)
  $consumed2 = ($memconsumed.Average/1024/1024)
  $consumed_cal = (( $consumed2*100)/$statmeminstalled)
  [int]$consumed = "{0:N0}" -f $consumed_cal
  
  $usage = "{0:P0}" -f ($memusage.Average/100)
  $installed = "{0:N0}" -f ($statmeminstalled)  

    # Second table: Colouring conditions
    if ($CPUMax -lt 60) { $CPUMax_status = "<td class=""Redfont""; style=""background-color:#6cc24a"">$($CPUMax) %" }
    elseif ($CPUMax -ge 60 -and $CPUMax -lt 80 )  { $CPUMax_status = "<td class=""Redfont""; style=""background-color:#ffa500"">$($CPUMax) %" }
    elseif ($CPUMax -ge 80) { $CPUMax_status = "<td class=""Redfont""; style=""background-color:#ff0000"">$($CPUMax) %</td>" }

    if ($CPUAvg -lt 70) { $CPUAvg_status = "<td class=""Redfont""; style=""background-color:#6cc24a"">$($CPUAvg) % </td>" }
    elseif ($CPUAvg -ge 70 -and $CPUAvg -lt 80 )  { $CPUAvg_status = "<td class=""Redfont""; style=""background-color:#ffa500"">$($CPUAvg) %</td>" }
    elseif ($CPUAvg -ge 80) { $CPUAvg_status = "<td class=""Redfont""; style=""background-color:#ff0000"">$($CPUAvg) %</td>" }
    
    if ($CPUMin -lt 60) { $CPUMin_status = "<td class=""Redfont""; style=""background-color:#6cc24a"">$($CPUMin) % </td>" }
    elseif ($CPUMin -ge 60 -and $CPUMin -lt 65 )  { $CPUMin_status = "<td class=""Redfont""; style=""background-color:#ffa500"">$($CPUMin) %</td>" }
    elseif ($CPUMin -gt 80) { $CPUMin_status = "<td class=""Redfont""; style=""background-color:#ff0000"">$($CPUMin) %</td>" }

    if ($MEMMax -lt 70) { $MEMMax_status = "<td class=""Redfont""; style=""background-color:#6cc24a"">$($MEMMax) % </td>" }
    elseif ($MEMMax -ge 70 -and $MEMMax -lt 80 )  { $MEMMax_status = "<td class=""Redfont""; style=""background-color:#ffa500"">$($MEMMax) %</td>" }
    elseif ($MEMMax -ge 80) { $MEMMax_status = "<td class=""Redfont""; style=""background-color:#ff0000"">$($MEMMax) %</td>" }

    if ($MEMAvg -lt 70) { $MEMAvg_status = "<td class=""Redfont""; style=""background-color:#6cc24a"">$($MEMAvg) % </td>" }
    elseif ($MEMAvg -ge 70 -and $MEMAvg -lt 80 )  { $MEMAvg_status = "<td class=""Redfont""; style=""background-color:#ffa500"">$($MEMAvg) %</td>" }
    elseif ($MEMAvg -ge 80) { $MEMAvg_status = "<td class=""Redfont""; style=""background-color:#ff0000"">$($MEMAvg) %</td>" }

    if ($MEMMin -lt 70) { $MEMMin_status = "<td class=""Redfont""; style=""background-color:#6cc24a"">$($MEMMin) % </td>" }
    elseif ($MEMMin -ge 70 -and $MEMMin -lt 80 )  { $MEMMin_status = "<td class=""Redfont""; style=""background-color:#ffa500"">$($MEMMin) %</td>" }
    elseif ($MEMMin -ge 80) { $MEMMin_status = "<td class=""Redfont""; style=""background-color:#ff0000"">$($MEMMin) %</td>" }
  
    $MemoryInstalled = $installed.ToString() + " GB"
    $MemoryAllocated = $allocated.ToString() + " GB"

    $esxView = Get-VMHost -name $vmHost | Get-View
    If ($esxView.OverallStatus -like "Red"){
        $OverallStatus = "<td class=""Redfont""; style=""background-color:#e30000"";>Critical</td>"
    }elseif($esxView.OverallStatus -like "yellow"){
        $OverallStatus = "<td class=""Redfont""; style=""background-color:#ffa500"";>Warning</td>"
    }else{
        $OverallStatus = "<td class=""Redfont""; style=""background-color:#6cc24a"";>Normal</td>"
    }
  
    $hoststat.CPUMax = $CPUMax_status
    $hoststat.CPUAvg = $CPUAvg_status
    $hoststat.CPUMin = $CPUMin_status
    $hoststat.MemoryMax = $MEMMax_status
    $hoststat.MemoryAvg = $MEMAvg_status
    $hoststat.MemoryMin = $MEMMin_status
    $hoststat.MemoryInstalled = $MemoryInstalled
    $hoststat.MemoryAllocated = $MemoryAllocated
    $hoststat.MemoryConsumed = $MemoryConsumed
    $hoststat.MemoryAvailable = $MemoryUsage
    $hoststat.OverallStatus = $OverallStatus
    $allhosts += $hoststat

    
}
  catch [Exception]{
    Write-Host "`tError while execution. Unable to find VMHost- $vmHost"
}
}
}
$allhosts | Select HostName, OverallStatus, Cluster, CPUMax, CPUAvg, CPUMin, MemoryMax, MemoryAvg, MemoryMin | Sort-Object HostName 
}

#--------------------------------------------------------Execution--------------------------------------------------------#

if (!(Test-Path $enscript))
{
if(!(Test-path $CredFilePath))
{
write-host "`tThe input credential file is unavailable" -ForegroundColor Red
exit
} 
$b=0

foreach($a in $credentials)
{
if(Test-Path "$path/key$b.txt")
{
Remove-Item "$path/key$b.txt"
}
$key = New-Object Byte[] 32
$a.vCenter >> $enscript
$a.User >> $enscript
$passwd = ConvertTo-SecureString $a.pwd -AsPlainText -Force
$securestringtext = $passwd | ConvertFrom-SecureString -Key $key >> $enscript
$keypath = [string]$Path + "\key$b.txt"
$key >> $keypath
$b++
}
(Get-Content $CredFilePath |  Select -First 1) | Out-File $CredFilePath
}

$z =0
$x = Get-Content $enscript
for($u=0;$u -lt $x.Length;$u++)
{
$keypath = [string]$Path + "\key$z.txt"
$key=Get-content $keypath
$vc = $x[$u]
$u++
$VCuser = $x[$u]
$u++
$password = $x[$u] | ConvertTo-SecureString -Key $key
$creds = New-Object System.Management.Automation.PSCredential($VCuser,$password)
$z++

try{    


Write-host "Connecting vCenter server....."
Connect-VIServer -server $vc -Credential $creds | Out-Null
write-host "`tSuccessfully connected to the $vc vCenter" -ForegroundColor Green

foreach($m in $vcnameTB2)
{
if($m.name -eq $vc)
{
$finalformat = @()
$clusterlist = $vCenterNameTB2.$vc
$cls = $clusterlist | Group-Object -Property cluster
$h = $clusterlist | Group-Object -Property hostname
# Host Performance Function
$perf = HostPerformance($cls)
$perf_report = $perf | Group-Object -Property cluster -AsHashTable -AsString
foreach($i in $cls)
{
    $htmlbody=@()
    $cluster_name=$i.Name
    Write-host "`tRetrieving values for cluster $cluster_name" -ForegroundColor Green
    foreach($s in $vcnameTBProp)
    {
        if ($s.name -eq $m.Name) {
            $clusterlist2 = $vcnameTB1.$vc
            foreach($t in $clusterlist2)
            {
                if($t.Cluster -eq $cluster_name){
                    $Vsan_SpaceUsage=Get-Datastore -name $t.Datastore -RelatedObject (Get-Cluster -Name $t.Cluster)
                    $free_capacity = [math]::round( ($Vsan_SpaceUsage.FreeSpaceGB * 100 ) / $Vsan_SpaceUsage.CapacityGB,2)
                    $free_cap = 100-($Vsan_SpaceUsage.FreeSpaceGB * 100 ) / $Vsan_SpaceUsage.CapacityGB
                    $utilized_percent = [math]::Ceiling($free_cap)
    
                    if ($utilized_percent -lt 70) { $Storpercentage_status = "<td class=""Redfont""; style=""background-color:#6cc24a"";>$($utilized_percent) % </td>" }
                    elseif ($utilized_percent -ge 70 -and $utilized_percent -lt 80 )  { $Storpercentage_status = "<td style=""background-color:#ffa500"";>$($utilized_percent)%</td>" }
                    elseif ($utilized_percent -ge 80) { $Storpercentage_status = "<td class=""Redfont""; style=""background-color:#ff0000""; >$($utilized_percent) %</td>" }


                    # vSan Health status
                    try {
                        $vSanHealth_status = "<td></td>"
                        $vsanViewClusterHealthSystem = Get-VSANView -Id "VsanVcClusterHealthSystem-vsan-cluster-health-system"
                        $data = $vsanViewClusterHealthSystem.VsanQueryVcClusterHealthSummary((Get-Cluster -Name $t.Cluster).id,$null,$null,$false,$null, $true,[vmware.vsan.views.vsanhealthperspective]::defaultView)
                        $vSanHealth = $data.OverallHealth
                        if ($vSanHealth -ieq "green") { $vSanHealth_status = "<td class='Redfont'; style='background-color:#6cc24a';>$($vSanHealth)</td>"}
                        elseif ($vSanHealth -ieq "yellow" )  { $vSanHealth_status = "<td class='Redfont'; style='background-color:#ffa500';>$($vSanHealth)</td>"}
                        elseif ($vSanHealth -ieq "red") { $vSanHealth_status = "<td class='Redfont'; style='background-color:#e30000';>$($vSanHealth)</td>"}
                        else {$vSanHealth_status = "<td>$($vSanHealth)</td>"}
                    }
                    catch [Exception] {
                        write-host "Unable to capture vsan health status for cluster $($t.Cluster)"
                        $vSanHealth_status = "<td></td>"
                    }


                    # Powered-Off & Powered-On of VM: Count
                    $count_poweredON = (Get-cluster -Name $t.Cluster | Get-VMHost | Get-VM | Where-Object {$_.powerstate -eq "PoweredOn"}).count
                    $count_poweredOFF = (Get-cluster -Name $t.Cluster | Get-VMHost | Get-VM | Where-Object {$_.powerstate -eq "PoweredOff"}).count
                    $count_poweredON_name = (Get-cluster -Name $t.Cluster | Get-VMHost | Get-VM | Where-Object {$_.powerstate -eq "PoweredOn"}).Name
                    $count_poweredOFF_name = (Get-cluster -Name $t.Cluster | Get-VMHost | Get-VM | Where-Object {$_.powerstate -eq "PoweredOff"}).Name
                    Write-Host $count_poweredON, $count_poweredOFF

                    # Snapshots: Count
                    $count_snapshot = (Get-cluster -Name $t.Cluster | Get-VMHost | Get-VM | Get-Snapshot | where {$_.Created -lt (Get-Date).AddDays(-3)} | Select-Object -Property VM -Unique).count
                    $count_snapshot_name = (Get-cluster -Name $t.Cluster | Get-VMHost | Get-VM | Get-Snapshot | where {$_.Created -lt (Get-Date).AddDays(-3)} | Select-Object -Property VM -Unique).VM
                    Write-Host $count_snapshot
                    
                    break }
             }
             break }
         
         } 

        # Exporting details to csv:        
        New-Object -TypeName PSCustomObject -Property @{
        vCenter = $m.Name
        Cluster_Name= $t.Cluster
        Storage_percentage = $utilized_percent
        Powered_On = $count_poweredON_name -join ";"
        Powered_Off = $count_poweredOFF_name -join ";"
        Snapshot_Count_greater_3days = $count_snapshot_name.Name -join ";"
        } | Select vCenter,Cluster_Name,Storage_percentage,Snapshot_Count_greater_3days,Powered_On,Powered_Off  | Export-Csv -Path $path_VMsInfo -NoTypeInformation -Append -Force


    # vCenter, Cluster, Storage, vSan Health, PoweredOn, PoweredOff and SnapShots html

    $mainTBrow1 = ""
    $mainTBrow1 += "<table><th> vCenter Name </th><th> Cluster Name </th><th>Cluster Report</th><th>Results</th>"
    $mainTBrow1 += "<tr><td rowspan='4';> $($t.vCenterDisplayName) </td><td rowspan='4';> $($t.ClusterDisplayName) </td><td>Storage Capacity Usage</td>$($Storpercentage_status)</tr><tr><td> vSan Health Status </td> $($vSanHealth_status.ToUpper()) </tr>"
    $mainTBrow1 += "<tr><td>VM Power Status<td>ON: $($count_poweredON) - OFF: $($count_poweredOFF)</td></tr>"
    $mainTBrow1 += "<tr><td> Snapshots > 3 days</td><td> $($count_snapshot) </td></tr></table><br>"
    
    # Host Performance html
    $data=@()
    $data=$perf_report.$cluster_name
    $final= $data | select hostname, OverallStatus, CPUMax, CPUAvg, CPUMin, MemoryMax, MemoryAvg, MemoryMin
    $hostname1 = ""
    $hostcheck = ""
    $CPUMax    = ""
    $CPUAvg    = ""
    $CPUMin    = ""
    $MemoryMax = ""
    $MemoryAvg = ""
    $MemoryMin = ""
    $OverallStat = ""
    foreach($row in $final){
        $hostname1   += "<tr><td>$($row.HostName)</td></tr>"
        $OverallStat += "<tr>$($row.OverallStatus)</tr>"
        $CPUMin      += "<tr>$($row.CPUMin)</tr>"
        $CPUAvg      += "<tr>$($row.CPUAvg)</tr>"
        $CPUMax      += "<tr>$($row.CPUMax)</tr>"
        $MemoryMin   += "<tr>$($row.MemoryMin)</tr>"
        $MemoryAvg   += "<tr>$($row.MemoryAvg)</tr>"
        $MemoryMax   += "<tr>$($row.MemoryMax)</tr>"
        $hostcheck   += $row.HostName
       }
    if($final.HostName -eq $null)
    {continue}
    $rowdata1 = "
        <tr>
        <td><table>$($hostname1)</table></td>
        <td><table style='table-layout:fixed;width: 100%'>$($OverallStat)</table></td>
        <td><table style='table-layout:fixed;width: 100%'>$($CPUMin)</table></td>
        <td><table style='table-layout:fixed;width: 100%'>$($CPUAvg)</table></td>
        <td><table style='table-layout:fixed;width: 100%'>$($CPUMax)</table></td>
        <td><table style='table-layout:fixed;width: 100%'>$($MemoryMin)</table></td>
        <td><table style='table-layout:fixed;width: 100%'>$($MemoryAvg)</table></td>
        <td><table style='table-layout:fixed;width: 100%'>$($MemoryMax)</table></td>
        </tr>"
    $htmlbody += $rowdata1
    
    $format="
    <center>
    $mainTBrow1
    </center>
    <center>
    <table>
    <th>ESXi Host</th>
    <th>Host Status</th>
    <th>CPU Min</th>
    <th>CPU Avg</th>
    <th>CPU Max</th>
    <th>Memory Min</th>
    <th>Memory Avg</th>
    <th>Memory Max</th>
    $htmlBody
    </table>
    </center>
    <br>
    <br>" 
    $finalformat +=$format  
    }
    $result += $finalformat
}
}

#----------------------------------------------Disconnect from vCenter-------------------------------------------------#

Write-Host "`tSuccessfully Report generated for vcenter $vc" -ForegroundColor Green
Disconnect-VIServer -Server $vc -Force -Confirm:$false
Write-Host "`tDisconnected from vcenter $vc" -ForegroundColor Green
}

catch [Exception] 
{
    Write-host "`tError happened during execution in $vc..Error log created" -ForegroundColor Red
    $line  = $_.InvocationInfo.ScriptLineNumber
    $Errorlog = New-Object -TypeName PsObject -Property @{                                          
		ComputerName = $comp
		Date 		 = $date
		Error      	 = $_
        Line         = $line
	    }
        $Errorlog | Export-Csv $errlog -NoTypeInformation -Append
	    Write-host "`tError log generated for $vc" -ForegroundColor Red
}
}


#-------------------------------------------------------Backup Input-------------------------------------------------------------#

Write-Host "Fetching the network path information from the inputfile.." -ForegroundColor Yellow

$backupdata = ""
$backupdata += "<table><th>vCenter Name</th><th>Last Backup File</th><th>Status</th>"

foreach($Npath in $Input_BackupLocation){    
        write-host "vCenter:" $Npath.vCenterName -fore Yellow
        Write-Host "Path:" $Npath.BackupLocation -fore Yellow
        $backupPathFile = Get-ChildItem $Npath.BackupLocation
        Write-Host $backupPathFile -ForegroundColor Green        
       
    if($(Get-ChildItem -Recurse -Path $Npath.BackupLocation -Filter * | Where-Object {$_ -match $($date_backup)})){ # The date is fetched as the folder name since its updated with date.
        $condition = $(Get-ChildItem -Recurse -Path $Npath.BackupLocation -Filter * | Where-Object {$_ -match $($date_backup)}).Name
                       Write-Host -ForegroundColor Green "Backup Directory(s) Found!" $condition
            
                if ($condition -match $date_backup){
                    $availability = "<td class=""Redfont""; style=""background-color:#6cc24a"";>Lastest Backup Available</td>"
                    }
                else {
                    $availability = "<td class=""Redfont""; style=""background-color:#ffa500"";>Latest Backup Not Available</td>"
                    }
        }

    else{
        write-host -ForegroundColor Red "Error occured! Check input argument(s)"
        }

    $backupdata += "<tr><td>$($Npath.vCenterName)</td><td>$($condition)</td>$($availability)</tr>"
}

$backupdata += "</table>"

#------------------------------------------------------HTML + CSS---------------------------------------------------------#

# CSS new updated
$HtmlStyle = @"
<style>
TABLE {table-layout:fixed ;align:center;width: 50px;text-align: center;table-layout:auto;width:0.1%;white-space: nowrap;border-width:0.5px;background-color:white; border-style: solid;border-color: black;border-collapse: collapse;margin-left:auto;margin-right:auto;font-size:15px}
TH {padding: 0px;text-align: center;width:0.1%;white-space: nowrap;border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #5F249F;color: #ffffff;}
TD {width:150px;padding: 0px;text-align: center;width:0.1%;white-space: nowrap;border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
td.redfont {color:white;}
</style>
"@

$html = "<html>

<head>
<meta name='viewport' content='width=device-width, initial-scale=1'>
<title>Test Report</title>
$HtmlStyle
</head>

<body style='background-color:ADD8E6';>
<style>
.header{
background-color: #5F249F;
color:white;
text-align: center;
padding-top:8px;
padding-bottom:8px;}
p{
font-size:40px;
font-weight: bold;
}
td.Redfont {color:white;}
</style>
<div class='header'>
<p class='alstom'>[ALSTOM TRANSPORT] VxRail Healthcheck and Utilization Report <p>
</div>
<nav class='navbar sticky-top navbar-dark' style='background-color: #5F249F;text-align:center;''>
</nav>
</center>
<div align='right'><h5>$date</h5></div>
<center>
$result
$backupdata
</center>
</body>
</html>"

$html | Out-File $html_report
Write-Host "`tHealth check HTML Report generated`t" -ForegroundColor Yellow

#-------------------------------------------------------------Email as HTML-------------------------------------------------------#
Write-Host "`tPreparing to send email..."  -ForegroundColor Yellow

$recepients = @('sivakumark@dxc.com')
$ccrecepients = @('dxcindiahci@dxc.com')
Send-MailMessage -SmtpServer  mail-apps.alstom.hub -to $recepients -Cc $ccrecepients -Subject '[ALSTOM TRANSPORT] VxRail Health Daily Status'  -From alstom_vxrail@alstomgroup.com -BodyAsHtml ([string](Get-Content $html_report))

Write-Host "`tEmail sent!" -ForegroundColor Green
#---------------------------------------------------------------THE END-----------------------------------------------------------#