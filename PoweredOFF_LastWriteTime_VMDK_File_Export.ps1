# =============================================================================================================
# Script:    PoweredOFF_LastWriteTime_VMDK_File_Export.ps1
# Version:   1.0
# Date:      January, 2022
# By:        Rajesh R
# =============================================================================================================

<#
.SYNOPSIS
  Fetch the PoweredOff VM's and filtering out the VMDK.
.DESCRIPTION
  This script will fetch the PoweredOff VM and its VMDK's last write time from vCenter.
.INPUTS
    $VIServer - The name or IP address of your VMWare vCenter server.
    $VMTemplateName - The name of the template user want to apply a patch
.OUTPUTS
  Log transcript file will be stored in C:\Temp\PoweredOffVMvmdk-LogFile.log
  
.EXAMPLE
  #This script will help us filter out the Last write time of the VMDK of a VM which is set to 100 days (Lastwritetime). 
  Based upon this condition the report is generated and stored in specified path.
    ./PoweredOffVMvmdk-Report.csv
#>

#-------------------------------------------------------[Initialization]-------------------------------------------------------#

Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
$ErrorActionPreference="stop"
$timeLimit = (Get-Date).AddDays(-100)
$path = "C:\temp\"
$report = "$path\PoweredOffVMvmdk-Report.csv" # Final Report
$errLogPath = "$path\PoweredOffVMvmdk-ErrorLogs.log" # Error Logs
$LogFile = "$path\PoweredOffVMvmdk-LogFile.log" # Detailed Logs
$vCenterList = Import-Csv -Path "$path\vCenterList.csv"

#---------------------------------------------------------[Credentials]---------------------------------------------------------#

$cred = get-credential "Enter Username"
$username = $cred.UserName
$password = $cred.Password
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
$PWD = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

#----------------------------------------------------------[Functions]----------------------------------------------------------#

Function Connect-vCenter{

try{

Connect-VIServer -Server $vCenter -Username $username -Password $PWD -ErrorAction Stop

}

catch{
$connErr = $_.exception.message
Write-Output "Unable to connect to vCenter. Error:$_" 
}
}

Function Get-PoweredOffVMvmdk{

try {
foreach ($Datacenter in (Get-Datacenter | Sort-Object -Property Name)) {
foreach ($Cluster in ($Datacenter | Get-Cluster | Sort-Object -Property Name)) {
foreach ($VM in ($Cluster | Get-VM | Where-Object {$_.PowerState -eq "PoweredOff"} | Sort-Object -Property Name)) {
#$VM = Get-VM -Name "StuartVault"
$datastores = Get-Datastore -VM $VM | Sort-Object -Property Name

foreach ($datastore in ($datastores)) {
Write-Output $datastore
$datastore.datacenter.name
$datastore.name
$lastwritetimes = Get-ChildItem -Recurse -Path "vmstore:\$($datastore.datacenter.name)\$($datastore.name)\$($VM.name)" -Include *vmdk* | Select-Object Name,DatastoreFullPath,LastWriteTime
Write-Output $lastwritetimes

foreach ($lastwritetime in $lastwritetimes) {
if ($lastwritetime.LastWriteTime -gt $timeLimit) {
Write-Host "`t`tLast write time of vmdk for $VM is over 100 days.`t`t" -ForegroundColor Green -BackgroundColor DarkRed

New-Object -TypeName PSObject -Property @{
"PowerOffVMName" = $VM.name
"Datacenter" = $datastore.datacenter.name
"DatastoreFullPath" = $lastwritetime.DatastoreFullPath
"VMDKFileName" = $lastwritetime.name
"LastWriteTime" = $lastwritetime.lastwritetime
} | Export-Csv -Path $report -Append -NoTypeInformation
}

else {
Write-Output "Last write time is not greater. Not appending data"
}
}
}
}
}
}
}

catch [Exception] {
$errorLog = $PSItem.exception.message
Write-Output "Something went wrong! Error:$errorLog" >> $errLogPath
Write-Host -ForegroundColor Red -BackgroundColor White "Error encountered: $_.exception"
}
}

Function Disconnect-vCenter{

try{
Disconnect-VIServer -server $vCenter -confirm:$false
Write-Host "Finished script execution. Successfully disconnected from $vCenter." -ForegroundColor Green
}

catch{
$disconErr = $_.exception.message
Write-Host "Something went wrong, please try disconnecting again."
Write-Output $disErr
}
}

Function Report-SendMail{

Start-Sleep -Seconds 3

try{

$dateStamp = ( Get-Date ).ToString(‘dd-MMM-yyyy h_mm_tt’)
$FromEml = "Enter email"
$ToEml = "Enter email"
$SubEml = "PoweredOff & Last Write Time of Virtual Machine VMDK Report - $dateStamp"
$SMTPCfg = "Enter SMTP server details"
$SMTPport = "12"
$vDisk = "F:\New_test\$vCenter-PoweredOffVMvmdk-Report.xlsx"
$bodyEml = "PoweredOff & Last Write Time of Virtual Machine VMDK Report"
$Arguments = "/SMTPserver $SMTPCfg /smtpport $SMTPport /mailto $ToEml /mailfrom $FromEml /mailsubject $SubEml /attachment $report"

Send-MailMessage –From $FromEml –To $ToEml -Bcc "Enter Email" –Subject $SubEml -Body $bodyEml -BodyAsHtml –SmtpServer $SMTPCfg -Attachments $report

}
catch{

$errorMail = $_.exception.message
Write-Output "Error: $errorMail"
Write-Host -ForegroundColor Yellow -BackgroundColor Red "`tCould not send email, please check e-mail configuration`t`t"

}
}

#----------------------------------------------------------[Execution]----------------------------------------------------------#

$ErrorActionPreference
Start-Transcript -Path $LogFile -Append
Write-Host ("-" *30) [Started Execution] ("-" *30)
$StartedTime = (Get-Date).ToString('MM/dd/yyyy hh:mm:ss tt K') + "$([TimeZoneInfo]::Local.Id)"
Write-Host "Execution started at: $StartedTime"

foreach ($row in $vCenterList){
$vCenter = $row.vcenter
Connect-vCenter
Get-PoweredOffVMvmdk
Disconnect-vCenter
}

if (test-path $report){
Report-SendMail
Write-Host -ForegroundColor Green -BackgroundColor Black "`tMail has been sent successfully.."
}

else{
Write-Host -ForegroundColor Red -BackgroundColor Black "`tMail not sent, 'Report' file missing.."
}

$EndTime = (Get-Date).ToString('MM/dd/yyyy hh:mm:ss tt K') + "$([TimeZoneInfo]::Local.Id)"
Write-Host "Execution Finished at: $EndTime"
Write-Host ("-" *30) [Finished Execution] ("-" *30)
Stop-Transcript

#-------------------------------------------------------[End of the Script]-----------------------------------------------------#