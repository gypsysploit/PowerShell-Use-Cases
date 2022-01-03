# =============================================================================================================
# Script:    Excel_Cleanup_Pivot.ps1
# Version:   1.0
# Date:      January, 2022
# By:        Rajesh R
# =============================================================================================================

<#
.SYNOPSIS
With this script you can do the clean-up of excel workbook, worksheet and create pivot and saves the excel file.
	
.DESCRIPTION
With this script you can do the clean-up of excel workbook, worksheet and create pivot.
This script must be started on a Windows machine where Excel Application and Excel Module is installed!

.PARAMETER InputFile
Path and Name of the xlsx file


.PARAMETER OutputFile
Path and Name of the output xlsx file (Saves on Original xlsx file)


.EXAMPLE
 .\Excel_Cleanup_Pivot.ps1 runs in the location C:\temp\FileName-ddMMyyyy.xlsx and OutputFile is saved in the same path as input path

Script will first check if the xlsx data file is present with current date, only if it matches then script executes and does the cleanup and saves the final data in the same path.

#>

#---------------------------------------------------------[ImportExcel Module]-----------------------------------------------------#

Write-Host -ForegroundColor Cyan "Importing Excel Module.."
Import-Module -Name ImportExcel
Write-Host -ForegroundColor Cyan "Excel Module Imported!"

#---------------------------------------------------------------[Begin]-------------------------------------------------------------#

$ErrorActionPreference = “SilentlyContinue”

Write-Host -ForegroundColor Green "Fetching the xlsx file in the directory.."

$file = "Path"
Set-Location $file

Write-Progress -Activity "Running Script" -Status "Progress.." -CurrentOperation "xlsx file is being Cleaned-up"

if ($($file | Get-ChildItem -Recurse -Filter *.xlsx | Where-Object {$_.Name -match (Get-Date -Format ddMMyyyy)})){
$FileFound = ($file | Get-ChildItem -Recurse -Filter *.xlsx | Where-Object {$_.Name -match (Get-Date -Format ddMMyyyy)}).FullName
Write-Host "Xlsx file found:`t$FileFound" -ForegroundColor Green -BackgroundColor Black
}

else{
throw (New-Object System.IO.FileNotFoundException ("File not found!"))
#$FileFound = ($WorkBook | Get-ChildItem -Recurse -Filter *.xlsx | Where-Object {$_.Name -match (Get-Date -Format ddMMyyyy)}).Name
}

Start-Sleep -Seconds 2

$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $true
$excel.displayalerts = $false
$Workbook = $Excel.workbooks.open($FileFound)
Write-Output "Removing unwanted Sheets.."
($workbook.Worksheets | Where-Object {$_.Name -ne 'vMemory'}).Delete()

$Worksheet = $Workbook.Worksheets.Item('vMemory')
Start-Sleep -Seconds 1
$range1 = $worksheet.Range("C:C,E:E,F:F,G:G,H:H,I:I,J:J,K:K,L:L,M:M,N:N,O:O,P:P,Q:Q,R:R,T:T,U:U,V:V,AB:AB,AC:AC,AD:AD,AE:AE,AF:AF,AG:AG")
$range1.EntireColumn.delete()
$workbook.Save()
Start-Sleep -Seconds 1

$ClusterRange = $Worksheet.UsedRange
$ClusterRange.AutoFilter()
$Range = $Worksheet.Range("H1").EntireColumn
Write-Output "Removing random IDs in the Folder column.."
$Search = $Range.Replace(" (*)", "", [Microsoft.Office.Interop.Excel.XlLookAt]::xlPart)
$ClusterRange = $Worksheet.UsedRange
$ClusterRange.AutoFilter()
$CriteriaArray = @('cluster-resources-fivp-01', 'cluster-resources-tel-01')
#$xlFilterValues = 7
#$ClusterRange.AutoFilter(6,$CriteriaArray, $xlFilterValues)
#$workbook.Save()
Write-Output "Filtering cluster rows with: $CriteriaArray. Removing other rows.."
for ($i = $worksheet.usedrange.rows.count; $i -gt 1; $i--)
{ 
    $text = $worksheet.Range("F$i").Text
    if ($CriteriaArray -contains $text)
    {
        $found = $true
    }
    else
    {
        $found = $false
    }

    if (!$found)
    {
        $WorkSheet.Cells.Item($i, $i).EntireRow.Delete() | Out-Null
    }
}

$workbook.Save()
$IgnoreArray = @('/GRE02-VxRail/NSX Controller Nodes', '/GRE02-VxRail/Discovered virtual machine', '/GRE02-VxRail', '/GRE02-VxRail/vcd1/Service VMs')

Write-Output "Deleting unwanted rows in Folder column info: $IgnoreArray"
for ($i = $worksheet.usedrange.rows.count; $i -gt 0; $i--)
{ 
    $text = $worksheet.Range("H$i").Text
    if ($IgnoreArray -contains $text)
    {
        $found = $true
    }
    else
    {
        $found = $false
    }

    if ($found)
    {
        $WorkSheet.Cells.Item($i, $i).EntireRow.Delete() | Out-Null
    }
}

$workbook.Save()
$WorkBook.Close($true)
$Excel.Quit()

$importvMemory = Import-Excel -Path $FileFound -ErrorAction SilentlyContinue -WorksheetName "vMemory"
$importvMemory | Select-Object Folder, VM, 'Size MB', Reservation, Powerstate, Datacenter, Cluster, Host, 'OS according to the configuration file' | Export-Excel $FileFound -AutoSize -WorksheetName "vMemory" -IncludePivotTable ` -PivotRows Folder ` -PivotData @{VM = "count"; Reservation = "sum"} ` -PivotDataToColumn


$Worksheet, $Workbook, $Excel | ForEach-Object {
        if ($_ -ne $null) {
            [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
        }
}

Remove-Variable -Name $Excel $Workbook $Worksheet
$Excel = $null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
$Excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Worksheet)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Range)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($range1)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)

Write-Host -ForegroundColor Magenta "Closing the Excel Application..."
Write-Host -ForegroundColor Green -BackgroundColor Black "Execution successfully completed!"

#----------------------------------------------------------[Script Ends]------------------------------------------------------------#