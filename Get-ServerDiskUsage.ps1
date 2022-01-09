# install-module ImportExcel
# $e = Open-ExcelPackage -Path .\server.xlsx
# $d = $e.Workbook.Worksheets['Disks']
# Close-ExcelPackage $e

function Get-ServerDisks {
    param (
        $server,
        $timestamp
    )
    if (test-path ".\data\$server.csv") {
        $serverdrives = Import-Csv ".\data\$server.csv"
        $drives = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $server | Where-Object { $_.DriveType -eq 3 } `
        | Select-Object DeviceID, FreeSpace
        foreach ($drive in $drives) {
            foreach ($serverdrive in $serverdrives) {
                if (($serverdrive.drive -contains ($drive).DeviceID) -and ($serverdrive.date -contains $timestamp)) {
                    $found = $true
                    break
                }
            }
            if (-not($found)) {
                $output = [PSCustomObject]@{
                    Drive     = $drive.DeviceID
                    FreeSpace = $drive.FreeSpace / "1MB"
                    Date      = $timestamp
                }
                Export-Csv -path ".\data\$server.csv" -InputObject $output -Append
            }
        }
    }
    else {
        #first run
        try {
            $drives = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $server -ErrorAction SilentlyContinue `
            | Where-Object { $_.DriveType -eq 3 } | Select-Object DeviceID, FreeSpace
        }
        catch {
        }
        if ($drives) {
            foreach ($drive in $drives) {
                $output = [PSCustomObject]@{
                    Drive     = $drive.DeviceID
                    FreeSpace = $drive.FreeSpace / "1MB"
                    Date      = $timestamp
                }
                Export-Csv -path ".\data\$server.csv" -InputObject $output -Append
            }
        }
    }
}

$servers = Get-Content .\servers.txt
$timestamp = (Get-Date).ToShortDateString()
$report = @()
foreach ($server in $servers) {
    Get-ServerDisks -server $server -timestamp $timestamp
    try {
        $serverdrives = Import-Csv ".\data\$server.csv"
        $drives = $serverdrives.drive | Sort-Object | Get-Unique
        foreach ($drive in $drives) {
            $totaldays = 0
            [long]$difference = 0
            #get initial free space
            $start = $serverdrives | Where-Object { $_.Drive -eq $drive } | Select-Object -First 1
            #get everything after that
            $end = $serverdrives | Where-Object { $_.Drive -eq $drive } | Select-Object -Last 1
            if ($end -ne $start) {
                $difference = [long]$start.freespace - [long]$end.freespace
                #create timespan and calculate rate of change per week/month
                $totaldays = (New-TimeSpan -start $start.Date -end $end.Date).Days
                if ($difference -gt 0) {
                    $rate = $difference / $totaldays
                    $forecast = [long]$end.freespace / $rate
                    $reportdrive = [PSCustomObject]@{
                        Server       = $server
                        Drive        = $end.Drive
                        FreeSpace    = $end.FreeSpace.split('.')[0]
                        Changed      = $difference
                        Days         = $totaldays
                        CapacityDate = (Get-Date).adddays($forecast)
                    }
                    $report += $reportdrive
                }
                else {
                    $reportdrive = [PSCustomObject]@{
                        Server       = $server
                        Drive        = $end.Drive
                        FreeSpace    = $end.FreeSpace.split('.')[0]
                        Changed      = $difference
                        Days         = $totaldays
                        CapacityDate = "No change or free space increased"
                    }
                    $report += $reportdrive
                }
            }
            else {
                $reportdrive = [PSCustomObject]@{
                    Server       = $server
                    Drive        = $end.Drive
                    FreeSpace    = $end.FreeSpace.split('.')[0]
                    Changed      = $difference
                    Days         = $totaldays
                    CapacityDate = "First run"
                }
                $report += $reportdrive
            }
        }
    }
    catch {
            $reportdrive = [PSCustomObject]@{
                Server       = $server
                Drive        = "N/A"
                FreeSpace    = "N/A"
                Changed      = "N/A"
                Days         = "N/A"
                CapacityDate = "Unable to collect drive info"
            }
            $report += $reportdrive
    }
    
}
$d = Get-Date
$reportdate = "-" + $d.Year.ToString() + "-" + ("{0:D2}" -f $d.Month).ToString() + "-" + ("{0:D2}" -f $d.day).ToString()
if (-not(Test-Path ".\DiskReport$reportdate.csv")) {
    $report | Export-Csv -Path ".\DiskReport$reportdate.csv"   
}