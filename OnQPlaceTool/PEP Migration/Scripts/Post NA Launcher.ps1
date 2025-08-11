function Ensure-PathExists {
    param (
        [string]$path
    )

    if (-Not (Test-Path -Path $path)) {
        $null = New-Item -ItemType Directory -Path $path
    }
}

$upRootPath = Split-Path -path $PSScriptRoot -Parent
$assetsPath = Join-Path -Path $upRootPath "Assets"
$timestampFile = Join-Path -Path $assetsPath "NA Launcher Time Stamp.txt"

if (-not (Test-Path $timestampFile)) {
    New-Item -ItemType File -Path $timestampFile -Force | Out-Null
}

$now = Get-Date

# this will stop the script from running multiple times if the backups occur in quick succession (which they do)
if (Test-Path $timestampFile) {
    $lastRun = Get-Content $timestampFile | Get-Date
    $now.ToString() | Set-Content $timestampFile
    try{$elapsed = ($now - $lastRun).TotalMinutes}catch{$elapsed = 0}
    if ($elapsed -lt 5) {
        exit
    }
}

# This checks if Night Audit was ran before midnight, if it was, we assume it is for PEP migration.

$databaseName = "HPMS3"
$query = @"
USE HPMS3;

SELECT 
    p.property_id,
    p.property_name,
    p.cur_bus_date,
    a.audit_marker_code,
    a.description,
    a.audit_procedure,
    a.date_completed,
    a.start_time,
    a.end_time,
    a.real_end_time
FROM dbo.PROPERTY p
CROSS JOIN dbo.AUDIT_MARKER a
WHERE a.audit_marker_code = 'DTBMP';
"@

# Run the SQL command
$sqlresult = Invoke-Sqlcmd -ServerInstance $env:COMPUTERNAME -Database $databaseName -Query $query

if (($sqlresult.cur_bus_date | Get-Date) -lt ($sqlresult.end_time | Get-Date)){
exit
}

# Write the current time to stop subsequent launches
$now.ToString() | Set-Content $timestampFile

$basePath = Split-Path -Path (split-path -Path $PSCommandPath)

Ensure-PathExists -path "$basePath\Logs"

#logic here to launch other scripts:

$scriptPaths = @(
"D:\PEP Migration\Scripts\OnQ Balancing Report Tool.ps1",
"D:\PEP Migration\Scripts\Disable Interface Services.ps1"
)

foreach ($script in $scriptPaths) {
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$script`"" -NoNewWindow
}

# this will stop the script from running multiple times if the backups occur in quick succession (which they do)
if (Test-Path $timestampFile) {
    $lastRun = Get-Content $timestampFile | Get-Date
    if (($now - $lastRun).TotalMinutes -lt 5) {
        exit
    }
}

# This checks if Night Audit was ran before midnight, if it was, we assume it is for PEP migration.

$databaseName = "HPMS3"
$query = @"
USE HPMS3;

SELECT 
    p.property_id,
    p.property_name,
    p.cur_bus_date,
    a.audit_marker_code,
    a.description,
    a.audit_procedure,
    a.date_completed,
    a.start_time,
    a.end_time,
    a.real_end_time
FROM dbo.PROPERTY p
CROSS JOIN dbo.AUDIT_MARKER a
WHERE a.audit_marker_code = 'DTBMP';
"@

# Run the SQL command
$sqlresult = Invoke-Sqlcmd -ServerInstance $env:COMPUTERNAME -Database $databaseName -Query $query

if (($sqlresult.cur_bus_date | Get-Date) -lt ($sqlresult.end_time | Get-Date)){
exit
}

# Write the current time to stop subsequent launches
$now.ToString() | Set-Content $timestampFile

$basePath = Split-Path -Path (split-path -Path $PSCommandPath)

Ensure-PathExists -path "$basePath\Logs"

#logic here to launch other scripts:

$scriptPaths = @(
"D:\PEP Migration\Scripts\OnQ Balancing Report Tool.ps1",
"D:\PEP Migration\Scripts\Disable Interface Services.ps1"
)

foreach ($script in $scriptPaths) {
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$script`"" -NoNewWindow

}

