# Set your log directory here
$LogDirectory = "D:\PEP Migration\Logs"
$LogFile = Join-Path $LogDirectory ("IFC_Service_Disable_Log_" + (Get-Date -Format "yyyy-MM-dd") + ".log")

# Function to log messages
function Log-Message {
    param ([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $LogFile -Append -Encoding UTF8
}

# Start logging
Log-Message "=== Script started ==="

# Define service name patterns (wildcards allowed)
$ServicePatterns = @(
    'OnQIFC*',
    'Lims',
    'V5_DoorLockServer'
)

foreach ($pattern in $ServicePatterns) {
    $IfcServices = $null
    $IfcServices = Get-Service -Name $pattern -ErrorAction SilentlyContinue

    if ($IfcServices.Count -gt 0) {
        foreach ($IfcService in $IfcServices) {
            try {
                Log-Message "Processing service: $($IfcService.Name) (Status: $($IfcService.Status))"

                if ($IfcService.Status -eq 'Running') {
                    $ServiceInfo = Get-WmiObject -Class Win32_Service -Filter "Name = '$($IfcService.Name)'"
                    $ProcessId = $ServiceInfo.ProcessId

                    if ($ProcessId) {
                        try {
                            Stop-Process -Id $ProcessId -Force -ErrorAction Stop
                            Log-Message "Stopped process ID $ProcessId for service $($IfcService.Name)"
                            Start-Sleep -Seconds 5
                        } catch {
                            Log-Message "WARNING: Failed to stop process ID $ProcessId - $_"
                        }
                    }

                    Stop-Service -Name $IfcService.Name -Force -ErrorAction Stop
                    Log-Message "Stopped service $($IfcService.Name)"
                    Start-Sleep -Seconds 5
                }

                Set-Service -Name $IfcService.Name -StartupType Disabled -ErrorAction Stop
                Log-Message "Disabled service $($IfcService.Name)"
            } catch {
                Log-Message "ERROR: Failed to process service $($IfcService.Name) - $_"
            }
        }
    } else {
        Log-Message "No services found matching pattern '$pattern'"
    }
}

Log-Message "=== Script completed ==="
Exit
