# Ensure TLS 1.2 is used for secure downloads
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Define target version and script path
$moduleName = "ImportExcel"
$targetVersion = "7.8.10"
$modulePath = Join-Path -Path $PSScriptRoot -ChildPath $moduleName

# Check if ImportExcel is installed
$installed = Get-Module -ListAvailable -Name $moduleName

if (-not $installed) {
    Write-Host "ImportExcel $targetVersion not found. Installing for current user..."

    # Install the module for the current user
    Install-Module -Name $moduleName -RequiredVersion $targetVersion -Scope CurrentUser -Force

    # Get the installed module path
    $installed = Get-Module -ListAvailable -Name $moduleName | Where-Object { $_.Version -eq $targetVersion }
}

if ($installed) {
    $sourcePath = Split-Path $installed.Path -Parent

    # Copy module to script directory
    Copy-Item -Path $sourcePath -Destination $modulePath -Recurse -Force
    Write-Host "Module copied to $modulePath"
}
else {
    Write-Error "Failed to locate the installed module."
}
