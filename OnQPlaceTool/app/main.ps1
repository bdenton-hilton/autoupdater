#region - NA Flag XML
$preNaFlagXml = @"
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <Triggers>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="Application"&gt;&lt;Select Path="Application"&gt;*[System[Provider[@Name='MSSQLSERVER'] and EventID=18265]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
    </EventTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT2H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe</Command>
      <Arguments>-ExecutionPolicy Bypass -File "D:\PEP Migration\Scripts\Pre NA Launcher.ps1"</Arguments>
    </Exec>
  </Actions>
</Task>
"@

$postNaFlagXml = @"
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <Triggers>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="Application"&gt;&lt;Select Path="Application"&gt;*[System[Provider[@Name='MSSQLSERVER'] and EventID=8957]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
    </EventTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT2H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe</Command>
      <Arguments>-ExecutionPolicy Bypass -File "D:\PEP Migration\Scripts\Post NA Launcher.ps1"</Arguments>
    </Exec>
  </Actions>
</Task>
"@
#endregion

function Get-DirectoryBlobTree {
    param (
        [string]$rootDirectory
    )

    if (-not (Test-Path $rootDirectory)) {
        return $null
    }

    $result = @()

    # Recursively get all files in the directory
    $files = Get-ChildItem -Path $rootDirectory -Recurse -File

    foreach ($file in $files) {
        $relativePath = $file.FullName.Substring($rootDirectory.Length).TrimStart('\')
        $blobSha1 = Compute-BlobSHA1 -filePath $file.FullName

        $result += [PSCustomObject]@{
            RelativePath = $relativePath
            BlobSHA1     = $blobSha1
        }
    }

    return $result
}


# Shared state reference
$sharedState = [ref]@{
    keepWindowOpen = $true
    windowClosed   = $false
    jobPaused      = $false
    dataTable      = @{}
}

#region - Draw Data Table

Add-Type -TypeDefinition @"
using System.ComponentModel;

public class InnStatus : INotifyPropertyChanged {
    public event PropertyChangedEventHandler PropertyChanged;

    private string _innCode;
    private string _status;
    private string _result;

    public string InnCode {
        get { return _innCode; }
        set {
            _innCode = value;
            if (PropertyChanged != null)
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("InnCode"));
        }
    }

    public string Status {
        get { return _status; }
        set {
            _status = value;
            if (PropertyChanged != null)
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Status"));
        }
    }

    public string Result {
        get { return _result; }
        set {
            _result = value;
            if (PropertyChanged != null)
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs("Result"));
        }
    }
}
"@


# Set up UI runspace
$runspace = [runspacefactory]::CreateRunspace()
$runspace.ApartmentState = 'STA'
$runspace.Open()

$ps = [powershell]::Create()
$ps.Runspace = $runspace

$ps.AddScript({
        param($stateRef)

        Add-Type -AssemblyName PresentationFramework

        $state = $stateRef.Value
        $dataCollection = [System.Collections.ObjectModel.ObservableCollection[InnStatus]]::new()

        function Write-DebugLine {
            param([string]$msg)
            $debugBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] $msg`r`n")
            $debugBox.ScrollToEnd()
        }

        function Sync-Collection {
            foreach ($key in $state.dataTable.Keys) {
                $entry = $state.dataTable[$key]
                $row = $dataCollection | Where-Object { $_.InnCode -eq $key }

                if ($row) {
                    if ($row.Status -ne $entry.Status) { $row.Status = $entry.Status }
                    if ($row.Result -ne $entry.Result) { $row.Result = $entry.Result }
                    Write-DebugLine "Updated $key`: $($row.Status), $($row.Result)"
                }
                else {
                    $newEntry = [InnStatus]::new()
                    $newEntry.InnCode = $key
                    $newEntry.Status = $entry.Status
                    $newEntry.Result = $entry.Result
                    $dataCollection.Add($newEntry)
                    Write-DebugLine "Added $key to table"
                }
            }
        }



        $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="OnQ Placement Status Tracker" Height="600" Width="436"
        WindowStartupLocation="CenterScreen"
        Background="White" FontSize="16" FontFamily="Lato"
        Topmost="False">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <DataGrid x:Name="DataTableGrid" Grid.Row="0"
                  ScrollViewer.HorizontalScrollBarVisibility="Auto"
                  ScrollViewer.CanContentScroll="False"
                  AutoGenerateColumns="False"
                  RowDetailsVisibilityMode="Collapsed"
                  HeadersVisibility="Column"
                  CanUserAddRows="False"
                  CanUserSortColumns="True"
                  GridLinesVisibility="None"
                  Margin="0,0,0,12"
                  FontSize="16">
            <DataGrid.Columns>
                <DataGridTextColumn Header="InnCode" Binding="{Binding InnCode}" />
                <DataGridTextColumn Header="Status" Binding="{Binding Status}" />
                <DataGridTextColumn Header="Result" Binding="{Binding Result}" Width="*" />
            </DataGrid.Columns>
        </DataGrid>

<Grid Grid.Row="1">
    <Grid.Resources>
        <Style TargetType="Button">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="5"> <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"
                                              VerticalAlignment="{TemplateBinding VerticalContentAlignment}"
                                              Content="{TemplateBinding Content}"
                                              Margin="{TemplateBinding Padding}" />
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Grid.Resources>

    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto" />
        <ColumnDefinition Width="*" />
        <ColumnDefinition Width="Auto" />
    </Grid.ColumnDefinitions>

    <Button x:Name="PauseButton" Grid.Column="0" Height="40" Content="Pause Jobs"
            Background="#FF409EFF" Foreground="White"
            BorderBrush="#FF3077C9" BorderThickness="1"
            Padding="20,10" Margin="0,0,8,0" Cursor="Hand" Visibility="Collapsed"/>

    <Button x:Name="ExportButton" Grid.Column="2" Width="100" Height="40" Content="Export"
            Background="#FF409EFF" Foreground="White"
            BorderBrush="#FF3077C9" BorderThickness="1"
            Padding="20,10" Cursor="Hand" />
</Grid>

        <TextBox x:Name="DebugOutput" Grid.Row="2" Height="100"
                 Margin="0,8,0,0" VerticalScrollBarVisibility="Auto"
                 TextWrapping="Wrap" IsReadOnly="True"
                 FontSize="14" FontFamily="Consolas"
                 Background="#FFF5F5F5" Foreground="Black"
                 BorderBrush="LightGray" BorderThickness="1" Visibility="Collapsed"/>
    </Grid>
</Window>
"@

        $reader = [System.Xml.XmlTextReader]::new([System.IO.StringReader]$xaml)
        $window = [Windows.Markup.XamlReader]::Load($reader)

        $grid = $window.FindName("DataTableGrid")
        $exportBtn = $window.FindName("ExportButton")
        $pauseBtn = $window.FindName("PauseButton")
        $debugBox = $window.FindName("DebugOutput")
        $grid.ItemsSource = $dataCollection

        $timer = [System.Windows.Threading.DispatcherTimer]::new()
        $timer.Interval = [TimeSpan]::FromMilliseconds(125)
        $timer.Add_Tick({
        
                Sync-Collection
                Write-DebugLine "Sync triggered. Total rows: $($dataCollection.Count)"
                if (-not $state.keepWindowOpen) {
                    $state.windowClosed = $true
                    $timer.Stop()
                    Write-DebugLine "Window closed by main loop."
                    $window.Close()
                }
            })

        $timer.Start()

        $exportBtn.Add_Click({
                $dialog = [Microsoft.Win32.SaveFileDialog]::new()
                $dialog.Filter = "CSV file (*.csv)|*.csv"
                if ($dialog.ShowDialog()) {
                    $dataCollection | Export-Csv -Path $dialog.FileName -NoTypeInformation
                    [System.Windows.MessageBox]::Show("Exported to $($dialog.FileName)", "Export Success")
                    Write-DebugLine "Exported CSV to $($dialog.FileName)"
                }
            })

        $pauseBtn.Add_Click({
                $state.jobPaused = -not $state.jobPaused
                $pauseBtn.Content = if ($state.jobPaused) { "Start Jobs" } else { "Pause Jobs" }
                Write-DebugLine "JobPaused toggled: $($state.jobPaused)"
            })

        $null = $window.ShowDialog()
        Write-DebugLine "Grid binding count: $($grid.Items.Count)"

        $state.windowClosed = $true
    }) | Out-Null


$ps.AddArgument($sharedState) | Out-Null
$null = $ps.BeginInvoke()

# Setup
$cred = $uiResult.AdmCred
$localFolder = $uiResult.FolderToPlace
$remoteRoot = $uiResult.TargetLocation
$localFolderName = Split-Path $localFolder -Leaf

$localBlob = Get-DirectoryBlobTree -rootDirectory $localFolder

# Ensure remote root ends with a single backslash
if (-not $remoteRoot.EndsWith("\")) {
    $remoteRoot += "\"
}
elseif ($remoteRoot -match "\\{2,}$") {
    $remoteRoot = $remoteRoot.TrimEnd('\') + "\"
}

$remoteTarget = $remoteRoot + $localFolderName
$zipPath = Join-Path $env:TEMP "$localFolderName.zip"
Remove-Item $zipPath -Force -ErrorAction Ignore

# Zip entire local folder
Add-Type -Assembly "System.IO.Compression.FileSystem"
[IO.Compression.ZipFile]::CreateFromDirectory($localFolder, $zipPath)

start-sleep -Milliseconds 200

# Create jobs per device
$jobs = @()

foreach ($action in $uiResult.Actions) {

    $server = $action.serverName
    $inncode = $action.inncode
    $createListener = $action.createListener

    if ([string]::IsNullOrEmpty($server)) {
        continue
    }

    $sharedState.Value.dataTable[$inncode] = @{
        Status = "Pending"
        Result = "Initializing"
    }

    
    $jobs += Start-Job -Name "$inncode" -ScriptBlock {
        param($cred, $server, $remoteTarget, $zipPath, $localFolderName, $postNaFlagXml, $preNaFlagXml, $createListener, $localBlob)

        try {
            $session = New-PSSession -ComputerName $server -Credential $cred -errorAction Stop
            $connection = $true
        }
        catch { 
            $connection = $false
        }

        if ($connection) {
            $createListenerResult = Invoke-Command -Session $session -ScriptBlock {
                param($dest, $createListener, $postNaFlagXml, $preNaFlagXml, $cred)
                if (-not (Test-Path -Path $dest -PathType Container)) {
                    New-Item -ItemType Directory -Path $dest | Out-Null
                }
                if ($createListener) {

                    $postMessage = "An error occured scheduling the Post-flag."
                    $taskName = "Post NA Migration Flag"
                    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
                        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
                    }
                    $task = [xml]$postNaFlagXml
                    $registeredTask = Register-ScheduledTask -Xml $task.OuterXml -TaskName $taskName -User $cred.UserName -Password $cred.GetNetworkCredential().Password
                    $registeredTaskName = $registeredTask.TaskName
                    $registeredTaskState = $registeredTask.State
                    if ($registeredTaskState -eq 3) {
                        $postMessage = "Successfully Scheduled $registeredTaskName."
                    }

                    $preMessage = "An error occured scheduling the Pre-flag. "
                    $taskName = "Pre NA Migration Flag"
                    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
                        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
                    }
                    $task = [xml]$preNaFlagXml
                    $registeredTask = Register-ScheduledTask -Xml $task.OuterXml -TaskName $taskName -User $cred.UserName -Password $cred.GetNetworkCredential().Password
                    $registeredTaskName = $registeredTask.TaskName
                    $registeredTaskState = $registeredTask.State
                    if ($registeredTaskState -eq 3) {
                        $preMessage =  "Successfully Scheduled $registeredTaskName. "
                    }

                    $taskMessage = $preMessage + $postMessage
                    return $taskMessage

                }
                else {
                    return ""
                }
            }-ArgumentList $remoteTarget, $createListener, $postNaFlagXml, $cred

            # Validate remote zip path
            if (-not $remoteTarget.EndsWith("\")) {
                $remoteTarget += "\"
            }

            $remoteBlob = Invoke-Command -Session $session -ScriptBlock {
                param($remoteTarget)
                function Get-DirectoryBlobTree {
                    param (
                        [string]$rootDirectory
                    )

                    if (-not (Test-Path $rootDirectory)) {
                        return $null
                    }

                    $result = @()

                    # Recursively get all files in the directory
                    $files = Get-ChildItem -Path $rootDirectory -Recurse -File

                    foreach ($file in $files) {
                        $relativePath = $file.FullName.Substring($rootDirectory.Length).TrimStart('\')
                        $blobSha1 = Compute-BlobSHA1 -filePath $file.FullName

                        $result += [PSCustomObject]@{
                            RelativePath = $relativePath
                            BlobSHA1     = $blobSha1
                        }
                    }

                    return $result
                }

                function Compute-BlobSHA1 {
                    param (
                        [string]$filePath
                    )

                    if (Test-Path $filePath) {
                        # Read the file content
                        $fileContent = [System.IO.File]::ReadAllBytes($filePath)

                        # Create the blob header
                        $blobHeader = [System.Text.Encoding]::UTF8.GetBytes("blob $($fileContent.Length)`0")

                        # Combine the header and the file content
                        $blobData = New-Object byte[] ($blobHeader.Length + $fileContent.Length)
                        [System.Buffer]::BlockCopy($blobHeader, 0, $blobData, 0, $blobHeader.Length)
                        [System.Buffer]::BlockCopy($fileContent, 0, $blobData, $blobHeader.Length, $fileContent.Length)

                        # Compute the SHA-1 hash
                        $sha1 = [System.Security.Cryptography.SHA1]::Create()
                        $hashBytes = $sha1.ComputeHash($blobData)
                        $sha1Hash = [BitConverter]::ToString($hashBytes) -replace '-', ''
                        $sha1Hash = $sha1Hash.ToLower()
                        return $sha1Hash
                    }
                    else {
                        return "0"
                    }

                }

                # Call the function with your desired path
                Get-DirectoryBlobTree -rootDirectory $remoteTarget
            } -ArgumentList $remoteTarget


            $transmit = $false
            $transmitMessage = "File sync skipped. No files to update. "

            if ($null -eq $remoteBlob) {
                $transmitMessage = "File sync completed. "
                $transmit = $true
            }
            elseif ($remoteBlob.Count -le $localBlob.Count) {
                $transmitMessage = "File sync completed. "
                $transmit = $true
            }
            elseif ($remoteBlob.Count -ge $localBlob.Count) {
                foreach ($blob in $localBlob) {
                    if ($blob.BlobSHA1 -ne ($remoteBlob | Where-object { $_.RelativePath -eq $blob.RelativePath }).BlobSHA1) {
                        $transmitMessage = "File sync completed. "
                        $transmit = $true
                        break
                    }
                }
            }

            if ($transmit) {

                $remoteZip = $remoteTarget + "$localFolderName.zip"

                # Copy zip to remote
                Copy-Item -Path $zipPath -Destination $remoteZip -ToSession $session -Force

                # Unpack, compare, update
                Invoke-Command -Session $session -ScriptBlock {
                    param($zip, $dest)

                    $tempUnpack = "$env:TEMP\temp_" + [Guid]::NewGuid().ToString()
                    New-Item -Path $tempUnpack -ItemType Directory | Out-Null

                    Add-Type -Assembly "System.IO.Compression.FileSystem"
                    [IO.Compression.ZipFile]::ExtractToDirectory($zip, $tempUnpack)
                    Remove-Item $zip -Force

                    # Get hashes of new files
                    $newFiles = Get-ChildItem -Path $tempUnpack -Recurse -File
                    $updates = @{}
                    foreach ($file in $newFiles) {
                        $rel = $file.FullName.Substring($tempUnpack.Length).TrimStart('\')
                        $hash = Get-FileHash -Path $file.FullName -Algorithm SHA256
                        $updates[$rel] = @{ FullPath = $file.FullName; Hash = $hash.Hash }
                    }

                    # Get hashes of existing remote files
                    $existing = @{}
                    if (Test-Path $dest) {
                        Get-ChildItem -Path $dest -Recurse -File | ForEach-Object {
                            $rel = $_.FullName.Substring($dest.Length).TrimStart('\')
                            $hash = Get-FileHash -Path $_.FullName -Algorithm SHA256
                            $existing[$rel] = $hash.Hash
                        }
                    }

                    # Update changed or missing files
                    foreach ($relPath in $updates.Keys) {
                        $file = $updates[$relPath]
                        $shouldCopy = (-not $existing.ContainsKey($relPath)) -or ($file.Hash -ne $existing[$relPath])
                        if ($shouldCopy) {
                            $targetPath = $dest + "\" + $relPath
                            $targetDir = Split-Path $targetPath
                            if (-not (Test-Path $targetDir)) {
                                New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
                            }
                            Copy-Item -Path $file.FullPath -Destination $targetPath -Force
                        }
                    }

                    # Clean temp
                    Remove-Item -Path $tempUnpack -Recurse -Force
                } -ArgumentList $remoteZip, $remoteTarget
            }

            Remove-PSSession $session

            $returnmessage = $transmitMessage + $createListenerResult
        }
        else {
            $returnmessage = "Failed To Connect"
        }
    
        return $returnmessage
    } -ArgumentList $cred, $server, $remoteTarget, $zipPath, $localFolderName, $postNaFlagXml, $preNaFlagXml, $createListener, $localBlob

    $script:sharedState.Value.dataTable["$inncode"] = @{
        Status = "Pending"
        Result = "Launched"
    }

}

# $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$i = 0
$dot = ""
while (($jobs.Count -gt (($jobs | Where-Object { $_.State -eq "Completed" -and $_.HasMoreData -eq $false }).Count + ($jobs | Where-Object { $_.State -eq "Failed" }).Count)) -or !$sharedState.Value.jobPaused) {
    $i++
    foreach ($job in $jobs) {
        try {
            $inncode = $job.Name
            $state = $job.State.ToString()
            switch ($state) {
                "Running" { 
                    $displayResult = "Sync in progress" + $dot
                }
                "Completed" {
                    if ($job.HasMoreData) {
                        $displayResult = Receive-Job $job
                        $jobs = $jobs | Where-Object { $_.State -ne "Completed" }
                        Remove-Job $job
                    }
                    else {} 
                }
                "Failed" {
                    $displayResult = "Failed"     
                    Remove-Job $job
                }
                Default { $displayResult = $null }
            }
            $script:sharedState.Value.dataTable[$inncode] = @{
                Status = $state
                Result = $displayResult
            }
        }
        catch {}
    }
    Start-Sleep -Milliseconds 250

    if ($sharedState.Value.windowClosed -eq $true) {
        break
    }
    if ($i -lt 4) {
        $dot = $dot + "."
    }
    else {
        $i = 0
        $dot = ""
    }
   
}

foreach ($job in $jobs) {
    Stop-Job $job -ErrorAction Ignore
    Remove-Job $job -ErrorAction Ignore
}


Remove-Item $zipPath -Force -ErrorAction Ignore

