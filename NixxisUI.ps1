#Requires -Version 5.1
# NixxisUI.ps1 - WPF GUI for Nixxis Maintenance Operations
# Invoke with:
#   [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
#   irm "https://raw.githubusercontent.com/repnixxis/NixxisUI/main/NixxisUI.ps1" | iex

#region --- Self-Elevation ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $scriptSource = $MyInvocation.MyCommand.Path
    if ($scriptSource) {
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptSource`""
    } else {
        $launchUrl = "https://raw.githubusercontent.com/repnixxis/NixxisUI/main/NixxisUI.ps1"
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; irm '$launchUrl' | iex`""
    }
    exit
}
#endregion

# Force TLS 1.2 — required for GitHub/web downloads on PowerShell 5.1 / Windows Server
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

#region --- XAML ---
[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Nixxis Maintenance Tool"
    Height="740" Width="1160"
    MinHeight="620" MinWidth="960"
    WindowStartupLocation="CenterScreen"
    Background="#1e1e1e"
    FontFamily="Segoe UI"
    FontSize="13">

    <Window.Resources>
        <Style x:Key="PrimaryBtn" TargetType="Button">
            <Setter Property="Background" Value="#0078d4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="Margin" Value="3,3"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#1489e0"/></Trigger>
                            <Trigger Property="IsPressed"   Value="True"><Setter TargetName="bd" Property="Background" Value="#005fa3"/></Trigger>
                            <Trigger Property="IsEnabled"   Value="False">
                                <Setter TargetName="bd" Property="Background" Value="#3a3a3a"/>
                                <Setter Property="Foreground" Value="#555"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="ActionBtn" TargetType="Button" BasedOn="{StaticResource PrimaryBtn}">
            <Setter Property="Background" Value="#2d2d30"/>
            <Setter Property="HorizontalAlignment" Value="Stretch"/>
        </Style>
        <Style x:Key="GreenBtn" TargetType="Button" BasedOn="{StaticResource ActionBtn}">
            <Setter Property="Background" Value="#1e6e42"/>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Foreground" Value="#9cdcfe"/>
            <Setter Property="BorderBrush" Value="#3a3a3a"/>
            <Setter Property="Margin" Value="0,0,0,8"/>
            <Setter Property="Padding" Value="8"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#3c3c3c"/>
            <Setter Property="Foreground" Value="#d4d4d4"/>
            <Setter Property="BorderBrush" Value="#555"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Margin" Value="2,2"/>
            <Setter Property="CaretBrush" Value="White"/>
        </Style>
        <Style TargetType="RadioButton">
            <Setter Property="Foreground" Value="#d4d4d4"/>
            <Setter Property="Margin" Value="0,4"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="#cccccc"/>
            <Setter Property="Padding" Value="2,2"/>
            <Setter Property="FontSize" Value="11"/>
        </Style>
        <Style TargetType="Separator">
            <Setter Property="Background" Value="#3a3a3a"/>
            <Setter Property="Margin" Value="0,5"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="58"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="38"/>
        </Grid.RowDefinitions>

        <!-- HEADER -->
        <Border Grid.Row="0" Background="#252526" BorderBrush="#0078d4" BorderThickness="0,0,0,2">
            <Grid Margin="14,0">
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="&#9881;" Foreground="#0078d4" FontSize="26" VerticalAlignment="Center" Margin="0,0,12,0"/>
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock Text="Nixxis Maintenance Tool" Foreground="White" FontSize="17" FontWeight="SemiBold"/>
                        <TextBlock Text="Automated Update &amp; Deployment  |  v1.4" Foreground="#777" FontSize="11"/>
                    </StackPanel>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,4,0">
                    <Border Background="#1e1e1e" CornerRadius="4" Padding="10,5" Margin="0,0,8,0">
                        <StackPanel Orientation="Horizontal">
                            <Ellipse x:Name="ellServiceDot" Width="9" Height="9" Fill="#888" Margin="0,0,7,0" VerticalAlignment="Center"/>
                            <TextBlock x:Name="tbServiceStatus" Text="Service: Unknown" Foreground="#aaa" FontSize="12" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Border>
                    <Button x:Name="btnRefreshStatus" Content="Refresh" Style="{StaticResource ActionBtn}" Width="80" Height="30" FontSize="12"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- MAIN -->
        <Grid Grid.Row="1" Margin="8,8,8,4">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="300"/>
                <ColumnDefinition Width="6"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- LEFT PANEL -->
            <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                <StackPanel Margin="0,0,4,0">

                    <!-- Working Directory -->
                    <GroupBox Header="  Working Directory  ">
                        <StackPanel>
                            <TextBlock Foreground="#888" FontSize="10" TextWrapping="Wrap" Margin="0,0,0,4"
                                       Text="ZIPs and staged files are saved here. A dated subfolder (YYYYMMDD) is created automatically for each run."/>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="tbWorkDir" Grid.Column="0" Text="C:\NixxisMaintenance\Update" FontSize="11"/>
                                <Button x:Name="btnBrowseWork" Grid.Column="1" Content="..." Width="34" Height="28"
                                        Style="{StaticResource ActionBtn}" Margin="4,2,0,2" FontSize="13"/>
                            </Grid>
                            <TextBlock x:Name="tbRunDir" Text="" Foreground="#569cd6" FontSize="10" Margin="2,4,0,0" TextWrapping="Wrap"/>
                        </StackPanel>
                    </GroupBox>

                    <!-- Source Mode -->
                    <GroupBox Header="  Source Mode  ">
                        <StackPanel>
                            <RadioButton x:Name="rbOnlineAuto"   Content="Online — Auto-Discover latest"  IsChecked="True"/>
                            <RadioButton x:Name="rbOnlineCustom" Content="Online — Custom URLs"/>
                            <RadioButton x:Name="rbOffline"      Content="Offline — Use local ZIP files"/>

                            <StackPanel x:Name="pnlCustomUrls" Visibility="Collapsed" Margin="8,6,0,0">
                                <Label Content="ClientProvisioning.zip URL (blank = auto):"/>
                                <TextBox x:Name="tbCPUrl"  Text="" FontSize="11"/>
                                <Label Content="ClientSoftware.zip URL (blank = auto):"/>
                                <TextBox x:Name="tbCSUrl"  Text="" FontSize="11"/>
                                <Label Content="NCS.zip URL (blank = auto):"/>
                                <TextBox x:Name="tbNCSUrl" Text="" FontSize="11"/>
                            </StackPanel>

                            <StackPanel x:Name="pnlOfflinePath" Visibility="Collapsed" Margin="8,6,0,0">
                                <Label Content="Folder containing the 3 ZIP files:"/>
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox x:Name="tbOfflinePath" Grid.Column="0" Text="" FontSize="11"/>
                                    <Button x:Name="btnBrowseOffline" Grid.Column="1" Content="..." Width="34" Height="28"
                                            Style="{StaticResource ActionBtn}" Margin="4,2,0,2" FontSize="13"/>
                                </Grid>
                                <TextBlock Foreground="#666" FontSize="10" TextWrapping="Wrap" Margin="2,2,0,0"
                                           Text="Required: ClientProvisioning.zip, ClientSoftware.zip, NCS.zip"/>
                            </StackPanel>
                        </StackPanel>
                    </GroupBox>

                    <!-- Full Update -->
                    <GroupBox Header="  Full Update  ">
                        <StackPanel>
                            <Button x:Name="btnRunFull"
                                    Content="&#9658;  Run Full Update"
                                    Style="{StaticResource PrimaryBtn}"
                                    HorizontalAlignment="Stretch"
                                    Height="44" FontSize="15" FontWeight="SemiBold"/>
                            <TextBlock Foreground="#666" FontSize="10" TextWrapping="Wrap" Margin="4,5,4,0"
                                       Text="Download  |  Extract  |  Stop Service  |  Backup  |  Cleanup  |  Deploy"/>
                        </StackPanel>
                    </GroupBox>

                    <!-- Individual Steps -->
                    <GroupBox Header="  Individual Steps  ">
                        <StackPanel>
                            <Button x:Name="btnDownload"     Content="&#9312; Download ZIPs"              Style="{StaticResource ActionBtn}" Height="33"/>
                            <Button x:Name="btnPrepare"      Content="&#9313; Extract / Prepare Files"    Style="{StaticResource ActionBtn}" Height="33"/>
                            <Separator/>
                            <Button x:Name="btnStopService"  Content="&#9314; Stop Nixxis Service"        Style="{StaticResource ActionBtn}" Height="33"/>
                            <Button x:Name="btnBackup"       Content="&#9315; Backup Current Nixxis"      Style="{StaticResource ActionBtn}" Height="33"/>
                            <Button x:Name="btnCleanup"      Content="&#9316; Cleanup Old Files"          Style="{StaticResource ActionBtn}" Height="33"/>
                            <Button x:Name="btnDeploy"       Content="&#9317; Deploy New Files"           Style="{StaticResource ActionBtn}" Height="33"/>
                            <Separator/>
                            <Button x:Name="btnStartService" Content="Start Nixxis Service"               Style="{StaticResource GreenBtn}"  Height="33"/>
                        </StackPanel>
                    </GroupBox>

                    <!-- Quick Actions -->
                    <GroupBox Header="  Quick Actions  ">
                        <UniformGrid Columns="1" Rows="3">
                            <Button x:Name="btnOpenWork"    Content="Open Working Directory" Style="{StaticResource ActionBtn}" Height="29" FontSize="11"/>
                            <Button x:Name="btnOpenLogs"    Content="Open Logs Folder"       Style="{StaticResource ActionBtn}" Height="29" FontSize="11"/>
                            <Button x:Name="btnOpenNixxis"  Content="Browse C:\Nixxis"        Style="{StaticResource ActionBtn}" Height="29" FontSize="11"/>
                        </UniformGrid>
                    </GroupBox>

                </StackPanel>
            </ScrollViewer>

            <!-- SPLITTER -->
            <GridSplitter Grid.Column="1" Width="4" HorizontalAlignment="Stretch" Background="#333"/>

            <!-- RIGHT PANEL - Log -->
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <Border Grid.Row="0" Background="#252526" CornerRadius="4,4,0,0" Padding="10,6" Margin="0,0,0,1">
                    <Grid>
                        <TextBlock Text="ACTIVITY LOG" Foreground="#9cdcfe" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                            <Button x:Name="btnClearLog" Content="Clear"    Style="{StaticResource ActionBtn}" Height="24" Width="55" FontSize="11" Margin="2,0"/>
                            <Button x:Name="btnSaveLog"  Content="Save Log" Style="{StaticResource ActionBtn}" Height="24" Width="68" FontSize="11" Margin="2,0"/>
                        </StackPanel>
                    </Grid>
                </Border>

                <RichTextBox x:Name="rtbLog" Grid.Row="1"
                             Background="#0d1117" Foreground="#c8c8c8"
                             BorderBrush="#3a3a3a" BorderThickness="1"
                             IsReadOnly="True"
                             FontFamily="Consolas,Courier New" FontSize="12"
                             VerticalScrollBarVisibility="Auto"
                             HorizontalScrollBarVisibility="Auto"
                             Padding="8">
                    <RichTextBox.Document>
                        <FlowDocument PageWidth="9999"/>
                    </RichTextBox.Document>
                </RichTextBox>

                <StackPanel Grid.Row="2" Background="#1a1a1a" Margin="0,1,0,0">
                    <ProgressBar x:Name="progressBar" Height="5" Value="0" Maximum="100"
                                 Background="#1a1a1a" Foreground="#0078d4" BorderThickness="0"/>
                    <Grid Margin="8,4">
                        <TextBlock x:Name="tbStatus"  Text="Ready" Foreground="#666" FontSize="11"/>
                        <TextBlock x:Name="tbElapsed" Text=""      Foreground="#555" FontSize="11" HorizontalAlignment="Right"/>
                    </Grid>
                </StackPanel>
            </Grid>
        </Grid>

        <!-- FOOTER -->
        <Border Grid.Row="2" Background="#252526" BorderBrush="#333" BorderThickness="0,1,0,0">
            <Grid Margin="12,0">
                <TextBlock x:Name="tbLogFile" Text="Log: initializing..." Foreground="#444" FontSize="10" VerticalAlignment="Center"/>
            </Grid>
        </Border>
    </Grid>
</Window>
'@
#endregion

#region --- Build Window ---
$reader  = [System.Xml.XmlNodeReader]::new($xaml)
$window  = [Windows.Markup.XamlReader]::Load($reader)
function ctrl($n) { $window.FindName($n) }

$rbOnlineAuto    = ctrl 'rbOnlineAuto'
$rbOnlineCustom  = ctrl 'rbOnlineCustom'
$rbOffline       = ctrl 'rbOffline'
$pnlCustomUrls   = ctrl 'pnlCustomUrls'
$pnlOfflinePath  = ctrl 'pnlOfflinePath'
$tbCPUrl         = ctrl 'tbCPUrl'
$tbCSUrl         = ctrl 'tbCSUrl'
$tbNCSUrl        = ctrl 'tbNCSUrl'
$tbOfflinePath   = ctrl 'tbOfflinePath'
$tbWorkDir       = ctrl 'tbWorkDir'
$tbRunDir        = ctrl 'tbRunDir'
$btnBrowseWork   = ctrl 'btnBrowseWork'
$btnBrowseOffline= ctrl 'btnBrowseOffline'
$btnRunFull      = ctrl 'btnRunFull'
$btnDownload     = ctrl 'btnDownload'
$btnPrepare      = ctrl 'btnPrepare'
$btnStopService  = ctrl 'btnStopService'
$btnBackup       = ctrl 'btnBackup'
$btnCleanup      = ctrl 'btnCleanup'
$btnDeploy       = ctrl 'btnDeploy'
$btnStartService = ctrl 'btnStartService'
$btnRefreshStatus= ctrl 'btnRefreshStatus'
$btnOpenWork     = ctrl 'btnOpenWork'
$btnOpenLogs     = ctrl 'btnOpenLogs'
$btnOpenNixxis   = ctrl 'btnOpenNixxis'
$btnClearLog     = ctrl 'btnClearLog'
$btnSaveLog      = ctrl 'btnSaveLog'
$rtbLog          = ctrl 'rtbLog'
$progressBar     = ctrl 'progressBar'
$tbStatus        = ctrl 'tbStatus'
$tbElapsed       = ctrl 'tbElapsed'
$tbServiceStatus = ctrl 'tbServiceStatus'
$ellServiceDot   = ctrl 'ellServiceDot'
$tbLogFile       = ctrl 'tbLogFile'

$allOpButtons = @($btnRunFull,$btnDownload,$btnPrepare,$btnStopService,$btnBackup,$btnCleanup,$btnDeploy,$btnStartService)
#endregion

#region --- Shared State ---
$sync = [hashtable]::Synchronized(@{
    Queue   = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
    Busy    = $false
    WorkDir = 'C:\NixxisMaintenance\Update'
    RunDir  = ''
    LogLines= [System.Collections.Generic.List[string]]::new()
})

# Log file — in Logs folder alongside WorkDir's parent
$logDate = Get-Date -Format 'yyyyMMdd_HHmmss'
$logDir  = 'C:\NixxisMaintenance\Logs'
$logFile = Join-Path $logDir "NixxisMaintenance_$logDate.log"
if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
$tbLogFile.Text = "Log: $logFile"

# Color map
$colorMap = @{
    INFO    = '#c8c8c8'; OK     = '#4ec9b0'; WARN   = '#dcdcaa'
    ERROR   = '#f44747'; CYAN   = '#9cdcfe'; MAGENTA= '#c586c0'
    GRAY    = '#666666'; HEADER = '#569cd6'
}

# Brush cache
$brushCache = @{}
function Get-Brush($hex) {
    if (-not $brushCache.ContainsKey($hex)) {
        $brushCache[$hex] = [System.Windows.Media.BrushConverter]::new().ConvertFromString($hex)
    }
    $brushCache[$hex]
}
#endregion

#region --- UI Helpers (call ONLY from UI thread) ---

# Add one log line directly to RTB — safe on UI thread, no Dispatcher needed
function Add-LogEntry([string]$Message, [string]$Level = 'INFO') {
    $ts   = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts] $Message"
    $hex  = if ($colorMap.ContainsKey($Level)) { $colorMap[$Level] } else { '#c8c8c8' }

    # Persist to file
    Add-Content -Path $logFile -Value $line -ErrorAction SilentlyContinue
    $sync.LogLines.Add($line)

    # RTB — already on UI thread
    $para          = [System.Windows.Documents.Paragraph]::new()
    $para.Margin   = [System.Windows.Thickness]::new(0)
    $run           = [System.Windows.Documents.Run]::new($line)
    $run.Foreground= Get-Brush $hex
    $para.Inlines.Add($run)
    $rtbLog.Document.Blocks.Add($para)
    $rtbLog.ScrollToEnd()
}

function Set-Status([string]$Text, [int]$Pct = -1) {
    $tbStatus.Text = $Text
    if ($Pct -ge 0) { $progressBar.Value = $Pct }
}

function Set-Busy([bool]$Busy) {
    $sync.Busy = $Busy
    foreach ($b in $allOpButtons) { $b.IsEnabled = -not $Busy }
    if (-not $Busy) { $progressBar.Value = 100 }
}

function Update-ServiceStatus {
    $svc = Get-Service -Name 'crappserver' -ErrorAction SilentlyContinue
    if (-not $svc) {
        $tbServiceStatus.Text = 'Service: Not Found'
        $ellServiceDot.Fill   = Get-Brush '#888888'
    } elseif ($svc.Status -eq 'Running') {
        $tbServiceStatus.Text = 'Service: Running'
        $ellServiceDot.Fill   = Get-Brush '#4ec9b0'
    } else {
        $tbServiceStatus.Text = "Service: $($svc.Status)"
        $ellServiceDot.Fill   = Get-Brush '#dcdcaa'
    }
}

# Compute and display the RunDir label beneath the WorkDir box
function Update-RunDirLabel {
    $dateFolder = Get-Date -Format 'yyyyMMdd'
    $tbRunDir.Text = "Active folder:  $($sync.WorkDir)\$dateFolder"
}
#endregion

#region --- Runspace Job Runner ---
function Start-NixxisJob {
    param([scriptblock]$Work, [string]$JobName = 'Operation')

    if ($sync.Busy) { Add-LogEntry "Another operation is already running." 'WARN'; return }

    # Determine and create the dated run directory
    $dateFolder     = Get-Date -Format 'yyyyMMdd'
    $sync.WorkDir   = $tbWorkDir.Text.Trim()
    $sync.RunDir    = Join-Path $sync.WorkDir $dateFolder
    if (-not (Test-Path $sync.RunDir)) {
        New-Item -Path $sync.RunDir -ItemType Directory -Force | Out-Null
    }
    Update-RunDirLabel
    Add-LogEntry "Working directory: $($sync.RunDir)" 'CYAN'

    Set-Busy $true
    Set-Status "Running: $JobName" 5

    # Snapshot UI values for runspace
    $mode      = if ($rbOnlineAuto.IsChecked)    { 'online-auto' }
                 elseif ($rbOnlineCustom.IsChecked) { 'online-custom' }
                 else { 'offline' }
    $cpUrl     = $tbCPUrl.Text.Trim()
    $csUrl     = $tbCSUrl.Text.Trim()
    $ncsUrl    = $tbNCSUrl.Text.Trim()
    $offPath   = $tbOfflinePath.Text.Trim()
    $runDir    = $sync.RunDir

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('sync',    $sync)
    $rs.SessionStateProxy.SetVariable('logFile', $logFile)
    $rs.SessionStateProxy.SetVariable('mode',    $mode)
    $rs.SessionStateProxy.SetVariable('cpUrl',   $cpUrl)
    $rs.SessionStateProxy.SetVariable('csUrl',   $csUrl)
    $rs.SessionStateProxy.SetVariable('ncsUrl',  $ncsUrl)
    $rs.SessionStateProxy.SetVariable('offlinePath', $offPath)
    $rs.SessionStateProxy.SetVariable('runDir',  $runDir)

    $ps = [powershell]::Create()
    $ps.Runspace = $rs

    # Logging helper available inside every runspace scriptblock
    $logHelper = {
        function Write-BgLog {
            param([string]$Message, [string]$Level = 'INFO')
            $sync.Queue.Enqueue(@{ Message = $Message; Level = $Level })
        }
    }

    $ps.AddScript($logHelper) | Out-Null
    $ps.AddScript($Work)       | Out-Null

    $startTime = [datetime]::Now
    $handle    = $ps.BeginInvoke()

    # DispatcherTimer — runs on UI thread, safe to touch controls directly
    $timer          = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [timespan]::FromMilliseconds(200)

    # Use a script-level variable to hold closure state
    $script:_job = @{ ps = $ps; rs = $rs; handle = $handle; timer = $timer; start = $startTime; name = $JobName }

    $timer.Add_Tick({
        $j = $script:_job

        # Drain log queue — already on UI thread, call Add-LogEntry directly
        $entry = $null
        while ($sync.Queue.TryDequeue([ref]$entry)) {
            Add-LogEntry $entry.Message $entry.Level
        }

        # Update elapsed
        $tbElapsed.Text = "Elapsed: $(([datetime]::Now - $j.start).ToString('mm\:ss'))"

        # Check completion
        if ($j.handle.IsCompleted) {
            $j.timer.Stop()

            # Final drain
            $entry = $null
            while ($sync.Queue.TryDequeue([ref]$entry)) {
                Add-LogEntry $entry.Message $entry.Level
            }

            # Collect runspace errors
            if ($j.ps.HadErrors) {
                foreach ($err in $j.ps.Streams.Error) {
                    Add-LogEntry "Error: $err" 'ERROR'
                }
            }
            try { $j.ps.EndInvoke($j.handle) } catch { Add-LogEntry "Job exception: $_" 'ERROR' }
            $j.ps.Dispose()
            $j.rs.Dispose()

            Set-Busy $false
            Set-Status "Done: $($j.name)" 100
            Update-ServiceStatus
        }
    })
    $timer.Start()
}
#endregion

#region --- Operation Scriptblocks ---

$sbDownload = {
    Write-BgLog '=== DOWNLOAD PHASE ===' 'HEADER'
    Write-BgLog "Saving to: $runDir" 'CYAN'

    $baseUrl = 'http://update.nixxis.net'

    function Get-LatestFolder([string]$url, [string]$desc) {
        Write-BgLog "Scanning: $url" 'GRAY'
        $resp    = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30
        $folders = @()
        foreach ($pat in @('href="([^"]+/)"', '(v\d+\.\d+)', '(\d+\.\d+\.\d+)')) {
            [regex]::Matches($resp.Content, $pat) | ForEach-Object {
                $n = $_.Groups[1].Value.TrimEnd('/')
                if ($n -notmatch '^(\.\.|\.|icons|cgi-bin)$' -and $n -notin $folders) { $folders += $n }
            }
        }
        if (-not $folders) { throw "No folders found at $url" }
        $latest = ($folders | Sort-Object {
            $v = $_ -replace '[^\d\.]', ''
            try { [version]$v } catch { [version]'0.0' }
        } | Select-Object -Last 1)
        Write-BgLog "Latest $desc`: $latest" 'OK'
        return $latest
    }

    function Get-ZipFile([string]$url, [string]$dest) {
        Write-BgLog "Downloading: $url" 'CYAN'
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($url, $dest)
        $wc.Dispose()
        $sz = [math]::Round((Get-Item $dest).Length / 1MB, 2)
        Write-BgLog "  Saved $(Split-Path $dest -Leaf) ($sz MB)" 'OK'
    }

    if ($mode -eq 'offline') {
        $src = if ($offlinePath) { $offlinePath } else { $runDir }
        Write-BgLog "Offline mode — source: $src" 'MAGENTA'
        foreach ($zip in @('ClientProvisioning.zip','ClientSoftware.zip','NCS.zip')) {
            $f = Join-Path $src $zip
            if (-not (Test-Path $f)) { throw "Missing offline file: $f" }
            Write-BgLog "  Found $zip" 'OK'
            if ((Resolve-Path $src).Path -ne (Resolve-Path $runDir).Path) {
                Copy-Item $f $runDir -Force
                Write-BgLog "  Copied to run folder" 'GRAY'
            }
        }
    } else {
        $resolvedCP  = $cpUrl
        $resolvedCS  = $csUrl
        $resolvedNCS = $ncsUrl

        if (-not $resolvedCP -or -not $resolvedCS -or -not $resolvedNCS) {
            $latestVer  = Get-LatestFolder $baseUrl 'version'
            $versionUrl = "$baseUrl/$latestVer"

            if (-not $resolvedCP -or -not $resolvedCS) {
                $clientUrl = "$versionUrl/Client"
                $latestCli = Get-LatestFolder $clientUrl 'client build'
                if (-not $resolvedCP)  { $resolvedCP = "$clientUrl/$latestCli/ClientProvisioning.zip" }
                if (-not $resolvedCS)  { $resolvedCS = "$clientUrl/$latestCli/ClientSoftware.zip" }
            }
            if (-not $resolvedNCS) {
                $serverUrl  = "$versionUrl/Server"
                $latestSrv  = Get-LatestFolder $serverUrl 'server build'
                $resolvedNCS = "$serverUrl/$latestSrv/NCS.zip"
            }
        }

        Write-BgLog "ClientProvisioning : $resolvedCP"  'GRAY'
        Write-BgLog "ClientSoftware     : $resolvedCS"  'GRAY'
        Write-BgLog "NCS                : $resolvedNCS" 'GRAY'

        Get-ZipFile $resolvedCP  (Join-Path $runDir 'ClientProvisioning.zip')
        Get-ZipFile $resolvedCS  (Join-Path $runDir 'ClientSoftware.zip')
        Get-ZipFile $resolvedNCS (Join-Path $runDir 'NCS.zip')
    }
    Write-BgLog "Download phase complete. Files in: $runDir" 'OK'
}

$sbPrepare = {
    Write-BgLog '=== PREPARE / EXTRACT PHASE ===' 'HEADER'
    Write-BgLog "Source folder: $runDir" 'CYAN'

    foreach ($z in @('NCS.zip','ClientProvisioning.zip','ClientSoftware.zip')) {
        if (-not (Test-Path (Join-Path $runDir $z))) { throw "Required ZIP not found in run folder: $z`nExpected in: $runDir" }
    }

    $appServer = Join-Path $runDir 'NixxisApplicationServer'
    if (Test-Path $appServer) {
        Write-BgLog "Removing old NixxisApplicationServer staging folder..." 'WARN'
        Remove-Item $appServer -Recurse -Force
    }
    New-Item $appServer -ItemType Directory -Force | Out-Null
    Write-BgLog "Created staging folder: $appServer" 'OK'

    Write-BgLog "Extracting NCS.zip..." 'CYAN'
    Expand-Archive (Join-Path $runDir 'NCS.zip') -DestinationPath $appServer -Force
    Write-BgLog "  Done" 'OK'

    $csDest = Join-Path $appServer 'ClientSoftware'
    New-Item $csDest -ItemType Directory -Force | Out-Null
    Write-BgLog "Extracting ClientSoftware.zip..." 'CYAN'
    Expand-Archive (Join-Path $runDir 'ClientSoftware.zip') -DestinationPath $csDest -Force
    Write-BgLog "  Done" 'OK'

    $provPath   = Join-Path $appServer 'CrAppServer\provisioning'
    $provClient = Join-Path $provPath 'client'
    New-Item $provClient -ItemType Directory -Force | Out-Null
    Copy-Item (Join-Path $runDir 'ClientSoftware.zip') $provPath -Force

    Write-BgLog "Extracting ClientProvisioning.zip..." 'CYAN'
    Expand-Archive (Join-Path $runDir 'ClientProvisioning.zip') -DestinationPath $provClient -Force
    Write-BgLog "  Done" 'OK'

    $settingsSrc = Join-Path $provClient 'settings'
    if (Test-Path $settingsSrc) {
        Move-Item $settingsSrc $provPath -Force
        Write-BgLog "Moved settings folder to provisioning root" 'OK'
    }
    Write-BgLog "Preparation phase complete. Staged in: $appServer" 'OK'
}

$sbStopService = {
    Write-BgLog '=== STOP SERVICE PHASE ===' 'HEADER'

    $clientProc = Get-Process 'nixxisclientdesktop' -ErrorAction SilentlyContinue
    if ($clientProc) {
        Write-BgLog "Killing nixxisclientdesktop.exe (PID $($clientProc.Id))..." 'WARN'
        & taskkill /IM nixxisclientdesktop.exe /F 2>&1 | Out-Null
        Start-Sleep 2
        if (Get-Process 'nixxisclientdesktop' -ErrorAction SilentlyContinue) {
            Write-BgLog "WARNING: nixxisclientdesktop still running." 'WARN'
        } else { Write-BgLog "nixxisclientdesktop terminated." 'OK' }
    } else { Write-BgLog "nixxisclientdesktop not running." 'GRAY' }

    Stop-Service -Name 'crappserver' -Force -ErrorAction SilentlyContinue
    $max = 60; $i = 0
    do {
        $svc  = Get-Service 'crappserver' -ErrorAction SilentlyContinue
        $proc = Get-Process 'crappserver' -ErrorAction SilentlyContinue
        if ((-not $svc -or $svc.Status -eq 'Stopped') -and -not $proc) {
            Start-Sleep 3
            if (-not (Get-Process 'crappserver' -ErrorAction SilentlyContinue)) {
                Write-BgLog "Crappserver fully stopped." 'OK'; break
            }
        }
        $i++
        Write-BgLog "Waiting for service to stop... attempt $i / $max" 'WARN'
        Stop-Service -Name 'crappserver' -Force -ErrorAction SilentlyContinue
        Start-Sleep 2
    } until ($i -ge $max)
    if ($i -ge $max) { throw "Service did not stop within 120 seconds." }
    Write-BgLog 'Service stop phase complete.' 'OK'
}

$sbBackup = {
    Write-BgLog '=== BACKUP PHASE ===' 'HEADER'
    $dateFolder = Get-Date -Format 'yyyyMMdd'
    $year       = (Get-Date).Year.ToString()
    $backupBase = 'C:\NixxisMaintenance\BackUp'
    $dest       = Join-Path $backupBase "$year\$dateFolder"

    foreach ($p in @($dest, (Join-Path $dest 'NMS'), (Join-Path $dest 'SQL'))) {
        New-Item $p -ItemType Directory -Force | Out-Null
    }
    @('NMS1','NMS2','NMS3','NMS4') | ForEach-Object {
        New-Item (Join-Path $dest "NMS\$_") -ItemType Directory -Force | Out-Null
    }
    Write-BgLog "Backup destination: $dest" 'CYAN'

    foreach ($src in @('C:\Nixxis\CrAppServer', 'C:\Nixxis\ClientSoftware')) {
        if (Test-Path $src) {
            $name = Split-Path $src -Leaf
            Write-BgLog "Backing up $name..." 'CYAN'
            Copy-Item $src $dest -Recurse -Force
            Write-BgLog "  Done" 'OK'
        } else { Write-BgLog "$src not found — skipping." 'WARN' }
    }
    Write-BgLog 'Backup phase complete.' 'OK'
}

$sbCleanup = {
    Write-BgLog '=== CLEANUP PHASE ===' 'HEADER'
    $base = 'C:\Nixxis\CrAppServer'
    if (-not (Test-Path $base)) {
        Write-BgLog "CrAppServer not found at $base — nothing to clean." 'WARN'
    } else {
        $files = @(
            "$base\AdminLink.dll","$base\agsXMPP.dll","$base\CRAppServerInterfaces.dll",
            "$base\CrAppServerReferences.dll","$base\CRShared.dll",
            "$base\Microsoft.ReportViewer.Common.dll","$base\Microsoft.ReportViewer.WinForms.dll",
            "$base\Nixxis.Install.dll","$base\NixxisClientReferences.dll","$base\RestSharp.dll",
            "$base\SampleConfigurationFiles.dll","$base\SocialMediaElements.dll","$base\SFAgent.dll",
            "$base\System.Windows.Controls.DataVisualization.Toolkit.dll","$base\Twitterizer2.dll",
            "$base\UscNetTools.dll","$base\CRReportingServer.exe","$base\ManitermSettings.dll",
            "$base\NCS.TranscriberSettings.dll","$base\NixxisSalesForceClient.dll",
            "$base\TicketControllerSettings.dll","$base\WebChatSettings.dll",
            "$base\Plugins\ManitermSettings.dll","$base\Plugins\TicketsControllerSettings.dll",
            "$base\Default_reporting\StateMachineContactList.bin",
            "$base\TicketsControllerSettings.dll","$base\TicketsController.dll"
        )
        foreach ($f in $files) {
            if (Test-Path $f) {
                try { Remove-Item $f -Force; Write-BgLog "Deleted: $(Split-Path $f -Leaf)" 'GRAY' }
                catch { Write-BgLog "Could not delete $(Split-Path $f -Leaf): $_" 'WARN' }
            }
        }
        Get-ChildItem $base -Filter '*.pdb' -File -ErrorAction SilentlyContinue | ForEach-Object {
            try { Remove-Item $_.FullName -Force; Write-BgLog "Deleted PDB: $($_.Name)" 'GRAY' }
            catch { Write-BgLog "Could not delete PDB $($_.Name): $_" 'WARN' }
        }
        $prov = Join-Path $base 'Provisioning'
        if (Test-Path $prov) {
            Get-ChildItem $prov | Where-Object { $_.Name -ne 'Settings' } | ForEach-Object {
                Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                Write-BgLog "Removed from Provisioning: $($_.Name)" 'GRAY'
            }
        }
    }
    $cs = 'C:\Nixxis\ClientSoftware'
    if (Test-Path $cs) {
        Remove-Item $cs -Recurse -Force -ErrorAction SilentlyContinue
        Write-BgLog "Removed ClientSoftware folder" 'OK'
    } else { Write-BgLog "ClientSoftware not found — skipping." 'WARN' }
    Write-BgLog 'Cleanup phase complete.' 'OK'
}

$sbDeploy = {
    Write-BgLog '=== DEPLOY PHASE ===' 'HEADER'
    $appServer = Join-Path $runDir 'NixxisApplicationServer'
    if (-not (Test-Path $appServer)) { throw "NixxisApplicationServer not found in run folder. Run Prepare step first.`nExpected: $appServer" }

    foreach ($f in @(
        @{ S='ClientSoftware';   D='C:\Nixxis\ClientSoftware'   },
        @{ S='CrAppServer';      D='C:\Nixxis\CrAppServer'      },
        @{ S='MediaServer';      D='C:\Nixxis\MediaServer'       },
        @{ S='Reporting';        D='C:\Nixxis\Reporting'         },
        @{ S='SampleConfigFiles';D='C:\Nixxis\SampleConfigFiles' },
        @{ S='SoundsSamples';    D='C:\Nixxis\SoundsSamples'     }
    )) {
        $src = Join-Path $appServer $f.S
        if (Test-Path $src) {
            Write-BgLog "Deploying $($f.S)  ->  $($f.D)..." 'CYAN'
            if (-not (Test-Path $f.D)) { New-Item $f.D -ItemType Directory -Force | Out-Null }
            Copy-Item "$src\*" $f.D -Recurse -Force
            Write-BgLog "  Done" 'OK'
        } else { Write-BgLog "Source not found in staging: $($f.S) — skipping." 'WARN' }
    }
    Write-BgLog 'Deploy phase complete.' 'OK'
}

$sbStartService = {
    Write-BgLog '=== START SERVICE ===' 'HEADER'
    $svc = Get-Service 'crappserver' -ErrorAction SilentlyContinue
    if (-not $svc) { throw "Service 'crappserver' not found on this machine." }
    Start-Service 'crappserver'
    Start-Sleep 3
    $svc.Refresh()
    Write-BgLog "Service status: $($svc.Status)" $(if ($svc.Status -eq 'Running') { 'OK' } else { 'WARN' })
}

# Full update chains all phases in sequence
$sbFullUpdate = [scriptblock]::Create(
    $sbDownload.ToString()     + "`n" +
    $sbPrepare.ToString()      + "`n" +
    $sbStopService.ToString()  + "`n" +
    $sbBackup.ToString()       + "`n" +
    $sbCleanup.ToString()      + "`n" +
    $sbDeploy.ToString()       + "`n" +
    "Write-BgLog '=== FULL UPDATE COMPLETE ===' 'OK'" + "`n" +
    "Write-BgLog 'Click Start Service when ready.' 'CYAN'"
)
#endregion

#region --- Event Handlers ---

# Mode radio toggles
$rbOnlineCustom.Add_Checked({ $pnlCustomUrls.Visibility  = 'Visible';   $pnlOfflinePath.Visibility = 'Collapsed' })
$rbOffline.Add_Checked({      $pnlOfflinePath.Visibility = 'Visible';   $pnlCustomUrls.Visibility  = 'Collapsed' })
$rbOnlineAuto.Add_Checked({   $pnlCustomUrls.Visibility  = 'Collapsed'; $pnlOfflinePath.Visibility = 'Collapsed' })

# WorkDir picker
$btnBrowseWork.Add_Click({
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description  = 'Select the NixxisMaintenance working directory'
    $dlg.SelectedPath = $tbWorkDir.Text
    if ($dlg.ShowDialog() -eq 'OK') {
        $tbWorkDir.Text = $dlg.SelectedPath
        $sync.WorkDir   = $dlg.SelectedPath
        Update-RunDirLabel
    }
})

# WorkDir text change — keep sync live
$tbWorkDir.Add_TextChanged({
    $sync.WorkDir = $tbWorkDir.Text.Trim()
    Update-RunDirLabel
})

# Offline folder picker
$btnBrowseOffline.Add_Click({
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description = 'Select folder containing the 3 Nixxis ZIP files'
    if ($dlg.ShowDialog() -eq 'OK') { $tbOfflinePath.Text = $dlg.SelectedPath }
})

$btnRefreshStatus.Add_Click({ Update-ServiceStatus })

$btnClearLog.Add_Click({
    $rtbLog.Document.Blocks.Clear()
    $sync.LogLines.Clear()
})
$btnSaveLog.Add_Click({
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title    = 'Save Log File'
    $dlg.Filter   = 'Text Files (*.txt)|*.txt|All Files (*.*)|*.*'
    $dlg.FileName = "NixxisLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    if ($dlg.ShowDialog() -eq $true) {
        $sync.LogLines | Set-Content $dlg.FileName -Encoding UTF8
        Add-LogEntry "Log saved: $($dlg.FileName)" 'OK'
    }
})

$btnOpenWork.Add_Click({
    $p = $tbWorkDir.Text.Trim()
    if (Test-Path $p) { Start-Process explorer $p }
    else { Add-LogEntry "Working directory not found: $p — it will be created on first run." 'WARN' }
})
$btnOpenLogs.Add_Click({
    if (Test-Path $logDir) { Start-Process explorer $logDir }
    else { Add-LogEntry "Logs folder not yet created." 'WARN' }
})
$btnOpenNixxis.Add_Click({
    if (Test-Path 'C:\Nixxis') { Start-Process explorer 'C:\Nixxis' }
    else { Add-LogEntry "C:\Nixxis not found on this machine." 'WARN' }
})

# Operation buttons
$btnRunFull.Add_Click({     Add-LogEntry '==== FULL UPDATE ====' 'HEADER';   Set-Status 'Running full update...' 0;  Start-NixxisJob $sbFullUpdate    'Full Update'   })
$btnDownload.Add_Click({    Add-LogEntry '==== DOWNLOAD ====' 'HEADER';      Set-Status 'Downloading ZIPs...' 0;    Start-NixxisJob $sbDownload      'Download'      })
$btnPrepare.Add_Click({     Add-LogEntry '==== PREPARE ====' 'HEADER';       Set-Status 'Preparing files...' 0;     Start-NixxisJob $sbPrepare       'Prepare'       })
$btnStopService.Add_Click({ Add-LogEntry '==== STOP SERVICE ====' 'HEADER';  Set-Status 'Stopping service...' 0;    Start-NixxisJob $sbStopService   'Stop Service'  })
$btnBackup.Add_Click({      Add-LogEntry '==== BACKUP ====' 'HEADER';        Set-Status 'Backing up...' 0;         Start-NixxisJob $sbBackup        'Backup'        })
$btnCleanup.Add_Click({     Add-LogEntry '==== CLEANUP ====' 'HEADER';       Set-Status 'Cleaning up...' 0;        Start-NixxisJob $sbCleanup       'Cleanup'       })
$btnDeploy.Add_Click({      Add-LogEntry '==== DEPLOY ====' 'HEADER';        Set-Status 'Deploying...' 0;          Start-NixxisJob $sbDeploy        'Deploy'        })
$btnStartService.Add_Click({Add-LogEntry '==== START SERVICE ====' 'HEADER'; Set-Status 'Starting service...' 0;   Start-NixxisJob $sbStartService  'Start Service' })
#endregion

#region --- Startup ---
Update-RunDirLabel
Add-LogEntry 'Nixxis Maintenance Tool ready.' 'HEADER'
Add-LogEntry "User: $env:USERNAME  |  Host: $env:COMPUTERNAME" 'GRAY'
Add-LogEntry "Log file: $logFile" 'GRAY'
Add-LogEntry "Default working directory: $($sync.WorkDir)" 'CYAN'
Add-LogEntry "A dated subfolder (YYYYMMDD) will be created in that directory on each run." 'GRAY'
Add-LogEntry '' 'INFO'
Update-ServiceStatus

$window.ShowDialog() | Out-Null
#endregion
