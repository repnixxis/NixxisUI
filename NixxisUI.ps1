#Requires -Version 5.1
# NixxisUI.ps1 - WPF GUI for Nixxis Maintenance Operations
# Invoke with: irm "https://raw.githubusercontent.com/repnixxis/NixxisUI/main/NixxisUI.ps1" | iex

#region --- Self-Elevation ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $scriptSource = $MyInvocation.MyCommand.Path
    if ($scriptSource) {
        # Launched from file — re-launch elevated
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptSource`""
    } else {
        # Launched via irm | iex — re-launch with the URL
        $launchUrl = "https://raw.githubusercontent.com/repnixxis/NixxisUI/main/NixxisUI.ps1"
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; irm '$launchUrl' | iex`""
    }
    exit
}
#endregion

# Force TLS 1.2 — required for GitHub/web downloads on PowerShell 5.1 / Windows Server
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

#region --- XAML Layout ---
[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Nixxis Maintenance Tool"
    Height="720" Width="1150"
    MinHeight="600" MinWidth="950"
    WindowStartupLocation="CenterScreen"
    Background="#1e1e1e"
    FontFamily="Segoe UI"
    FontSize="13"
    WindowStyle="SingleBorderWindow">

    <Window.Resources>
        <!-- Primary blue button -->
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
                            <Trigger Property="IsEnabled"   Value="False"><Setter TargetName="bd" Property="Background" Value="#3a3a3a"/><Setter Property="Foreground" Value="#666666"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Muted action button -->
        <Style x:Key="ActionBtn" TargetType="Button" BasedOn="{StaticResource PrimaryBtn}">
            <Setter Property="Background" Value="#2d2d30"/>
            <Setter Property="HorizontalAlignment" Value="Stretch"/>
        </Style>
        <!-- Green button -->
        <Style x:Key="GreenBtn" TargetType="Button" BasedOn="{StaticResource ActionBtn}">
            <Setter Property="Background" Value="#1e6e42"/>
        </Style>
        <!-- GroupBox -->
        <Style TargetType="GroupBox">
            <Setter Property="Foreground" Value="#9cdcfe"/>
            <Setter Property="BorderBrush" Value="#3a3a3a"/>
            <Setter Property="Margin" Value="0,0,0,8"/>
            <Setter Property="Padding" Value="8"/>
        </Style>
        <!-- TextBox -->
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#3c3c3c"/>
            <Setter Property="Foreground" Value="#d4d4d4"/>
            <Setter Property="BorderBrush" Value="#555"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Margin" Value="2,2"/>
            <Setter Property="CaretBrush" Value="White"/>
        </Style>
        <!-- RadioButton -->
        <Style TargetType="RadioButton">
            <Setter Property="Foreground" Value="#d4d4d4"/>
            <Setter Property="Margin" Value="0,4"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <!-- Label -->
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="#cccccc"/>
            <Setter Property="Padding" Value="2,2"/>
            <Setter Property="FontSize" Value="11"/>
        </Style>
        <!-- Separator -->
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

        <!-- ===== HEADER ===== -->
        <Border Grid.Row="0" Background="#252526" BorderBrush="#0078d4" BorderThickness="0,0,0,2">
            <Grid Margin="14,0">
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="⚙" Foreground="#0078d4" FontSize="26" VerticalAlignment="Center" Margin="0,0,12,0"/>
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock Text="Nixxis Maintenance Tool" Foreground="White" FontSize="17" FontWeight="SemiBold"/>
                        <TextBlock Text="Automated Update &amp; Deployment  •  v1.4" Foreground="#777" FontSize="11"/>
                    </StackPanel>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,4,0">
                    <Border Background="#1e1e1e" CornerRadius="4" Padding="10,5" Margin="0,0,8,0">
                        <StackPanel Orientation="Horizontal">
                            <Ellipse x:Name="ellServiceDot" Width="9" Height="9" Fill="#888" Margin="0,0,7,0" VerticalAlignment="Center"/>
                            <TextBlock x:Name="tbServiceStatus" Text="Service: Unknown" Foreground="#aaa" FontSize="12" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Border>
                    <Button x:Name="btnRefreshStatus" Content="↺  Refresh" Style="{StaticResource ActionBtn}" Width="90" Height="30" FontSize="12"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- ===== MAIN CONTENT ===== -->
        <Grid Grid.Row="1" Margin="8,8,8,4">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="295"/>
                <ColumnDefinition Width="6"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- LEFT PANEL -->
            <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                <StackPanel Margin="0,0,4,0">

                    <!-- Source Mode -->
                    <GroupBox Header="  Source Mode  ">
                        <StackPanel>
                            <RadioButton x:Name="rbOnlineAuto"   Content="Online — Auto-Discover latest"  IsChecked="True"/>
                            <RadioButton x:Name="rbOnlineCustom" Content="Online — Custom URLs"/>
                            <RadioButton x:Name="rbOffline"      Content="Offline — Use local ZIP files"/>

                            <!-- Custom URL panel -->
                            <StackPanel x:Name="pnlCustomUrls" Visibility="Collapsed" Margin="8,6,0,0">
                                <Label Content="ClientProvisioning.zip URL (blank = auto):"/>
                                <TextBox x:Name="tbCPUrl"  Text="" FontSize="11"/>
                                <Label Content="ClientSoftware.zip URL (blank = auto):"/>
                                <TextBox x:Name="tbCSUrl"  Text="" FontSize="11"/>
                                <Label Content="NCS.zip URL (blank = auto):"/>
                                <TextBox x:Name="tbNCSUrl" Text="" FontSize="11"/>
                            </StackPanel>

                            <!-- Offline path panel -->
                            <StackPanel x:Name="pnlOfflinePath" Visibility="Collapsed" Margin="8,6,0,0">
                                <Label Content="Folder containing the 3 ZIP files:"/>
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox x:Name="tbOfflinePath" Grid.Column="0" Text="" FontSize="11"/>
                                    <Button x:Name="btnBrowse" Grid.Column="1" Content="…" Width="34" Height="28"
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
                                    Content="▶  Run Full Update"
                                    Style="{StaticResource PrimaryBtn}"
                                    HorizontalAlignment="Stretch"
                                    Height="44" FontSize="15" FontWeight="SemiBold"/>
                            <TextBlock Foreground="#666" FontSize="10" TextWrapping="Wrap" Margin="4,5,4,0"
                                       Text="Download → Extract → Stop Service → Backup → Cleanup → Deploy"/>
                        </StackPanel>
                    </GroupBox>

                    <!-- Individual Steps -->
                    <GroupBox Header="  Individual Steps  ">
                        <StackPanel>
                            <Button x:Name="btnDownload"     Content="① Download ZIPs"              Style="{StaticResource ActionBtn}" Height="33"/>
                            <Button x:Name="btnPrepare"      Content="② Extract / Prepare Files"    Style="{StaticResource ActionBtn}" Height="33"/>
                            <Separator/>
                            <Button x:Name="btnStopService"  Content="③ Stop Nixxis Service"        Style="{StaticResource ActionBtn}" Height="33"/>
                            <Button x:Name="btnBackup"       Content="④ Backup Current Nixxis"      Style="{StaticResource ActionBtn}" Height="33"/>
                            <Button x:Name="btnCleanup"      Content="⑤ Cleanup Old Files"          Style="{StaticResource ActionBtn}" Height="33"/>
                            <Button x:Name="btnDeploy"       Content="⑥ Deploy New Files"           Style="{StaticResource ActionBtn}" Height="33"/>
                            <Separator/>
                            <Button x:Name="btnStartService" Content="Start Nixxis Service"         Style="{StaticResource GreenBtn}"  Height="33"/>
                        </StackPanel>
                    </GroupBox>

                    <!-- Quick Actions -->
                    <GroupBox Header="  Quick Actions  ">
                        <UniformGrid Columns="1" Rows="3">
                            <Button x:Name="btnOpenBackup"  Content="📂 Open Backup Folder" Style="{StaticResource ActionBtn}" Height="29" FontSize="11"/>
                            <Button x:Name="btnOpenLogs"    Content="📋 Open Logs Folder"   Style="{StaticResource ActionBtn}" Height="29" FontSize="11"/>
                            <Button x:Name="btnOpenNixxis"  Content="📁 Browse C:\Nixxis"   Style="{StaticResource ActionBtn}" Height="29" FontSize="11"/>
                        </UniformGrid>
                    </GroupBox>

                </StackPanel>
            </ScrollViewer>

            <!-- SPLITTER -->
            <GridSplitter Grid.Column="1" Width="4" HorizontalAlignment="Stretch" Background="#333"/>

            <!-- RIGHT PANEL — Log -->
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <Border Grid.Row="0" Background="#252526" CornerRadius="4,4,0,0" Padding="10,6" Margin="0,0,0,1">
                    <Grid>
                        <TextBlock Text="ACTIVITY LOG" Foreground="#9cdcfe" FontSize="11"
                                   FontWeight="SemiBold" VerticalAlignment="Center"/>
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
                        <TextBlock x:Name="tbStatus" Text="Ready — select a mode and click Run Full Update or an individual step."
                                   Foreground="#666" FontSize="11"/>
                        <TextBlock x:Name="tbElapsed" Text="" Foreground="#555" FontSize="11" HorizontalAlignment="Right"/>
                    </Grid>
                </StackPanel>
            </Grid>
        </Grid>

        <!-- ===== FOOTER ===== -->
        <Border Grid.Row="2" Background="#252526" BorderBrush="#333" BorderThickness="0,1,0,0">
            <Grid Margin="12,0">
                <TextBlock Text="Requires Administrator • Logs: C:\NixxisMaintenance\Logs • Backup: C:\NixxisMaintenance\BackUp"
                           Foreground="#444" FontSize="10" VerticalAlignment="Center"/>
                <TextBlock x:Name="tbLogFile" Text="" Foreground="#444" FontSize="10"
                           VerticalAlignment="Center" HorizontalAlignment="Right"/>
            </Grid>
        </Border>
    </Grid>
</Window>
'@
#endregion

#region --- Build Window ---
$reader   = [System.Xml.XmlNodeReader]::new($xaml)
$window   = [Windows.Markup.XamlReader]::Load($reader)

# Helper to find controls
function ctrl($name) { $window.FindName($name) }

$rbOnlineAuto   = ctrl 'rbOnlineAuto'
$rbOnlineCustom = ctrl 'rbOnlineCustom'
$rbOffline      = ctrl 'rbOffline'
$pnlCustomUrls  = ctrl 'pnlCustomUrls'
$pnlOfflinePath = ctrl 'pnlOfflinePath'
$tbCPUrl        = ctrl 'tbCPUrl'
$tbCSUrl        = ctrl 'tbCSUrl'
$tbNCSUrl       = ctrl 'tbNCSUrl'
$tbOfflinePath  = ctrl 'tbOfflinePath'
$btnBrowse      = ctrl 'btnBrowse'
$btnRunFull     = ctrl 'btnRunFull'
$btnDownload    = ctrl 'btnDownload'
$btnPrepare     = ctrl 'btnPrepare'
$btnStopService = ctrl 'btnStopService'
$btnBackup      = ctrl 'btnBackup'
$btnCleanup     = ctrl 'btnCleanup'
$btnDeploy      = ctrl 'btnDeploy'
$btnStartService= ctrl 'btnStartService'
$btnRefreshStatus=ctrl 'btnRefreshStatus'
$btnOpenBackup  = ctrl 'btnOpenBackup'
$btnOpenLogs    = ctrl 'btnOpenLogs'
$btnOpenNixxis  = ctrl 'btnOpenNixxis'
$btnClearLog    = ctrl 'btnClearLog'
$btnSaveLog     = ctrl 'btnSaveLog'
$rtbLog         = ctrl 'rtbLog'
$progressBar    = ctrl 'progressBar'
$tbStatus       = ctrl 'tbStatus'
$tbElapsed      = ctrl 'tbElapsed'
$tbServiceStatus= ctrl 'tbServiceStatus'
$ellServiceDot  = ctrl 'ellServiceDot'
$tbLogFile      = ctrl 'tbLogFile'
#endregion

#region --- Shared State ---
$sync = [hashtable]::Synchronized(@{
    Queue    = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
    Busy     = $false
    Abort    = $false
    LogLines = [System.Collections.Generic.List[string]]::new()
})

# Log paths
$logDate    = Get-Date -Format 'yyyyMMdd_HHmmss'
$logDir     = 'C:\NixxisMaintenance\Logs'
$logFile    = Join-Path $logDir "NixxisMaintenance_$logDate.log"
if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
$window.FindName('tbLogFile').Text = "Log: $logFile"

# Color map: level → hex color
$colorMap = @{
    INFO    = '#c8c8c8'
    OK      = '#4ec9b0'
    WARN    = '#dcdcaa'
    ERROR   = '#f44747'
    CYAN    = '#9cdcfe'
    MAGENTA = '#c586c0'
    GRAY    = '#666666'
    HEADER  = '#569cd6'
}
#endregion

#region --- UI Helper Functions (must run on UI thread) ---
function Write-UILog {
    param([string]$Message, [string]$Level = 'INFO')
    $ts  = Get-Date -Format 'HH:mm:ss'
    $color = if ($colorMap.ContainsKey($Level)) { $colorMap[$Level] } else { '#c8c8c8' }
    $full = "[$ts] $Message"

    # Persist to file
    Add-Content -Path $logFile -Value $full -ErrorAction SilentlyContinue
    $sync.LogLines.Add($full)

    # UI — must be on dispatcher thread
    $window.Dispatcher.Invoke([action]{
        $doc  = $rtbLog.Document
        $para = [System.Windows.Documents.Paragraph]::new()
        $run  = [System.Windows.Documents.Run]::new($full)
        $run.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($color)
        $para.Inlines.Add($run)
        $para.Margin = [System.Windows.Thickness]::new(0)
        $doc.Blocks.Add($para)
        $rtbLog.ScrollToEnd()
    })
}

function Set-UIStatus {
    param([string]$Text, [int]$Progress = -1)
    $window.Dispatcher.Invoke([action]{
        $tbStatus.Text = $Text
        if ($Progress -ge 0) { $progressBar.Value = $Progress }
    })
}

function Set-UIBusy {
    param([bool]$Busy)
    $window.Dispatcher.Invoke([action]{
        $sync.Busy = $Busy
        $btnRunFull.IsEnabled     = -not $Busy
        $btnDownload.IsEnabled    = -not $Busy
        $btnPrepare.IsEnabled     = -not $Busy
        $btnStopService.IsEnabled = -not $Busy
        $btnBackup.IsEnabled      = -not $Busy
        $btnCleanup.IsEnabled     = -not $Busy
        $btnDeploy.IsEnabled      = -not $Busy
        $btnStartService.IsEnabled= -not $Busy
        if ($Busy) { $progressBar.Value = 0 }
    })
}

function Update-ServiceStatus {
    $svc = Get-Service -Name 'crappserver' -ErrorAction SilentlyContinue
    $window.Dispatcher.Invoke([action]{
        if (-not $svc) {
            $tbServiceStatus.Text     = 'Service: Not Found'
            $ellServiceDot.Fill       = [System.Windows.Media.Brushes]::Gray
        } elseif ($svc.Status -eq 'Running') {
            $tbServiceStatus.Text     = "Service: Running"
            $ellServiceDot.Fill       = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#4ec9b0')
        } else {
            $tbServiceStatus.Text     = "Service: $($svc.Status)"
            $ellServiceDot.Fill       = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#dcdcaa')
        }
    })
}
#endregion

#region --- Runspace Worker ---
# All heavy work runs in a background runspace. Logging is funnelled back via the
# $sync.Queue; a DispatcherTimer drains the queue on the UI thread.

function Start-NixxisJob {
    param([scriptblock]$Work, [string]$JobName = 'Operation')

    if ($sync.Busy) {
        Write-UILog "Another operation is already running. Please wait." 'WARN'
        return
    }
    Set-UIBusy $true
    $sync.Abort = $false

    # Capture current UI values for the runspace
    $mode       = if ($rbOnlineAuto.IsChecked)   { 'online-auto' }
                  elseif ($rbOnlineCustom.IsChecked) { 'online-custom' }
                  else { 'offline' }
    $cpUrl      = $tbCPUrl.Text.Trim()
    $csUrl      = $tbCSUrl.Text.Trim()
    $ncsUrl     = $tbNCSUrl.Text.Trim()
    $offPath    = $tbOfflinePath.Text.Trim()

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

    $ps = [powershell]::Create()
    $ps.Runspace = $rs

    # Shared logging helper inside runspace
    $logHelper = {
        function Write-BgLog {
            param([string]$Message, [string]$Level = 'INFO')
            $sync.Queue.Enqueue(@{ Message = $Message; Level = $Level })
            # Also write direct to file
            $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            Add-Content -Path $logFile -Value "[$ts] $Message" -ErrorAction SilentlyContinue
        }
    }

    $ps.AddScript($logHelper) | Out-Null
    $ps.AddScript($Work)       | Out-Null

    $startTime = [datetime]::Now
    $handle    = $ps.BeginInvoke()

    # Timer to drain log queue and detect completion
    $timer          = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [timespan]::FromMilliseconds(150)
    $timer.Add_Tick({
        # Drain log queue
        $entry = $null
        while ($sync.Queue.TryDequeue([ref]$entry)) {
            Write-UILog -Message $entry.Message -Level $entry.Level
        }
        # Update elapsed
        $elapsed = [datetime]::Now - $startTime
        $tbElapsed.Text = "Elapsed: $($elapsed.ToString('mm\:ss'))"

        # Check completion
        if ($handle.IsCompleted) {
            $timer.Stop()
            # Drain remaining
            $entry = $null
            while ($sync.Queue.TryDequeue([ref]$entry)) {
                Write-UILog -Message $entry.Message -Level $entry.Level
            }
            # Collect errors
            if ($ps.HadErrors) {
                foreach ($err in $ps.Streams.Error) {
                    Write-UILog "Error: $err" 'ERROR'
                }
            }
            $ps.Dispose()
            $rs.Dispose()
            Set-UIBusy $false
            Update-ServiceStatus
            $progressBar.Value = 100
        }
    })
    $timer.Start()
}
#endregion

#region --- Core Operations (run inside runspace) ---

$sbDownload = {
    Write-BgLog '=== DOWNLOAD PHASE ===' 'HEADER'

    $baseUrl = 'http://update.nixxis.net'

    function Get-LatestFolder($url, $desc) {
        Write-BgLog "Fetching directory: $url" 'CYAN'
        $resp    = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30
        $folders = @()
        foreach ($pat in @('href="([^"]+/)"','(v\d+\.\d+)','(\d+\.\d+\.\d+)')) {
            [regex]::Matches($resp.Content, $pat) | ForEach-Object {
                $n = $_.Groups[1].Value.TrimEnd('/')
                if ($n -notmatch '^(\.\.|\.|icons|cgi-bin)$' -and $n -notin $folders) { $folders += $n }
            }
        }
        if (-not $folders) { throw "No folders found at $url" }
        $latest = ($folders | Sort-Object {
            $v = $_ -replace '[^\d\.]',''
            try { [version]$v } catch { [version]'0.0' }
        } | Select-Object -Last 1)
        Write-BgLog "Latest $desc`: $latest" 'OK'
        return $latest
    }

    function Get-ZipFile($url, $dest) {
        Write-BgLog "Downloading: $url" 'CYAN'
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($url, $dest)
        $sz = [math]::Round((Get-Item $dest).Length / 1MB, 2)
        Write-BgLog "  → Saved $(Split-Path $dest -Leaf) ($sz MB)" 'OK'
        $wc.Dispose()
    }

    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    if (-not $scriptDir) { $scriptDir = $env:TEMP }

    if ($mode -eq 'offline') {
        $src = if ($offlinePath) { $offlinePath } else { $scriptDir }
        Write-BgLog "Offline mode — source: $src" 'MAGENTA'
        @('ClientProvisioning.zip','ClientSoftware.zip','NCS.zip') | ForEach-Object {
            $f = Join-Path $src $_
            if (-not (Test-Path $f)) { throw "Missing offline file: $f" }
            Write-BgLog "  Found $_" 'OK'
            if ($src -ne $scriptDir) { Copy-Item $f $scriptDir -Force }
        }
    } else {
        # Resolve URLs
        $resolvedCP  = $cpUrl
        $resolvedCS  = $csUrl
        $resolvedNCS = $ncsUrl

        if (-not $resolvedCP -or -not $resolvedCS -or -not $resolvedNCS) {
            $latestVer    = Get-LatestFolder $baseUrl    'version'
            $versionUrl   = "$baseUrl/$latestVer"

            if (-not $resolvedCP -or -not $resolvedCS) {
                $clientUrl  = "$versionUrl/Client"
                $latestCli  = Get-LatestFolder $clientUrl 'client build'
                if (-not $resolvedCP)  { $resolvedCP  = "$clientUrl/$latestCli/ClientProvisioning.zip" }
                if (-not $resolvedCS)  { $resolvedCS  = "$clientUrl/$latestCli/ClientSoftware.zip" }
            }
            if (-not $resolvedNCS) {
                $serverUrl  = "$versionUrl/Server"
                $latestSrv  = Get-LatestFolder $serverUrl 'server build'
                $resolvedNCS = "$serverUrl/$latestSrv/NCS.zip"
            }
        }

        Write-BgLog "ClientProvisioning URL : $resolvedCP"  'GRAY'
        Write-BgLog "ClientSoftware URL     : $resolvedCS"  'GRAY'
        Write-BgLog "NCS URL                : $resolvedNCS" 'GRAY'

        Get-ZipFile $resolvedCP  (Join-Path $scriptDir 'ClientProvisioning.zip')
        Get-ZipFile $resolvedCS  (Join-Path $scriptDir 'ClientSoftware.zip')
        Get-ZipFile $resolvedNCS (Join-Path $scriptDir 'NCS.zip')
    }
    Write-BgLog 'Download phase complete.' 'OK'
}

$sbPrepare = {
    Write-BgLog '=== PREPARE / EXTRACT PHASE ===' 'HEADER'
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    if (-not $scriptDir) { $scriptDir = $env:TEMP }

    @('NCS.zip','ClientProvisioning.zip','ClientSoftware.zip') | ForEach-Object {
        if (-not (Test-Path (Join-Path $scriptDir $_))) { throw "Required ZIP not found: $_" }
    }

    $appServer = Join-Path $scriptDir 'NixxisApplicationServer'
    if (Test-Path $appServer) {
        Write-BgLog "Removing old NixxisApplicationServer folder..." 'WARN'
        Remove-Item $appServer -Recurse -Force
    }
    New-Item $appServer -ItemType Directory -Force | Out-Null
    Write-BgLog "Created NixxisApplicationServer" 'OK'

    Write-BgLog "Extracting NCS.zip..." 'CYAN'
    Expand-Archive (Join-Path $scriptDir 'NCS.zip') -DestinationPath $appServer -Force
    Write-BgLog "  Done" 'OK'

    $csSrc = Join-Path $appServer 'ClientSoftware'
    New-Item $csSrc -ItemType Directory -Force | Out-Null
    Write-BgLog "Extracting ClientSoftware.zip..." 'CYAN'
    Expand-Archive (Join-Path $scriptDir 'ClientSoftware.zip') -DestinationPath $csSrc -Force
    Write-BgLog "  Done" 'OK'

    $provPath   = Join-Path $appServer 'CrAppServer\provisioning'
    $provClient = Join-Path $provPath 'client'
    New-Item $provPath   -ItemType Directory -Force | Out-Null
    New-Item $provClient -ItemType Directory -Force | Out-Null
    Copy-Item (Join-Path $scriptDir 'ClientSoftware.zip') $provPath -Force
    Write-BgLog "Extracting ClientProvisioning.zip..." 'CYAN'
    Expand-Archive (Join-Path $scriptDir 'ClientProvisioning.zip') -DestinationPath $provClient -Force
    Write-BgLog "  Done" 'OK'

    $settingsSrc = Join-Path $provClient 'settings'
    if (Test-Path $settingsSrc) {
        Move-Item $settingsSrc $provPath -Force
        Write-BgLog "Moved settings folder" 'OK'
    }
    Write-BgLog 'Preparation phase complete.' 'OK'
}

$sbStopService = {
    Write-BgLog '=== STOP SERVICE PHASE ===' 'HEADER'
    $svcName  = 'crappserver'
    $procName = 'crappserver'

    # Kill desktop client
    $client = Get-Process 'nixxisclientdesktop' -ErrorAction SilentlyContinue
    if ($client) {
        Write-BgLog "Killing nixxisclientdesktop.exe (PID $($client.Id))..." 'WARN'
        & taskkill /IM nixxisclientdesktop.exe /F 2>&1 | Out-Null
        Start-Sleep 2
        if (Get-Process 'nixxisclientdesktop' -ErrorAction SilentlyContinue) {
            Write-BgLog "WARNING: nixxisclientdesktop still running." 'WARN'
        } else { Write-BgLog "nixxisclientdesktop terminated." 'OK' }
    } else { Write-BgLog "nixxisclientdesktop not running." 'GRAY' }

    # Stop service
    Stop-Service -Name $svcName -Force -ErrorAction SilentlyContinue
    $max = 60; $i = 0
    do {
        $svc  = Get-Service $svcName -ErrorAction SilentlyContinue
        $proc = Get-Process $procName -ErrorAction SilentlyContinue
        if ($svc.Status -eq 'Stopped' -and -not $proc) {
            Start-Sleep 3
            if (-not (Get-Process $procName -ErrorAction SilentlyContinue)) {
                Write-BgLog "Crappserver fully stopped." 'OK'; break
            }
        }
        $i++
        Write-BgLog "Waiting for service to stop… attempt $i/$max" 'WARN'
        Stop-Service -Name $svcName -Force -ErrorAction SilentlyContinue
        Start-Sleep 2
    } until ($i -ge $max)
    if ($i -ge $max) { throw "Service did not stop within 120 seconds." }
    Write-BgLog 'Service stop phase complete.' 'OK'
}

$sbBackup = {
    Write-BgLog '=== BACKUP PHASE ===' 'HEADER'
    $base   = 'C:\NixxisMaintenance\BackUp'
    $date   = Get-Date -Format 'yyyyMMdd'
    $year   = (Get-Date).Year.ToString()
    $dest   = Join-Path $base "$year\$date"
    $nms    = Join-Path $dest 'NMS'
    $sql    = Join-Path $dest 'SQL'
    foreach ($p in $dest,$nms,$sql) { New-Item $p -ItemType Directory -Force | Out-Null }
    @('NMS1','NMS2','NMS3','NMS4') | ForEach-Object { New-Item (Join-Path $nms $_) -ItemType Directory -Force | Out-Null }
    Write-BgLog "Backup folder: $dest" 'CYAN'

    foreach ($src in @('C:\Nixxis\CrAppServer','C:\Nixxis\ClientSoftware')) {
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
        Write-BgLog "CrAppServer not found — nothing to clean." 'WARN'
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
                catch { Write-BgLog "Could not delete $f : $_" 'WARN' }
            }
        }

        # .pdb files
        Get-ChildItem $base -Filter '*.pdb' -File -ErrorAction SilentlyContinue | ForEach-Object {
            try { Remove-Item $_.FullName -Force; Write-BgLog "Deleted PDB: $($_.Name)" 'GRAY' }
            catch { Write-BgLog "Could not delete PDB $($_.Name): $_" 'WARN' }
        }

        # Provisioning folder (except Settings)
        $prov = Join-Path $base 'Provisioning'
        if (Test-Path $prov) {
            Get-ChildItem $prov | Where-Object { $_.Name -ne 'Settings' } | ForEach-Object {
                Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                Write-BgLog "Removed from Provisioning: $($_.Name)" 'GRAY'
            }
        }
    }

    # ClientSoftware folder
    $cs = 'C:\Nixxis\ClientSoftware'
    if (Test-Path $cs) {
        Remove-Item $cs -Recurse -Force -ErrorAction SilentlyContinue
        Write-BgLog "Removed ClientSoftware folder" 'OK'
    } else { Write-BgLog "ClientSoftware not found — skipping." 'WARN' }

    Write-BgLog 'Cleanup phase complete.' 'OK'
}

$sbDeploy = {
    Write-BgLog '=== DEPLOY PHASE ===' 'HEADER'
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    if (-not $scriptDir) { $scriptDir = $env:TEMP }
    $src = Join-Path $scriptDir 'NixxisApplicationServer'
    if (-not (Test-Path $src)) { throw "NixxisApplicationServer folder not found. Run Prepare step first." }

    $folders = @(
        @{ S = 'ClientSoftware' ; D = 'C:\Nixxis\ClientSoftware'   },
        @{ S = 'CrAppServer'    ; D = 'C:\Nixxis\CrAppServer'      },
        @{ S = 'MediaServer'    ; D = 'C:\Nixxis\MediaServer'       },
        @{ S = 'Reporting'      ; D = 'C:\Nixxis\Reporting'         },
        @{ S = 'SampleConfigFiles'; D='C:\Nixxis\SampleConfigFiles' },
        @{ S = 'SoundsSamples'  ; D = 'C:\Nixxis\SoundsSamples'     }
    )

    if (-not (Test-Path 'C:\Nixxis')) { New-Item 'C:\Nixxis' -ItemType Directory -Force | Out-Null }

    foreach ($f in $folders) {
        $srcPath = Join-Path $src $f.S
        if (Test-Path $srcPath) {
            Write-BgLog "Deploying $($f.S) → $($f.D)..." 'CYAN'
            if (-not (Test-Path $f.D)) { New-Item $f.D -ItemType Directory -Force | Out-Null }
            Copy-Item "$srcPath\*" $f.D -Recurse -Force
            Write-BgLog "  Done" 'OK'
        } else { Write-BgLog "Source not found: $($f.S) — skipping." 'WARN' }
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
    if ($svc.Status -eq 'Running') {
        Write-BgLog "Crappserver is Running." 'OK'
    } else {
        Write-BgLog "Service status after start: $($svc.Status)" 'WARN'
    }
}

# Full update: chains all phases
$sbFullUpdate = [scriptblock]::Create(@"
$($sbDownload.ToString())
$($sbPrepare.ToString())
$($sbStopService.ToString())
$($sbBackup.ToString())
$($sbCleanup.ToString())
$($sbDeploy.ToString())
Write-BgLog '=============================' 'HEADER'
Write-BgLog 'FULL UPDATE COMPLETE' 'OK'
Write-BgLog 'Start the Nixxis service manually or click Start Service.' 'CYAN'
"@)
#endregion

#region --- Event Handlers ---

# Mode radio buttons
$rbOnlineCustom.Add_Checked({ $pnlCustomUrls.Visibility = 'Visible'; $pnlOfflinePath.Visibility = 'Collapsed' })
$rbOffline.Add_Checked({      $pnlOfflinePath.Visibility = 'Visible'; $pnlCustomUrls.Visibility = 'Collapsed' })
$rbOnlineAuto.Add_Checked({   $pnlCustomUrls.Visibility = 'Collapsed'; $pnlOfflinePath.Visibility = 'Collapsed' })

# Browse offline folder
$btnBrowse.Add_Click({
    $dlg         = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description = 'Select folder containing the 3 Nixxis ZIP files'
    if ($dlg.ShowDialog() -eq 'OK') { $tbOfflinePath.Text = $dlg.SelectedPath }
})

# Refresh service status
$btnRefreshStatus.Add_Click({ Update-ServiceStatus })

# Log controls
$btnClearLog.Add_Click({
    $rtbLog.Document.Blocks.Clear()
    $sync.LogLines.Clear()
})
$btnSaveLog.Add_Click({
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title      = 'Save Log File'
    $dlg.Filter     = 'Text Files (*.txt)|*.txt|All Files (*.*)|*.*'
    $dlg.FileName   = "NixxisLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    if ($dlg.ShowDialog() -eq $true) {
        $sync.LogLines | Set-Content $dlg.FileName -Encoding UTF8
        Write-UILog "Log saved to: $($dlg.FileName)" 'OK'
    }
})

# Quick actions
$btnOpenBackup.Add_Click({  $p = 'C:\NixxisMaintenance\BackUp'; if (Test-Path $p) { Start-Process explorer $p } else { Write-UILog "Backup folder not found yet." 'WARN' } })
$btnOpenLogs.Add_Click({    $p = 'C:\NixxisMaintenance\Logs';   if (Test-Path $p) { Start-Process explorer $p } else { Write-UILog "Logs folder not found yet." 'WARN' } })
$btnOpenNixxis.Add_Click({  $p = 'C:\Nixxis';                   if (Test-Path $p) { Start-Process explorer $p } else { Write-UILog "C:\Nixxis not found on this machine." 'WARN' } })

# Operation buttons
$btnRunFull.Add_Click({
    Write-UILog '==== STARTING FULL UPDATE ====' 'HEADER'
    Set-UIStatus 'Running full update…' 0
    Start-NixxisJob -Work $sbFullUpdate -JobName 'Full Update'
})
$btnDownload.Add_Click({
    Write-UILog '==== DOWNLOAD STEP ====' 'HEADER'
    Set-UIStatus 'Downloading ZIPs…' 0
    Start-NixxisJob -Work $sbDownload -JobName 'Download'
})
$btnPrepare.Add_Click({
    Write-UILog '==== PREPARE STEP ====' 'HEADER'
    Set-UIStatus 'Preparing files…' 0
    Start-NixxisJob -Work $sbPrepare -JobName 'Prepare'
})
$btnStopService.Add_Click({
    Write-UILog '==== STOP SERVICE ====' 'HEADER'
    Set-UIStatus 'Stopping Nixxis service…' 0
    Start-NixxisJob -Work $sbStopService -JobName 'Stop Service'
})
$btnBackup.Add_Click({
    Write-UILog '==== BACKUP STEP ====' 'HEADER'
    Set-UIStatus 'Backing up files…' 0
    Start-NixxisJob -Work $sbBackup -JobName 'Backup'
})
$btnCleanup.Add_Click({
    Write-UILog '==== CLEANUP STEP ====' 'HEADER'
    Set-UIStatus 'Cleaning up old files…' 0
    Start-NixxisJob -Work $sbCleanup -JobName 'Cleanup'
})
$btnDeploy.Add_Click({
    Write-UILog '==== DEPLOY STEP ====' 'HEADER'
    Set-UIStatus 'Deploying new files…' 0
    Start-NixxisJob -Work $sbDeploy -JobName 'Deploy'
})
$btnStartService.Add_Click({
    Write-UILog '==== START SERVICE ====' 'HEADER'
    Set-UIStatus 'Starting Nixxis service…' 0
    Start-NixxisJob -Work $sbStartService -JobName 'Start Service'
})
#endregion

#region --- Startup ---
Write-UILog '  Nixxis Maintenance Tool  ' 'HEADER'
Write-UILog "  Log file: $logFile"         'GRAY'
Write-UILog "  Running as: $env:USERNAME"  'GRAY'
Write-UILog "  Host: $env:COMPUTERNAME"    'GRAY'
Write-UILog '  Select a mode and click Run Full Update, or use individual steps.' 'CYAN'
Write-UILog '' 'INFO'
Update-ServiceStatus

$window.ShowDialog() | Out-Null
#endregion
