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
    Title="Nixxis NCS Installer and Updater"
    Height="760" Width="1180"
    MinHeight="620" MinWidth="960"
    WindowStartupLocation="CenterScreen"
    Background="#0f141a"
    FontFamily="Segoe UI"
    FontSize="13">

    <Window.Resources>
        <Style x:Key="PrimaryBtn" TargetType="Button">
            <Setter Property="Background" Value="#0e7490"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Margin" Value="2,2"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#0891b2"/></Trigger>
                            <Trigger Property="IsPressed"   Value="True"><Setter TargetName="bd" Property="Background" Value="#0f5f73"/></Trigger>
                            <Trigger Property="IsEnabled"   Value="False">
                                <Setter TargetName="bd" Property="Background" Value="#29323d"/>
                                <Setter Property="Foreground" Value="#607080"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="ActionBtn" TargetType="Button" BasedOn="{StaticResource PrimaryBtn}">
            <Setter Property="Background" Value="#1b2530"/>
            <Setter Property="HorizontalAlignment" Value="Stretch"/>
            <Setter Property="Height" Value="27"/>
            <Setter Property="FontSize" Value="11"/>
        </Style>
        <Style x:Key="UpdatePrimaryBtn" TargetType="Button" BasedOn="{StaticResource PrimaryBtn}">
            <Setter Property="Background" Value="#0e7490"/>
            <Setter Property="Height" Value="32"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style x:Key="InstallPrimaryBtn" TargetType="Button" BasedOn="{StaticResource PrimaryBtn}">
            <Setter Property="Background" Value="#2f855a"/>
            <Setter Property="Height" Value="32"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style x:Key="GreenBtn" TargetType="Button" BasedOn="{StaticResource ActionBtn}">
            <Setter Property="Background" Value="#2f855a"/>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Foreground" Value="#8bc3d8"/>
            <Setter Property="BorderBrush" Value="#2c3a46"/>
            <Setter Property="Margin" Value="0,0,0,6"/>
            <Setter Property="Padding" Value="7"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#17202b"/>
            <Setter Property="Foreground" Value="#dde7ef"/>
            <Setter Property="BorderBrush" Value="#334354"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="5,3"/>
            <Setter Property="Margin" Value="2,1"/>
            <Setter Property="CaretBrush" Value="White"/>
        </Style>
        <Style TargetType="RadioButton">
            <Setter Property="Foreground" Value="#d3dfeb"/>
            <Setter Property="Margin" Value="0,2"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#d3dfeb"/>
            <Setter Property="Margin" Value="0,2"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="#b8c6d3"/>
            <Setter Property="Padding" Value="2,2"/>
            <Setter Property="FontSize" Value="11"/>
        </Style>
        <Style TargetType="Separator">
            <Setter Property="Background" Value="#2c3a46"/>
            <Setter Property="Margin" Value="0,4"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="58"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="38"/>
        </Grid.RowDefinitions>

        <!-- HEADER -->
        <Border Grid.Row="0" Background="#111a22" BorderBrush="#0e7490" BorderThickness="0,0,0,2">
            <Grid Margin="14,0">
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="&#9881;" Foreground="#0ea5c5" FontSize="24" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock Text="Nixxis NCS Installer / Updater" Foreground="White" FontSize="16" FontWeight="SemiBold"/>
                        <TextBlock Text="Explicit flows for Fresh Install and Existing Update" Foreground="#7e93a8" FontSize="10"/>
                    </StackPanel>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,4,0">
                    <Border Background="#0f141a" CornerRadius="5" Padding="9,4" Margin="0,0,6,0">
                        <StackPanel Orientation="Horizontal">
                            <Ellipse x:Name="ellServiceDot" Width="9" Height="9" Fill="#888" Margin="0,0,7,0" VerticalAlignment="Center"/>
                            <TextBlock x:Name="tbServiceStatus" Text="Service: Unknown" Foreground="#b8c6d3" FontSize="11" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Border>
                    <Button x:Name="btnRefreshStatus" Content="Refresh" Style="{StaticResource ActionBtn}" Width="74"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- MAIN -->
        <Grid Grid.Row="1" Margin="8,8,8,4">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="320"/>
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

                    <!-- Nixxis Install Path -->
                    <GroupBox Header="  Nixxis Install Path  ">
                        <StackPanel>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="tbInstallPath" Grid.Column="0" Text="C:\Nixxis" FontSize="11"/>
                                <Button x:Name="btnBrowseInstall" Grid.Column="1" Content="..." Width="34" Height="28"
                                        Style="{StaticResource ActionBtn}" Margin="4,2,0,2" FontSize="13"/>
                            </Grid>
                            <TextBlock Foreground="#666" FontSize="10" TextWrapping="Wrap" Margin="2,2,0,0"
                                       Text="This path is used for deploy, service install, tools, MoveFiles, firewall and reporting actions."/>
                        </StackPanel>
                    </GroupBox>

                    <!-- Operation Mode -->
                    <GroupBox Header="  Operation Mode  ">
                        <StackPanel>
                            <RadioButton x:Name="rbModeUpdate" Content="Update Existing NCS" IsChecked="True"/>
                            <TextBlock Foreground="#879aab" FontSize="10" Margin="18,0,0,2" Text="Use when NCS already exists and you want to refresh binaries."/>
                            <RadioButton x:Name="rbModeInstall" Content="Fresh Install NCS"/>
                            <TextBlock Foreground="#879aab" FontSize="10" Margin="18,0,0,2" Text="Use on new servers or complete rebuilds with initial setup tasks."/>
                            <Border Background="#152532" CornerRadius="5" Padding="8,6" Margin="0,4,0,0">
                                <TextBlock x:Name="tbModeHint" Foreground="#9bd3e6" FontSize="10" TextWrapping="Wrap"
                                           Text="Mode: Update Existing NCS. Update workflow buttons are visible below."/>
                            </Border>
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

                    <StackPanel x:Name="pnlUpdateFlow">
                        <!-- Update Flow -->
                        <GroupBox Header="  Update Existing NCS  ">
                            <StackPanel>
                                <Button x:Name="btnRunFull"
                                        Content="Run Update Flow"
                                        Style="{StaticResource UpdatePrimaryBtn}"
                                        HorizontalAlignment="Stretch"/>
                                <TextBlock Foreground="#7f93a7" FontSize="10" TextWrapping="Wrap" Margin="4,4,4,0"
                                           Text="Download, extract, stop service, backup, cleanup and deploy binaries."/>
                            </StackPanel>
                        </GroupBox>

                        <GroupBox Header="  Update Steps (Advanced)  ">
                            <StackPanel>
                                <Button x:Name="btnDownload"     Content="1. Download ZIPs"              Style="{StaticResource ActionBtn}"/>
                                <Button x:Name="btnPrepare"      Content="2. Extract and Prepare"        Style="{StaticResource ActionBtn}"/>
                                <Separator/>
                                <Button x:Name="btnStopService"  Content="3. Stop Nixxis Service"       Style="{StaticResource ActionBtn}"/>
                                <Button x:Name="btnBackup"       Content="4. Backup Current NCS"        Style="{StaticResource ActionBtn}"/>
                                <Button x:Name="btnCleanup"      Content="5. Cleanup Old Files"         Style="{StaticResource ActionBtn}"/>
                                <Button x:Name="btnDeploy"       Content="6. Deploy New Files"          Style="{StaticResource ActionBtn}"/>
                                <Separator/>
                                <Button x:Name="btnStartService" Content="Start Nixxis Service"         Style="{StaticResource GreenBtn}"/>
                            </StackPanel>
                        </GroupBox>
                    </StackPanel>

                    <StackPanel x:Name="pnlInstallFlow" Visibility="Collapsed">
                        <!-- Fresh Install Options -->
                        <GroupBox Header="  Fresh Install Options  ">
                            <StackPanel>
                                <CheckBox x:Name="cbEnsureDotNet48"       Content="Ensure .NET Framework 4.8" IsChecked="True"/>
                                <CheckBox x:Name="cbInstallMoveFiles"     Content="Install MoveFiles service" IsChecked="True"/>
                                <CheckBox x:Name="cbCreateReportingUser"  Content="Create local Reporting user" IsChecked="False"/>
                                <CheckBox x:Name="cbConfigureFirewall"    Content="Create firewall rules (TCP/UDP)" IsChecked="True"/>
                                <CheckBox x:Name="cbDeployTranscription"  Content="Deploy transcription helper files" IsChecked="False"/>
                            </StackPanel>
                        </GroupBox>

                        <!-- Fresh Install Steps -->
                        <GroupBox Header="  Fresh Install NCS  ">
                            <StackPanel>
                                <Button x:Name="btnRunInitialSetup"       Content="Run Fresh Install Flow" Style="{StaticResource InstallPrimaryBtn}"/>
                                <TextBlock Foreground="#7f93a7" FontSize="10" TextWrapping="Wrap" Margin="4,4,4,2"
                                           Text="Includes install tasks: service install, tools, firewall, MoveFiles and optional helpers."/>
                                <Separator/>
                                <Button x:Name="btnEnsureDotNet48"        Content="Check / Install .NET 4.8"         Style="{StaticResource ActionBtn}"/>
                                <Button x:Name="btnInstallService"        Content="Install CrAppServer Service"      Style="{StaticResource ActionBtn}"/>
                                <Button x:Name="btnCopyTools"             Content="Copy Tools Folder"                Style="{StaticResource ActionBtn}"/>
                                <Button x:Name="btnInstallMoveFiles"      Content="Install MoveFiles Service"        Style="{StaticResource ActionBtn}"/>
                                <Button x:Name="btnCreateReportingUser"   Content="Create Reporting Local User"      Style="{StaticResource ActionBtn}"/>
                                <Button x:Name="btnConfigFirewall"        Content="Configure Firewall Rules"         Style="{StaticResource ActionBtn}"/>
                                <Button x:Name="btnDeployTranscription"   Content="Deploy Transcription Helpers"     Style="{StaticResource ActionBtn}"/>
                                <Button x:Name="btnLaunchDeployReports"   Content="Launch DeployReports.exe"        Style="{StaticResource GreenBtn}"/>
                            </StackPanel>
                        </GroupBox>
                    </StackPanel>

                    <!-- Quick Actions -->
                    <GroupBox Header="  Quick Actions  ">
                        <UniformGrid Columns="1" Rows="4">
                            <Button x:Name="btnOpenWork"    Content="Open Working Directory" Style="{StaticResource ActionBtn}"/>
                            <Button x:Name="btnOpenLogs"    Content="Open Logs Folder"       Style="{StaticResource ActionBtn}"/>
                            <Button x:Name="btnOpenNixxis"  Content="Open Install Folder"    Style="{StaticResource ActionBtn}"/>
                            <Button x:Name="btnOpenReportBin" Content="Open Reporting Bin Hint" Style="{StaticResource ActionBtn}"/>
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

                <Border Grid.Row="0" Background="#111a22" CornerRadius="5,5,0,0" Padding="10,6" Margin="0,0,0,1">
                    <Grid>
                        <TextBlock Text="ACTIVITY LOG" Foreground="#8bc3d8" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                            <Button x:Name="btnClearLog" Content="Clear"    Style="{StaticResource ActionBtn}" Height="24" Width="55" FontSize="11" Margin="2,0"/>
                            <Button x:Name="btnSaveLog"  Content="Save Log" Style="{StaticResource ActionBtn}" Height="24" Width="68" FontSize="11" Margin="2,0"/>
                        </StackPanel>
                    </Grid>
                </Border>

                <RichTextBox x:Name="rtbLog" Grid.Row="1"
                             Background="#0b1117" Foreground="#c8d3dd"
                             BorderBrush="#2c3a46" BorderThickness="1"
                             IsReadOnly="True"
                             FontFamily="Consolas,Courier New" FontSize="12"
                             VerticalScrollBarVisibility="Auto"
                             HorizontalScrollBarVisibility="Auto"
                             Padding="8">
                    <RichTextBox.Document>
                        <FlowDocument PageWidth="9999"/>
                    </RichTextBox.Document>
                </RichTextBox>

                <StackPanel Grid.Row="2" Background="#111a22" Margin="0,1,0,0">
                    <ProgressBar x:Name="progressBar" Height="5" Value="0" Maximum="100"
                                 Background="#111a22" Foreground="#0e7490" BorderThickness="0"/>
                    <Grid Margin="8,4">
                        <TextBlock x:Name="tbStatus"  Text="Ready" Foreground="#666" FontSize="11"/>
                        <TextBlock x:Name="tbElapsed" Text=""      Foreground="#555" FontSize="11" HorizontalAlignment="Right"/>
                    </Grid>
                </StackPanel>
            </Grid>
        </Grid>

        <!-- FOOTER -->
        <Border Grid.Row="2" Background="#111a22" BorderBrush="#2c3a46" BorderThickness="0,1,0,0">
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

$rbModeUpdate   = ctrl 'rbModeUpdate'
$rbModeInstall  = ctrl 'rbModeInstall'
$tbModeHint     = ctrl 'tbModeHint'
$pnlUpdateFlow  = ctrl 'pnlUpdateFlow'
$pnlInstallFlow = ctrl 'pnlInstallFlow'
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
$tbInstallPath   = ctrl 'tbInstallPath'
$btnBrowseWork   = ctrl 'btnBrowseWork'
$btnBrowseInstall= ctrl 'btnBrowseInstall'
$btnBrowseOffline= ctrl 'btnBrowseOffline'
$btnRunFull      = ctrl 'btnRunFull'
$btnRunInitialSetup = ctrl 'btnRunInitialSetup'
$btnDownload     = ctrl 'btnDownload'
$btnPrepare      = ctrl 'btnPrepare'
$btnStopService  = ctrl 'btnStopService'
$btnBackup       = ctrl 'btnBackup'
$btnCleanup      = ctrl 'btnCleanup'
$btnDeploy       = ctrl 'btnDeploy'
$btnStartService = ctrl 'btnStartService'
$btnInstallService = ctrl 'btnInstallService'
$btnEnsureDotNet48 = ctrl 'btnEnsureDotNet48'
$btnCopyTools      = ctrl 'btnCopyTools'
$btnInstallMoveFiles = ctrl 'btnInstallMoveFiles'
$btnCreateReportingUser = ctrl 'btnCreateReportingUser'
$btnConfigFirewall = ctrl 'btnConfigFirewall'
$btnDeployTranscription = ctrl 'btnDeployTranscription'
$btnLaunchDeployReports = ctrl 'btnLaunchDeployReports'
$cbEnsureDotNet48 = ctrl 'cbEnsureDotNet48'
$cbInstallMoveFiles = ctrl 'cbInstallMoveFiles'
$cbCreateReportingUser = ctrl 'cbCreateReportingUser'
$cbConfigureFirewall = ctrl 'cbConfigureFirewall'
$cbDeployTranscription = ctrl 'cbDeployTranscription'
$btnRefreshStatus= ctrl 'btnRefreshStatus'
$btnOpenWork     = ctrl 'btnOpenWork'
$btnOpenLogs     = ctrl 'btnOpenLogs'
$btnOpenNixxis   = ctrl 'btnOpenNixxis'
$btnOpenReportBin= ctrl 'btnOpenReportBin'
$btnClearLog     = ctrl 'btnClearLog'
$btnSaveLog      = ctrl 'btnSaveLog'
$rtbLog          = ctrl 'rtbLog'
$progressBar     = ctrl 'progressBar'
$tbStatus        = ctrl 'tbStatus'
$tbElapsed       = ctrl 'tbElapsed'
$tbServiceStatus = ctrl 'tbServiceStatus'
$ellServiceDot   = ctrl 'ellServiceDot'
$tbLogFile       = ctrl 'tbLogFile'

$allOpButtons = @(
    $btnRunFull,$btnRunInitialSetup,$btnDownload,$btnPrepare,$btnStopService,$btnBackup,$btnCleanup,$btnDeploy,$btnStartService,
    $btnEnsureDotNet48,$btnInstallService,$btnCopyTools,$btnInstallMoveFiles,$btnCreateReportingUser,$btnConfigFirewall,$btnDeployTranscription,$btnLaunchDeployReports
)
#endregion

#region --- Shared State ---
$sync = [hashtable]::Synchronized(@{
    Queue   = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
    Busy    = $false
    WorkDir = 'C:\NixxisMaintenance\Update'
    InstallPath = 'C:\Nixxis'
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

function Update-OperationModeUI {
    if ($rbModeInstall.IsChecked) {
        $pnlInstallFlow.Visibility = 'Visible'
        $pnlUpdateFlow.Visibility = 'Collapsed'
        $tbModeHint.Text = 'Mode: Fresh Install NCS. Install workflow controls are visible.'
        Add-LogEntry 'Switched to Fresh Install mode.' 'CYAN'
    } else {
        $pnlInstallFlow.Visibility = 'Collapsed'
        $pnlUpdateFlow.Visibility = 'Visible'
        $tbModeHint.Text = 'Mode: Update Existing NCS. Update workflow controls are visible.'
        Add-LogEntry 'Switched to Update Existing mode.' 'CYAN'
    }
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
    $installPath = $tbInstallPath.Text.Trim().TrimEnd('\\')
    $optEnsureDotNet48 = [bool]$cbEnsureDotNet48.IsChecked
    $optInstallMoveFiles = [bool]$cbInstallMoveFiles.IsChecked
    $optCreateReportingUser = [bool]$cbCreateReportingUser.IsChecked
    $optConfigureFirewall = [bool]$cbConfigureFirewall.IsChecked
    $optDeployTranscription = [bool]$cbDeployTranscription.IsChecked
    $runDir    = $sync.RunDir
    $sync.InstallPath = $installPath

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
    $rs.SessionStateProxy.SetVariable('installPath', $installPath)
    $rs.SessionStateProxy.SetVariable('optEnsureDotNet48', $optEnsureDotNet48)
    $rs.SessionStateProxy.SetVariable('optInstallMoveFiles', $optInstallMoveFiles)
    $rs.SessionStateProxy.SetVariable('optCreateReportingUser', $optCreateReportingUser)
    $rs.SessionStateProxy.SetVariable('optConfigureFirewall', $optConfigureFirewall)
    $rs.SessionStateProxy.SetVariable('optDeployTranscription', $optDeployTranscription)
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

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    function Expand-ZipRobust {
        param(
            [Parameter(Mandatory = $true)][string]$ZipPath,
            [Parameter(Mandatory = $true)][string]$DestinationPath
        )

        try {
            if (-not (Test-Path $DestinationPath)) {
                New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
            }

            $stream = [System.IO.File]::OpenRead($ZipPath)
            try {
                $archive = New-Object System.IO.Compression.ZipArchive(
                    $stream,
                    [System.IO.Compression.ZipArchiveMode]::Read,
                    $false,
                    [System.Text.Encoding]::UTF8
                )
                try {
                    foreach ($entry in $archive.Entries) {
                        $entryDestPath = Join-Path -Path $DestinationPath -ChildPath $entry.FullName
                        if ($entry.FullName.EndsWith('/') -or $entry.FullName.EndsWith('\\')) {
                            New-Item -Path $entryDestPath -ItemType Directory -Force | Out-Null
                            continue
                        }

                        $parentDir = Split-Path -Parent $entryDestPath
                        if (-not (Test-Path $parentDir)) {
                            New-Item -Path $parentDir -ItemType Directory -Force | Out-Null
                        }

                        $entryStream = $entry.Open()
                        try {
                            $fileStream = [System.IO.File]::Create($entryDestPath)
                            try {
                                $entryStream.CopyTo($fileStream)
                            }
                            finally {
                                $fileStream.Dispose()
                            }
                        }
                        finally {
                            $entryStream.Dispose()
                        }
                    }
                }
                finally {
                    $archive.Dispose()
                }
            }
            finally {
                $stream.Dispose()
            }
        }
        catch {
            Expand-Archive -Path $ZipPath -DestinationPath $DestinationPath -Force
        }
    }

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
    Expand-ZipRobust -ZipPath (Join-Path $runDir 'NCS.zip') -DestinationPath $appServer
    Write-BgLog "  Done" 'OK'

    $csDest = Join-Path $appServer 'ClientSoftware'
    New-Item $csDest -ItemType Directory -Force | Out-Null
    Write-BgLog "Extracting ClientSoftware.zip..." 'CYAN'
    Expand-ZipRobust -ZipPath (Join-Path $runDir 'ClientSoftware.zip') -DestinationPath $csDest
    Write-BgLog "  Done" 'OK'

    $provPath   = Join-Path $appServer 'CrAppServer\provisioning'
    $provClient = Join-Path $provPath 'client'
    New-Item $provClient -ItemType Directory -Force | Out-Null
    Copy-Item (Join-Path $runDir 'ClientSoftware.zip') $provPath -Force

    Write-BgLog "Extracting ClientProvisioning.zip..." 'CYAN'
    Expand-ZipRobust -ZipPath (Join-Path $runDir 'ClientProvisioning.zip') -DestinationPath $provClient
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

    foreach ($src in @((Join-Path $installPath 'CrAppServer'), (Join-Path $installPath 'ClientSoftware'))) {
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
    $base = Join-Path $installPath 'CrAppServer'
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
    $cs = Join-Path $installPath 'ClientSoftware'
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

    if (-not (Test-Path $installPath)) {
        New-Item -Path $installPath -ItemType Directory -Force | Out-Null
        Write-BgLog "Created install root: $installPath" 'OK'
    }

    foreach ($f in @(
        @{ S='ClientSoftware';   D=(Join-Path $installPath 'ClientSoftware') },
        @{ S='CrAppServer';      D=(Join-Path $installPath 'CrAppServer')    },
        @{ S='MediaServer';      D=(Join-Path $installPath 'MediaServer')    },
        @{ S='Reporting';        D=(Join-Path $installPath 'Reporting')      },
        @{ S='SoundsSamples';    D=(Join-Path $installPath 'SoundsSamples')  }
    )) {
        $src = Join-Path $appServer $f.S
        if (Test-Path $src) {
            Write-BgLog "Deploying $($f.S)  ->  $($f.D)..." 'CYAN'
            if (-not (Test-Path $f.D)) { New-Item $f.D -ItemType Directory -Force | Out-Null }
            Copy-Item "$src\*" $f.D -Recurse -Force
            Write-BgLog "  Done" 'OK'
        } else { Write-BgLog "Source not found in staging: $($f.S) — skipping." 'WARN' }
    }

    $sampleSource = Join-Path $appServer 'SampleConfigFiles'
    $sampleDestination = Join-Path $installPath 'CrAppServer'
    if (Test-Path $sampleSource) {
        Write-BgLog "Processing SampleConfigFiles -> $sampleDestination" 'CYAN'
        $copied = 0
        $skipped = 0
        Get-ChildItem -Path $sampleSource -File -Recurse | ForEach-Object {
            if ($_.Name -like 'NCC*') {
                $skipped++
                return
            }

            $relativePath = $_.FullName.Substring($sampleSource.Length).TrimStart('\\','/')
            $destRelative = if ($relativePath -match '\\.sample$') {
                $relativePath -replace '\\.sample$', ''
            } else {
                $relativePath
            }
            $destFile = Join-Path $sampleDestination $destRelative
            $destDir = Split-Path -Parent $destFile
            if (-not (Test-Path $destDir)) {
                New-Item -Path $destDir -ItemType Directory -Force | Out-Null
            }
            Copy-Item -Path $_.FullName -Destination $destFile -Force
            $copied++
        }
        Write-BgLog "SampleConfigFiles complete: $copied copied, $skipped skipped" 'OK'
    } else {
        Write-BgLog 'SampleConfigFiles not found in staging - skipping.' 'WARN'
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

$sbEnsureDotNet48 = {
    Write-BgLog '=== .NET FRAMEWORK 4.8 PHASE ===' 'HEADER'

    function Get-DotNetRelease {
        $regPath = 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full'
        if (Test-Path $regPath) {
            return [int]((Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).Release)
        }
        return 0
    }

    $dotNet48Required = 528040
    $dotNetRelease = Get-DotNetRelease
    if ($dotNetRelease -ge $dotNet48Required) {
        Write-BgLog ".NET Framework 4.8 already installed (release key: $dotNetRelease)" 'OK'
        return
    }

    Write-BgLog '.NET Framework 4.8 not found. Downloading installer...' 'WARN'
    $dotNet48Url = 'https://go.microsoft.com/fwlink/?LinkId=2085155'
    $dotNet48Installer = Join-Path -Path $env:TEMP -ChildPath 'ndp48-x86-x64-allos-enu.exe'

    $webClient48 = New-Object System.Net.WebClient
    try {
        $webClient48.DownloadFile($dotNet48Url, $dotNet48Installer)
    }
    finally {
        $webClient48.Dispose()
    }

    Write-BgLog '.NET 4.8 installer downloaded. Installing (silent)...' 'CYAN'
    $process = Start-Process -FilePath $dotNet48Installer -ArgumentList "/quiet /norestart /log `"$env:TEMP\dotnet48_install.log`"" -Wait -PassThru
    if ($process.ExitCode -ne 0 -and $process.ExitCode -ne 3010) {
        throw "Installer exited with code $($process.ExitCode). See $env:TEMP\dotnet48_install.log"
    }

    Remove-Item -Path $dotNet48Installer -Force -ErrorAction SilentlyContinue
    $dotNetRelease = Get-DotNetRelease
    if ($dotNetRelease -ge $dotNet48Required) {
        Write-BgLog '.NET Framework 4.8 installed successfully.' 'OK'
        Write-BgLog 'A reboot is recommended before continuing service operations.' 'WARN'
    } else {
        throw '.NET 4.8 installation could not be verified.'
    }
}

$sbInstallService = {
    Write-BgLog '=== SERVICE INSTALLATION PHASE ===' 'HEADER'
    $nixxisLogsPath = Join-Path $installPath 'Logs'
    if (-not (Test-Path $nixxisLogsPath)) {
        New-Item -Path $nixxisLogsPath -ItemType Directory -Force | Out-Null
        Write-BgLog "Created Logs folder: $nixxisLogsPath" 'OK'
    }

    $crAppServerDir = Join-Path $installPath 'CrAppServer'
    $crAppServerExe = Join-Path $crAppServerDir 'CrAppServer.exe'
    if (-not (Test-Path $crAppServerExe)) {
        throw "CrAppServer.exe not found at $crAppServerExe"
    }

    Push-Location $crAppServerDir
    try {
        $svcOutput = & $crAppServerExe -install 2>&1
        foreach ($line in $svcOutput) {
            Write-BgLog "  [CrAppServer] $line" 'GRAY'
        }
    }
    finally {
        Pop-Location
    }
    Write-BgLog 'Service installation phase complete.' 'OK'
}

$sbCopyTools = {
    Write-BgLog '=== TOOLS COPY PHASE ===' 'HEADER'
    $toolsSource = Join-Path (Join-Path $runDir 'NixxisApplicationServer') 'Tools'
    $toolsDestination = Join-Path $installPath 'Tools'

    if (-not (Test-Path $toolsSource)) {
        throw "Tools source folder not found: $toolsSource"
    }
    if (-not (Test-Path $toolsDestination)) {
        New-Item -Path $toolsDestination -ItemType Directory -Force | Out-Null
    }
    Copy-Item -Path "$toolsSource\*" -Destination $toolsDestination -Recurse -Force
    Write-BgLog "Tools folder copied: $toolsDestination" 'OK'
}

$sbInstallMoveFiles = {
    Write-BgLog '=== MOVEFILES INSTALL PHASE ===' 'HEADER'
    $installDrive = Split-Path -Qualifier $installPath
    $moveFilesExe = Join-Path $installPath 'Tools\MoveFiles\MoveFiles.exe'
    $installUtilExe = "$installDrive\Windows\Microsoft.NET\Framework64\v4.0.30319\installutil.exe"

    if (-not (Test-Path $moveFilesExe)) {
        throw "MoveFiles.exe not found at $moveFilesExe"
    }
    if (-not (Test-Path $installUtilExe)) {
        throw "installutil.exe not found at $installUtilExe"
    }

    $mfOutput = & $installUtilExe $moveFilesExe 2>&1
    foreach ($line in $mfOutput) {
        Write-BgLog "  [installutil] $line" 'GRAY'
    }

    $sampleXml = Join-Path $installPath 'Tools\MoveFiles\SampleMoveFiles.xml'
    $targetXml = Join-Path $installPath 'Tools\MoveFiles\MoveFiles.xml'
    if (Test-Path $sampleXml) {
        if (Test-Path $targetXml) {
            Remove-Item -Path $sampleXml -Force
            Write-BgLog 'MoveFiles.xml already exists - removed SampleMoveFiles.xml' 'GRAY'
        } else {
            Rename-Item -Path $sampleXml -NewName 'MoveFiles.xml' -Force
            Write-BgLog 'Renamed SampleMoveFiles.xml -> MoveFiles.xml' 'OK'
        }
    }

    Write-BgLog "Action required: review MoveFiles configuration at $targetXml" 'WARN'
    Write-BgLog 'MoveFiles installation phase complete.' 'OK'
}

$sbCreateReportingUser = {
    Write-BgLog '=== REPORTING USER PHASE ===' 'HEADER'
    $reportingUsername = 'Reporting'
    $reportingPassword = ConvertTo-SecureString 'Rep0rting' -AsPlainText -Force

    $existingUser = Get-LocalUser -Name $reportingUsername -ErrorAction SilentlyContinue
    if ($existingUser) {
        Write-BgLog "Local user '$reportingUsername' already exists - skipping." 'WARN'
    } else {
        New-LocalUser -Name $reportingUsername -Password $reportingPassword -PasswordNeverExpires -UserMayNotChangePassword -Description 'Nixxis Reporting Services account' | Out-Null
        Write-BgLog "Local user '$reportingUsername' created." 'OK'
    }
}

$sbConfigFirewall = {
    Write-BgLog '=== FIREWALL PHASE ===' 'HEADER'
    $crAppServerExeFw = Join-Path $installPath 'CrAppServer\CrAppServer.exe'
    if (-not (Test-Path $crAppServerExeFw)) {
        throw "CrAppServer.exe not found for firewall rule at $crAppServerExeFw"
    }

    foreach ($proto in @('TCP', 'UDP')) {
        $ruleName = "Nixxis CrAppServer $proto"
        Remove-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue
        New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Program $crAppServerExeFw -Protocol $proto -Action Allow -Profile Any | Out-Null
        Write-BgLog "Firewall rule created: $ruleName" 'OK'
    }
}

$sbDeployTranscription = {
    Write-BgLog '=== TRANSCRIPTION HELPERS PHASE ===' 'HEADER'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $transcriptDestination = Join-Path $installPath 'CrAppServer\provisioning\client'
    $transcriptTempPath = Join-Path $runDir 'TranscriptClientHelpers'
    $baseUrl = 'https://raw.githubusercontent.com/NixxisIntegration/TranscriptionClientHelpers/refs/heads/main'
    $transcriptFiles = @(
        'DarkTranscriptions.css',
        'LightTranscriptions.css',
        'defaultTranscriptions.css',
        'handleTranscriptions.js',
        'handleTranscriptions_en.js',
        'handleTranscriptions_fr.js'
    )

    if (Test-Path $transcriptTempPath) {
        Remove-Item -Path $transcriptTempPath -Recurse -Force
    }
    New-Item -Path $transcriptTempPath -ItemType Directory -Force | Out-Null
    if (-not (Test-Path $transcriptDestination)) {
        New-Item -Path $transcriptDestination -ItemType Directory -Force | Out-Null
    }

    foreach ($fileName in $transcriptFiles) {
        $fileUrl = "$baseUrl/$fileName"
        $tempFilePath = Join-Path $transcriptTempPath $fileName
        Invoke-WebRequest -Uri $fileUrl -OutFile $tempFilePath -UseBasicParsing -TimeoutSec 60
        Copy-Item -Path $tempFilePath -Destination (Join-Path $transcriptDestination $fileName) -Force
        Write-BgLog "Deployed: $fileName" 'OK'
    }

    Remove-Item -Path $transcriptTempPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-BgLog 'Transcription helpers phase complete.' 'OK'
}

$sbLaunchDeployReports = {
    Write-BgLog '=== DEPLOY REPORTS LAUNCH ===' 'HEADER'
    $deployReportsExe = Join-Path $installPath 'Reporting\Deploy\DeployReports.exe'
    if (Test-Path $deployReportsExe) {
        Start-Process -FilePath $deployReportsExe
        Write-BgLog "Launched: $deployReportsExe" 'OK'
    } else {
        throw "DeployReports.exe not found at $deployReportsExe"
    }

    Write-BgLog 'Manual action: copy contents of Reporting\ToCopyReportingServer into SQL Reporting Services ReportServer\bin.' 'WARN'
}

$sbInitialSetupFull = [scriptblock]::Create(
    "if ($optEnsureDotNet48) {" + "`n" +
    $sbEnsureDotNet48.ToString() + "`n" +
    "} else { Write-BgLog 'Skipped .NET 4.8 check by option.' 'GRAY' }" + "`n" +
    $sbDownload.ToString() + "`n" +
    $sbPrepare.ToString() + "`n" +
    $sbStopService.ToString() + "`n" +
    $sbBackup.ToString() + "`n" +
    $sbCleanup.ToString() + "`n" +
    $sbDeploy.ToString() + "`n" +
    $sbInstallService.ToString() + "`n" +
    $sbCopyTools.ToString() + "`n" +
    "if ($optInstallMoveFiles) {" + "`n" +
    $sbInstallMoveFiles.ToString() + "`n" +
    "} else { Write-BgLog 'Skipped MoveFiles installation by option.' 'GRAY' }" + "`n" +
    "if ($optCreateReportingUser) {" + "`n" +
    $sbCreateReportingUser.ToString() + "`n" +
    "} else { Write-BgLog 'Skipped Reporting user creation by option.' 'GRAY' }" + "`n" +
    "if ($optConfigureFirewall) {" + "`n" +
    $sbConfigFirewall.ToString() + "`n" +
    "} else { Write-BgLog 'Skipped firewall configuration by option.' 'GRAY' }" + "`n" +
    "if ($optDeployTranscription) {" + "`n" +
    $sbDeployTranscription.ToString() + "`n" +
    "} else { Write-BgLog 'Skipped transcription helpers by option.' 'GRAY' }" + "`n" +
    "Write-BgLog '=== INITIAL SETUP COMPLETE ===' 'OK'"
)

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
$rbModeInstall.Add_Checked({ Update-OperationModeUI })
$rbModeUpdate.Add_Checked({ Update-OperationModeUI })
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

$btnBrowseInstall.Add_Click({
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description = 'Select Nixxis install directory'
    $dlg.SelectedPath = $tbInstallPath.Text
    if ($dlg.ShowDialog() -eq 'OK') {
        $tbInstallPath.Text = $dlg.SelectedPath.TrimEnd('\\')
        $sync.InstallPath = $tbInstallPath.Text
    }
})

# WorkDir text change — keep sync live
$tbWorkDir.Add_TextChanged({
    $sync.WorkDir = $tbWorkDir.Text.Trim()
    Update-RunDirLabel
})

$tbInstallPath.Add_TextChanged({
    $sync.InstallPath = $tbInstallPath.Text.Trim().TrimEnd('\\')
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
    $p = $tbInstallPath.Text.Trim().TrimEnd('\\')
    if (Test-Path $p) { Start-Process explorer $p }
    else { Add-LogEntry "Install path not found: $p" 'WARN' }
})
$btnOpenReportBin.Add_Click({
    $reportPath = Join-Path $tbInstallPath.Text.Trim().TrimEnd('\\') 'Reporting\ToCopyReportingServer'
    if (Test-Path $reportPath) { Start-Process explorer $reportPath }
    else { Add-LogEntry "Reporting helper folder not found: $reportPath" 'WARN' }
})

# Operation buttons
$btnRunInitialSetup.Add_Click({    Add-LogEntry '==== INITIAL SETUP ====' 'HEADER';   Set-Status 'Running initial setup...' 0; Start-NixxisJob $sbInitialSetupFull 'Initial Setup' })
$btnEnsureDotNet48.Add_Click({     Add-LogEntry '==== ENSURE .NET 4.8 ====' 'HEADER'; Set-Status 'Checking .NET...' 0; Start-NixxisJob $sbEnsureDotNet48 'Ensure .NET 4.8' })
$btnRunFull.Add_Click({     Add-LogEntry '==== FULL UPDATE ====' 'HEADER';   Set-Status 'Running full update...' 0;  Start-NixxisJob $sbFullUpdate    'Full Update'   })
$btnDownload.Add_Click({    Add-LogEntry '==== DOWNLOAD ====' 'HEADER';      Set-Status 'Downloading ZIPs...' 0;    Start-NixxisJob $sbDownload      'Download'      })
$btnPrepare.Add_Click({     Add-LogEntry '==== PREPARE ====' 'HEADER';       Set-Status 'Preparing files...' 0;     Start-NixxisJob $sbPrepare       'Prepare'       })
$btnStopService.Add_Click({ Add-LogEntry '==== STOP SERVICE ====' 'HEADER';  Set-Status 'Stopping service...' 0;    Start-NixxisJob $sbStopService   'Stop Service'  })
$btnBackup.Add_Click({      Add-LogEntry '==== BACKUP ====' 'HEADER';        Set-Status 'Backing up...' 0;         Start-NixxisJob $sbBackup        'Backup'        })
$btnCleanup.Add_Click({     Add-LogEntry '==== CLEANUP ====' 'HEADER';       Set-Status 'Cleaning up...' 0;        Start-NixxisJob $sbCleanup       'Cleanup'       })
$btnDeploy.Add_Click({      Add-LogEntry '==== DEPLOY ====' 'HEADER';        Set-Status 'Deploying...' 0;          Start-NixxisJob $sbDeploy        'Deploy'        })
$btnStartService.Add_Click({Add-LogEntry '==== START SERVICE ====' 'HEADER'; Set-Status 'Starting service...' 0;   Start-NixxisJob $sbStartService  'Start Service' })
$btnInstallService.Add_Click({     Add-LogEntry '==== INSTALL SERVICE ====' 'HEADER'; Set-Status 'Installing service...' 0; Start-NixxisJob $sbInstallService 'Install Service' })
$btnCopyTools.Add_Click({          Add-LogEntry '==== COPY TOOLS ====' 'HEADER';      Set-Status 'Copying tools...' 0;      Start-NixxisJob $sbCopyTools 'Copy Tools' })
$btnInstallMoveFiles.Add_Click({   Add-LogEntry '==== INSTALL MOVEFILES ====' 'HEADER'; Set-Status 'Installing MoveFiles...' 0; Start-NixxisJob $sbInstallMoveFiles 'Install MoveFiles' })
$btnCreateReportingUser.Add_Click({Add-LogEntry '==== CREATE REPORTING USER ====' 'HEADER'; Set-Status 'Creating Reporting user...' 0; Start-NixxisJob $sbCreateReportingUser 'Create Reporting User' })
$btnConfigFirewall.Add_Click({     Add-LogEntry '==== CONFIG FIREWALL ====' 'HEADER'; Set-Status 'Configuring firewall...' 0; Start-NixxisJob $sbConfigFirewall 'Configure Firewall' })
$btnDeployTranscription.Add_Click({Add-LogEntry '==== DEPLOY TRANSCRIPTION ====' 'HEADER'; Set-Status 'Deploying transcription helpers...' 0; Start-NixxisJob $sbDeployTranscription 'Deploy Transcription Helpers' })
$btnLaunchDeployReports.Add_Click({Add-LogEntry '==== LAUNCH DEPLOY REPORTS ====' 'HEADER'; Set-Status 'Launching DeployReports...' 0; Start-NixxisJob $sbLaunchDeployReports 'Launch DeployReports' })
#endregion

#region --- Startup ---
Update-RunDirLabel
Update-OperationModeUI
Add-LogEntry 'Nixxis Maintenance Tool ready.' 'HEADER'
Add-LogEntry "User: $env:USERNAME  |  Host: $env:COMPUTERNAME" 'GRAY'
Add-LogEntry "Log file: $logFile" 'GRAY'
Add-LogEntry "Default working directory: $($sync.WorkDir)" 'CYAN'
Add-LogEntry "Install path: $($sync.InstallPath)" 'CYAN'
Add-LogEntry "A dated subfolder (YYYYMMDD) will be created in that directory on each run." 'GRAY'
Add-LogEntry '' 'INFO'
Update-ServiceStatus

$window.ShowDialog() | Out-Null
#endregion
