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

# Force TLS 1.2 - required for GitHub/web downloads on PowerShell 5.1 / Windows Server
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

$script:AppVersion = '1.4'

#region --- XAML ---
[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="NCS Installer / Updater"
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
        <Style TargetType="TabControl">
            <Setter Property="Background" Value="#131b24"/>
            <Setter Property="Foreground" Value="#e6edf3"/>
            <Setter Property="BorderBrush" Value="#3b4d5f"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Foreground" Value="#d8e3ee"/>
            <Setter Property="Background" Value="#1a2430"/>
            <Setter Property="Margin" Value="0,0,2,0"/>
            <Setter Property="Padding" Value="12,4"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="tabBd" Background="{TemplateBinding Background}" BorderBrush="#3b4d5f" BorderThickness="1,1,1,0" CornerRadius="4,4,0,0" Padding="{TemplateBinding Padding}">
                            <ContentPresenter ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="tabBd" Property="Background" Value="#243445"/>
                                <Setter Property="Foreground" Value="#ffffff"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="tabBd" Property="Background" Value="#223140"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ListView">
            <Setter Property="Background" Value="#151f2a"/>
            <Setter Property="Foreground" Value="#e6edf3"/>
            <Setter Property="BorderBrush" Value="#3b4d5f"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="ListViewItem">
            <Setter Property="Foreground" Value="#e6edf3"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Padding" Value="3,2"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListViewItem">
                        <Border x:Name="lvItemBd" Background="{TemplateBinding Background}" BorderBrush="Transparent" BorderThickness="0" Padding="{TemplateBinding Padding}">
                            <GridViewRowPresenter VerticalAlignment="Center" Content="{TemplateBinding Content}" Columns="{TemplateBinding GridView.ColumnCollection}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="lvItemBd" Property="Background" Value="#284057"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="lvItemBd" Property="Background" Value="#1d2d3c"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="PlanRowStyle" TargetType="ListViewItem" BasedOn="{StaticResource {x:Type ListViewItem}}">
            <Style.Triggers>
                <DataTrigger Binding="{Binding Badge}" Value="CREATE">
                    <Setter Property="Background" Value="#173228"/>
                    <Setter Property="Foreground" Value="#d9ffee"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Badge}" Value="OVERWRITE">
                    <Setter Property="Background" Value="#3a3118"/>
                    <Setter Property="Foreground" Value="#fff1c4"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Badge}" Value="DELETE">
                    <Setter Property="Background" Value="#3d1f29"/>
                    <Setter Property="Foreground" Value="#ffdbe0"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Badge}" Value="SKIP">
                    <Setter Property="Background" Value="#25313f"/>
                    <Setter Property="Foreground" Value="#dbe7f3"/>
                </DataTrigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="RiskRowStyle" TargetType="ListViewItem" BasedOn="{StaticResource {x:Type ListViewItem}}">
            <Style.Triggers>
                <DataTrigger Binding="{Binding Status}" Value="GREEN">
                    <Setter Property="Background" Value="#173228"/>
                    <Setter Property="Foreground" Value="#d9ffee"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Status}" Value="YELLOW">
                    <Setter Property="Background" Value="#3a3118"/>
                    <Setter Property="Foreground" Value="#fff1c4"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Status}" Value="RED">
                    <Setter Property="Background" Value="#3d1f29"/>
                    <Setter Property="Foreground" Value="#ffdbe0"/>
                </DataTrigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="ValidationRowStyle" TargetType="ListViewItem" BasedOn="{StaticResource {x:Type ListViewItem}}">
            <Style.Triggers>
                <DataTrigger Binding="{Binding Status}" Value="PASS">
                    <Setter Property="Background" Value="#173228"/>
                    <Setter Property="Foreground" Value="#d9ffee"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Status}" Value="WARN">
                    <Setter Property="Background" Value="#3a3118"/>
                    <Setter Property="Foreground" Value="#fff1c4"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Status}" Value="FAIL">
                    <Setter Property="Background" Value="#3d1f29"/>
                    <Setter Property="Foreground" Value="#ffdbe0"/>
                </DataTrigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="TimelineRowStyle" TargetType="ListViewItem" BasedOn="{StaticResource {x:Type ListViewItem}}">
            <Style.Triggers>
                <DataTrigger Binding="{Binding State}" Value="In Progress">
                    <Setter Property="Background" Value="#183349"/>
                    <Setter Property="Foreground" Value="#d8ecff"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding State}" Value="Done">
                    <Setter Property="Background" Value="#173228"/>
                    <Setter Property="Foreground" Value="#d9ffee"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding State}" Value="Pending">
                    <Setter Property="Background" Value="#25313f"/>
                    <Setter Property="Foreground" Value="#dbe7f3"/>
                </DataTrigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="GridViewColumnHeader">
            <Setter Property="Background" Value="#1c2a38"/>
            <Setter Property="Foreground" Value="#f2f7fb"/>
            <Setter Property="BorderBrush" Value="#3b4d5f"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="Padding" Value="6,4"/>
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
                        <TextBlock x:Name="tbHeaderTitle" Text="NCS Installer / Updater" Foreground="White" FontSize="16" FontWeight="SemiBold"/>
                        <TextBlock x:Name="tbHeaderSubtitle" Text="Explicit flows for Fresh Install and Existing Update | v1.1" Foreground="#7e93a8" FontSize="10"/>
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
                            <Label Content="Windows service name:"/>
                            <TextBox x:Name="tbServiceName" Text="crappserver" FontSize="11"/>
                            <TextBlock Foreground="#666" FontSize="10" TextWrapping="Wrap" Margin="2,2,0,0"
                                       Text="Install path and service name are used for status, start/stop, service install and firewall actions."/>
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
                            <RadioButton x:Name="rbOnlineAuto"   Content="Online - Auto-Discover latest"  IsChecked="True"/>
                            <RadioButton x:Name="rbOnlineSelect" Content="Online - Select discovered versions"/>
                            <RadioButton x:Name="rbOnlineCustom" Content="Online - Custom URLs"/>
                            <RadioButton x:Name="rbOffline"      Content="Offline - Use local ZIP files"/>

                            <StackPanel x:Name="pnlOnlineSelect" Visibility="Collapsed" Margin="8,6,0,0">
                                <Button x:Name="btnDiscoverVersions" Content="Discover Available Versions" Style="{StaticResource ActionBtn}" Width="190" HorizontalAlignment="Left"/>
                                <TextBlock x:Name="tbDiscoveredBaseVersion" Foreground="#7f93a7" FontSize="10" Margin="2,4,0,2" Text="Base version: (not discovered)"/>
                                <Label Content="Client version:"/>
                                <ComboBox x:Name="cbClientVersion" MinHeight="26" Margin="2,1"/>
                                <Label Content="Server version:"/>
                                <ComboBox x:Name="cbServerVersion" MinHeight="26" Margin="2,1"/>
                                <TextBlock Foreground="#666" FontSize="10" TextWrapping="Wrap" Margin="2,2,0,0"
                                           Text="Use Discover first, then select the versions you want to deploy."/>
                            </StackPanel>

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

                        <GroupBox Header="  HTTP.Config Wizard (Install Only)  ">
                            <StackPanel>
                                <TextBlock Foreground="#7f93a7" FontSize="10" TextWrapping="Wrap" Margin="2,0,0,6"
                                           Text="Configure critical http.config values before deploy. This applies only during Fresh Install."/>

                                <Label Content="Domain host list (semicolon separated):"/>
                                <TextBox x:Name="tbHttpHostList" FontSize="11" Text="localhost:8088;localhost"/>
                                <TextBlock x:Name="tbHttpErrHost" Foreground="#f44747" FontSize="10" Margin="2,1,0,4" Text=""/>

                                <CheckBox x:Name="cbHttpSqlIntegrated" Content="Use SQL Integrated Security" IsChecked="True" Margin="0,2,0,0"/>
                                <Label Content="SQL Server / Data Source:"/>
                                <TextBox x:Name="tbHttpSqlServer" FontSize="11" Text="localhost"/>
                                <TextBlock x:Name="tbHttpErrSqlServer" Foreground="#f44747" FontSize="10" Margin="2,1,0,2" Text=""/>
                                <Label Content="SQL Initial Catalog pattern:"/>
                                <TextBox x:Name="tbHttpSqlCatalog" FontSize="11" Text="{0}_{1}"/>
                                <TextBlock x:Name="tbHttpErrSqlCatalog" Foreground="#f44747" FontSize="10" Margin="2,1,0,4" Text=""/>

                                <Label Content="SQL User (required when integrated security is off):"/>
                                <TextBox x:Name="tbHttpSqlUser" FontSize="11" Text=""/>
                                <TextBlock x:Name="tbHttpErrSqlUser" Foreground="#f44747" FontSize="10" Margin="2,1,0,2" Text=""/>
                                <Label Content="SQL Password:"/>
                                <PasswordBox x:Name="pbHttpSqlPassword" FontSize="11"/>
                                <TextBlock x:Name="tbHttpErrSqlPassword" Foreground="#f44747" FontSize="10" Margin="2,1,0,4" Text=""/>

                                <Label Content="Admin credential user (NixxisAdminRole):"/>
                                <TextBox x:Name="tbHttpAdminUser" FontSize="11" Text="DefaultAdmin"/>
                                <TextBlock x:Name="tbHttpErrAdminUser" Foreground="#f44747" FontSize="10" Margin="2,1,0,2" Text=""/>
                                <Label Content="Admin credential password:"/>
                                <PasswordBox x:Name="pbHttpAdminPassword" FontSize="11"/>
                                <TextBlock x:Name="tbHttpErrAdminPassword" Foreground="#f44747" FontSize="10" Margin="2,1,0,4" Text=""/>

                                <Label Content="Reporting credential (domain\user):"/>
                                <TextBox x:Name="tbHttpReportingUser" FontSize="11" Text=""/>
                                <TextBlock x:Name="tbHttpErrReportingUser" Foreground="#f44747" FontSize="10" Margin="2,1,0,2" Text=""/>
                                <Label Content="Reporting credential password:"/>
                                <PasswordBox x:Name="pbHttpReportingPassword" FontSize="11"/>
                                <TextBlock x:Name="tbHttpErrReportingPassword" Foreground="#f44747" FontSize="10" Margin="2,1,0,4" Text=""/>

                                <Label Content="Report Server URL:"/>
                                <TextBox x:Name="tbHttpReportServerUrl" FontSize="11" Text="http://localhost/ReportServer"/>
                                <TextBlock x:Name="tbHttpErrReportServer" Foreground="#f44747" FontSize="10" Margin="2,1,0,4" Text=""/>

                                <CheckBox x:Name="cbHttpAddAgentReactivation" Content="Add key in ACD: agentReactivationEnabled" IsChecked="True"/>
                                <CheckBox x:Name="cbHttpAgentReactivationValue" Content="agentReactivationEnabled value = true (unchecked = false)" IsChecked="False" Margin="14,0,0,4"/>

                                <CheckBox x:Name="cbHttpAddClientVirtualizationFilter" Content="Add key in Admin: ClientVirtualizationFilter" IsChecked="True"/>
                                <Label Content="ClientVirtualizationFilter value:"/>
                                <TextBox x:Name="tbHttpClientVirtualizationFilter" FontSize="11" Text=".+"/>
                                <TextBlock x:Name="tbHttpErrClientVirtualizationFilter" Foreground="#f44747" FontSize="10" Margin="2,1,0,4" Text=""/>

                                <Label Content="Recording FTP hosts (semicolon separated):"/>
                                <TextBox x:Name="tbHttpRecordingHosts" FontSize="11" Text=""/>
                                <TextBlock x:Name="tbHttpErrRecordingHosts" Foreground="#f44747" FontSize="10" Margin="2,1,0,2" Text=""/>
                                <Label Content="Recording FTP user:"/>
                                <TextBox x:Name="tbHttpRecordingUser" FontSize="11" Text="recording"/>
                                <TextBlock x:Name="tbHttpErrRecordingUser" Foreground="#f44747" FontSize="10" Margin="2,1,0,2" Text=""/>
                                <Label Content="Recording FTP password:"/>
                                <PasswordBox x:Name="pbHttpRecordingPassword" FontSize="11"/>
                                <TextBlock x:Name="tbHttpErrRecordingPassword" Foreground="#f44747" FontSize="10" Margin="2,1,0,2" Text=""/>
                                <Label Content="Recording folder (example: 509):"/>
                                <TextBox x:Name="tbHttpRecordingFolder" FontSize="11" Text=""/>
                                <TextBlock x:Name="tbHttpErrRecordingFolder" Foreground="#f44747" FontSize="10" Margin="2,1,0,4" Text=""/>

                                <CheckBox x:Name="cbHttpAddSupervision" Content="Add application supervision" IsChecked="False" Margin="0,0,0,6"/>

                                <StackPanel Orientation="Horizontal">
                                    <Button x:Name="btnHttpValidate" Content="Validate Config Inputs" Style="{StaticResource ActionBtn}" Width="145" Margin="0,0,6,0"/>
                                    <Button x:Name="btnHttpPreview" Content="Preview XML Changes" Style="{StaticResource ActionBtn}" Width="130" Margin="0,0,6,0"/>
                                    <Button x:Name="btnHttpSaveProfile" Content="Save Install Profile" Style="{StaticResource ActionBtn}" Width="130"/>
                                </StackPanel>
                                <TextBlock x:Name="tbHttpConfigSummary" Foreground="#b6c7d8" FontSize="10" Margin="2,5,0,0" Text="No validation run yet."/>
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

            <!-- RIGHT PANEL - Operations Console -->
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <Border Grid.Row="0" Background="#111a22" CornerRadius="5,5,0,0" Padding="10,6" Margin="0,0,0,1">
                    <Grid>
                        <TextBlock Text="OPERATIONS CONSOLE" Foreground="#8bc3d8" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                            <Button x:Name="btnClearLog" Content="Clear"    Style="{StaticResource ActionBtn}" Height="24" Width="55" FontSize="11" Margin="2,0"/>
                            <Button x:Name="btnSaveLog"  Content="Save Log" Style="{StaticResource ActionBtn}" Height="24" Width="68" FontSize="11" Margin="2,0"/>
                            <Button x:Name="btnExportSignedReport" Content="Export Signed Report" Style="{StaticResource ActionBtn}" Height="24" Width="140" FontSize="11" Margin="2,0"/>
                        </StackPanel>
                    </Grid>
                </Border>

                <TabControl x:Name="tcOps" Grid.Row="1" Background="#131b24" BorderBrush="#3b4d5f" BorderThickness="1">
                    <TabItem Header="Log">
                        <RichTextBox x:Name="rtbLog"
                                     Background="#151f2a" Foreground="#f3f7fb"
                                     BorderThickness="0"
                                     IsReadOnly="True"
                                     FontFamily="Consolas,Courier New" FontSize="12"
                                     VerticalScrollBarVisibility="Auto"
                                     HorizontalScrollBarVisibility="Auto"
                                     Padding="8">
                            <RichTextBox.Document>
                                <FlowDocument PageWidth="9999"/>
                            </RichTextBox.Document>
                        </RichTextBox>
                    </TabItem>

                    <TabItem Header="Plan">
                        <Grid Margin="8">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <StackPanel Grid.Row="0" Orientation="Horizontal">
                                <Button x:Name="btnGeneratePlan" Content="Generate Plan" Style="{StaticResource ActionBtn}" Width="120"/>
                                <TextBlock x:Name="tbPlanSummary" Foreground="#e6edf3" Margin="10,6,0,0" Text="Plan not generated."/>
                            </StackPanel>
                            <CheckBox x:Name="chkPlanApproved" Grid.Row="1" Margin="2,6,0,4" Content="I reviewed and approve this plan"/>
                            <TextBlock Grid.Row="2" Foreground="#b6c7d8" FontSize="10" Text="Badges: Create / Overwrite / Delete / Skip"/>
                            <ListView x:Name="lvPlan" Grid.Row="3" Margin="0,6,0,0" Background="#151f2a" BorderBrush="#3b4d5f" BorderThickness="1" ItemContainerStyle="{StaticResource PlanRowStyle}">
                                <ListView.View>
                                    <GridView>
                                        <GridViewColumn Header="Badge" Width="85" DisplayMemberBinding="{Binding Badge}"/>
                                        <GridViewColumn Header="Action" Width="210" DisplayMemberBinding="{Binding Action}"/>
                                        <GridViewColumn Header="Path" Width="280" DisplayMemberBinding="{Binding Path}"/>
                                        <GridViewColumn Header="Details" Width="260" DisplayMemberBinding="{Binding Details}"/>
                                    </GridView>
                                </ListView.View>
                            </ListView>
                        </Grid>
                    </TabItem>

                    <TabItem Header="Preflight">
                        <Grid Margin="8">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <StackPanel Grid.Row="0" Orientation="Horizontal">
                                <Button x:Name="btnRunPreflight" Content="Run Preflight" Style="{StaticResource ActionBtn}" Width="120"/>
                                <TextBlock x:Name="tbPreflightSummary" Foreground="#e6edf3" Margin="10,6,0,0" Text="Preflight not run."/>
                            </StackPanel>
                            <ListView x:Name="lvPreflight" Grid.Row="1" Margin="0,8,0,0" Background="#151f2a" BorderBrush="#3b4d5f" BorderThickness="1" ItemContainerStyle="{StaticResource RiskRowStyle}">
                                <ListView.View>
                                    <GridView>
                                        <GridViewColumn Header="Status" Width="85" DisplayMemberBinding="{Binding Status}"/>
                                        <GridViewColumn Header="Risk" Width="220" DisplayMemberBinding="{Binding Risk}"/>
                                        <GridViewColumn Header="Details" Width="520" DisplayMemberBinding="{Binding Details}"/>
                                    </GridView>
                                </ListView.View>
                            </ListView>
                        </Grid>
                    </TabItem>

                    <TabItem Header="Timeline">
                        <Grid Margin="8">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <Grid Grid.Row="0" Margin="0,0,0,6">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock x:Name="tbTimelineCurrent" Grid.Column="0" Foreground="#e6edf3" Text="Current: -"/>
                                <TextBlock x:Name="tbTimelineNext" Grid.Column="1" Foreground="#e6edf3" Text="Next: -"/>
                                <TextBlock x:Name="tbTimelineEta" Grid.Column="2" Foreground="#e6edf3" Text="ETA: -" HorizontalAlignment="Right"/>
                            </Grid>
                            <ListView x:Name="lvTimeline" Grid.Row="1" Background="#151f2a" BorderBrush="#3b4d5f" BorderThickness="1" ItemContainerStyle="{StaticResource TimelineRowStyle}">
                                <ListView.View>
                                    <GridView>
                                        <GridViewColumn Header="Step" Width="250" DisplayMemberBinding="{Binding Step}"/>
                                        <GridViewColumn Header="State" Width="120" DisplayMemberBinding="{Binding State}"/>
                                        <GridViewColumn Header="Started" Width="120" DisplayMemberBinding="{Binding Started}"/>
                                        <GridViewColumn Header="Finished" Width="120" DisplayMemberBinding="{Binding Finished}"/>
                                        <GridViewColumn Header="Notes" Width="240" DisplayMemberBinding="{Binding Notes}"/>
                                    </GridView>
                                </ListView.View>
                            </ListView>
                        </Grid>
                    </TabItem>

                    <TabItem Header="Validation">
                        <Grid Margin="8">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <StackPanel Grid.Row="0" Orientation="Horizontal">
                                <Button x:Name="btnRunValidation" Content="Run Post-Install Validation" Style="{StaticResource ActionBtn}" Width="190"/>
                                <TextBlock x:Name="tbValidationSummary" Foreground="#e6edf3" Margin="10,6,0,0" Text="Validation not run."/>
                            </StackPanel>
                            <ListView x:Name="lvValidation" Grid.Row="1" Margin="0,8,0,0" Background="#151f2a" BorderBrush="#3b4d5f" BorderThickness="1" ItemContainerStyle="{StaticResource ValidationRowStyle}">
                                <ListView.View>
                                    <GridView>
                                        <GridViewColumn Header="Status" Width="85" DisplayMemberBinding="{Binding Status}"/>
                                        <GridViewColumn Header="Check" Width="220" DisplayMemberBinding="{Binding Check}"/>
                                        <GridViewColumn Header="Details" Width="520" DisplayMemberBinding="{Binding Details}"/>
                                    </GridView>
                                </ListView.View>
                            </ListView>
                        </Grid>
                    </TabItem>
                </TabControl>

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
$tbHeaderTitle  = ctrl 'tbHeaderTitle'
$tbHeaderSubtitle = ctrl 'tbHeaderSubtitle'
$rbOnlineAuto    = ctrl 'rbOnlineAuto'
$rbOnlineSelect  = ctrl 'rbOnlineSelect'
$rbOnlineCustom  = ctrl 'rbOnlineCustom'
$rbOffline       = ctrl 'rbOffline'
$pnlOnlineSelect = ctrl 'pnlOnlineSelect'
$pnlCustomUrls   = ctrl 'pnlCustomUrls'
$pnlOfflinePath  = ctrl 'pnlOfflinePath'
$btnDiscoverVersions = ctrl 'btnDiscoverVersions'
$cbClientVersion = ctrl 'cbClientVersion'
$cbServerVersion = ctrl 'cbServerVersion'
$tbDiscoveredBaseVersion = ctrl 'tbDiscoveredBaseVersion'
$tbCPUrl         = ctrl 'tbCPUrl'
$tbCSUrl         = ctrl 'tbCSUrl'
$tbNCSUrl        = ctrl 'tbNCSUrl'
$tbOfflinePath   = ctrl 'tbOfflinePath'
$tbWorkDir       = ctrl 'tbWorkDir'
$tbRunDir        = ctrl 'tbRunDir'
$tbInstallPath   = ctrl 'tbInstallPath'
$tbServiceName   = ctrl 'tbServiceName'
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
$btnExportSignedReport = ctrl 'btnExportSignedReport'
$tcOps           = ctrl 'tcOps'
$btnGeneratePlan = ctrl 'btnGeneratePlan'
$tbPlanSummary   = ctrl 'tbPlanSummary'
$chkPlanApproved = ctrl 'chkPlanApproved'
$lvPlan          = ctrl 'lvPlan'
$btnRunPreflight = ctrl 'btnRunPreflight'
$tbPreflightSummary = ctrl 'tbPreflightSummary'
$lvPreflight     = ctrl 'lvPreflight'
$tbTimelineCurrent = ctrl 'tbTimelineCurrent'
$tbTimelineNext  = ctrl 'tbTimelineNext'
$tbTimelineEta   = ctrl 'tbTimelineEta'
$lvTimeline      = ctrl 'lvTimeline'
$btnRunValidation = ctrl 'btnRunValidation'
$tbValidationSummary = ctrl 'tbValidationSummary'
$lvValidation    = ctrl 'lvValidation'
$rtbLog          = ctrl 'rtbLog'
$progressBar     = ctrl 'progressBar'
$tbStatus        = ctrl 'tbStatus'
$tbElapsed       = ctrl 'tbElapsed'
$tbServiceStatus = ctrl 'tbServiceStatus'
$ellServiceDot   = ctrl 'ellServiceDot'
$tbLogFile       = ctrl 'tbLogFile'
$tbHttpHostList = ctrl 'tbHttpHostList'
$cbHttpSqlIntegrated = ctrl 'cbHttpSqlIntegrated'
$tbHttpSqlServer = ctrl 'tbHttpSqlServer'
$tbHttpSqlCatalog = ctrl 'tbHttpSqlCatalog'
$tbHttpSqlUser = ctrl 'tbHttpSqlUser'
$pbHttpSqlPassword = ctrl 'pbHttpSqlPassword'
$tbHttpAdminUser = ctrl 'tbHttpAdminUser'
$pbHttpAdminPassword = ctrl 'pbHttpAdminPassword'
$tbHttpReportingUser = ctrl 'tbHttpReportingUser'
$pbHttpReportingPassword = ctrl 'pbHttpReportingPassword'
$tbHttpReportServerUrl = ctrl 'tbHttpReportServerUrl'
$cbHttpAddAgentReactivation = ctrl 'cbHttpAddAgentReactivation'
$cbHttpAgentReactivationValue = ctrl 'cbHttpAgentReactivationValue'
$cbHttpAddClientVirtualizationFilter = ctrl 'cbHttpAddClientVirtualizationFilter'
$tbHttpClientVirtualizationFilter = ctrl 'tbHttpClientVirtualizationFilter'
$tbHttpRecordingHosts = ctrl 'tbHttpRecordingHosts'
$tbHttpRecordingUser = ctrl 'tbHttpRecordingUser'
$pbHttpRecordingPassword = ctrl 'pbHttpRecordingPassword'
$tbHttpRecordingFolder = ctrl 'tbHttpRecordingFolder'
$cbHttpAddSupervision = ctrl 'cbHttpAddSupervision'
$btnHttpValidate = ctrl 'btnHttpValidate'
$btnHttpPreview = ctrl 'btnHttpPreview'
$btnHttpSaveProfile = ctrl 'btnHttpSaveProfile'
$tbHttpConfigSummary = ctrl 'tbHttpConfigSummary'
$tbHttpErrHost = ctrl 'tbHttpErrHost'
$tbHttpErrSqlServer = ctrl 'tbHttpErrSqlServer'
$tbHttpErrSqlCatalog = ctrl 'tbHttpErrSqlCatalog'
$tbHttpErrSqlUser = ctrl 'tbHttpErrSqlUser'
$tbHttpErrSqlPassword = ctrl 'tbHttpErrSqlPassword'
$tbHttpErrAdminUser = ctrl 'tbHttpErrAdminUser'
$tbHttpErrAdminPassword = ctrl 'tbHttpErrAdminPassword'
$tbHttpErrReportingUser = ctrl 'tbHttpErrReportingUser'
$tbHttpErrReportingPassword = ctrl 'tbHttpErrReportingPassword'
$tbHttpErrReportServer = ctrl 'tbHttpErrReportServer'
$tbHttpErrClientVirtualizationFilter = ctrl 'tbHttpErrClientVirtualizationFilter'
$tbHttpErrRecordingHosts = ctrl 'tbHttpErrRecordingHosts'
$tbHttpErrRecordingUser = ctrl 'tbHttpErrRecordingUser'
$tbHttpErrRecordingPassword = ctrl 'tbHttpErrRecordingPassword'
$tbHttpErrRecordingFolder = ctrl 'tbHttpErrRecordingFolder'

$allOpButtons = @(
    $btnRunFull,$btnRunInitialSetup,$btnDownload,$btnPrepare,$btnStopService,$btnBackup,$btnCleanup,$btnDeploy,$btnStartService,
    $btnEnsureDotNet48,$btnInstallService,$btnCopyTools,$btnInstallMoveFiles,$btnCreateReportingUser,$btnConfigFirewall,$btnDeployTranscription,$btnLaunchDeployReports,
    $btnHttpValidate,$btnHttpPreview,$btnHttpSaveProfile
)

$script:PlanItems = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$script:PreflightItems = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$script:TimelineItems = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$script:ValidationItems = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$lvPlan.ItemsSource = $script:PlanItems
$lvPreflight.ItemsSource = $script:PreflightItems
$lvTimeline.ItemsSource = $script:TimelineItems
$lvValidation.ItemsSource = $script:ValidationItems

$script:TimelineStepOrder = @()
$script:TimelineCurrentIndex = -1
$script:TimelineStartedAt = $null
$script:HttpConfigValidationErrors = @{}
$script:HttpConfigErrorTargets = @{
    host = $tbHttpErrHost
    sqlServer = $tbHttpErrSqlServer
    sqlCatalog = $tbHttpErrSqlCatalog
    sqlUser = $tbHttpErrSqlUser
    sqlPassword = $tbHttpErrSqlPassword
    adminUser = $tbHttpErrAdminUser
    adminPassword = $tbHttpErrAdminPassword
    reportingUser = $tbHttpErrReportingUser
    reportingPassword = $tbHttpErrReportingPassword
    reportServerUrl = $tbHttpErrReportServer
    clientVirtualizationFilter = $tbHttpErrClientVirtualizationFilter
    recordingHosts = $tbHttpErrRecordingHosts
    recordingUser = $tbHttpErrRecordingUser
    recordingPassword = $tbHttpErrRecordingPassword
    recordingFolder = $tbHttpErrRecordingFolder
}
#endregion

#region --- Shared State ---
$sync = [hashtable]::Synchronized(@{
    Queue   = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
    Busy    = $false
    WorkDir = 'C:\NixxisMaintenance\Update'
    InstallPath = 'C:\Nixxis'
    ServiceName = 'crappserver'
    SelectedBaseVersion = ''
    SelectedClientVersion = ''
    SelectedServerVersion = ''
    RunDir  = ''
    LogLines= [System.Collections.Generic.List[string]]::new()
})

# Log file - in Logs folder alongside WorkDir's parent
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

# Add one log line directly to RTB - safe on UI thread, no Dispatcher needed
function Add-LogEntry([string]$Message, [string]$Level = 'INFO') {
    $ts   = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts] $Message"
    $hex  = if ($colorMap.ContainsKey($Level)) { $colorMap[$Level] } else { '#c8c8c8' }

    # Persist to file
    Add-Content -Path $logFile -Value $line -ErrorAction SilentlyContinue
    $sync.LogLines.Add($line)

    # RTB - already on UI thread
    $para          = [System.Windows.Documents.Paragraph]::new()
    $para.Margin   = [System.Windows.Thickness]::new(0)
    $run           = [System.Windows.Documents.Run]::new($line)
    $run.Foreground= Get-Brush $hex
    $para.Inlines.Add($run)
    $rtbLog.Document.Blocks.Add($para)
    $rtbLog.ScrollToEnd()

    Update-TimelineFromLog $Message
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
    $serviceName = $tbServiceName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($serviceName)) {
        $serviceName = 'crappserver'
        $tbServiceName.Text = $serviceName
    }

    $svc = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
    if (-not $svc) {
        $tbServiceStatus.Text = "Service '$serviceName': Not Found"
        $ellServiceDot.Fill   = Get-Brush '#888888'
    } elseif ($svc.Status -eq 'Running') {
        $tbServiceStatus.Text = "Service '$serviceName': Running"
        $ellServiceDot.Fill   = Get-Brush '#4ec9b0'
    } else {
        $tbServiceStatus.Text = "Service '$serviceName': $($svc.Status)"
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
        $tbHttpConfigSummary.Text = 'HTTP.Config wizard active for Fresh Install.'
        Add-LogEntry 'Switched to Fresh Install mode.' 'CYAN'
    } else {
        $pnlInstallFlow.Visibility = 'Collapsed'
        $pnlUpdateFlow.Visibility = 'Visible'
        $tbModeHint.Text = 'Mode: Update Existing NCS. Update workflow controls are visible.'
        $tbHttpConfigSummary.Text = 'HTTP.Config wizard is ignored in Update mode.'
        Add-LogEntry 'Switched to Update Existing mode.' 'CYAN'
    }
    Update-HttpConfigUiState
}

function Get-DiscoveredWebFolders {
    param([string]$Url)

    $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 30
    $folders = @()
    foreach ($pat in @('href="([^"]+/)"', '(v\d+\.\d+)', '(\d+\.\d+\.\d+)')) {
        [regex]::Matches($resp.Content, $pat) | ForEach-Object {
            $n = $_.Groups[1].Value.TrimEnd('/')
            if ($n -notmatch '^(\.\.|\.|icons|cgi-bin)$' -and $n -notin $folders) {
                $folders += $n
            }
        }
    }

    return $folders
}

function Get-LatestVersionValue {
    param([string[]]$Values)

    if (-not $Values -or $Values.Count -eq 0) { return '' }
    return ($Values | Sort-Object {
        $v = $_ -replace '[^\d\.]', ''
        try { [version]$v } catch { [version]'0.0' }
    } | Select-Object -Last 1)
}

function Update-SourceModeUI {
    if ($rbOffline.IsChecked) {
        $pnlOfflinePath.Visibility = 'Visible'
        $pnlCustomUrls.Visibility = 'Collapsed'
        $pnlOnlineSelect.Visibility = 'Collapsed'
    } elseif ($rbOnlineCustom.IsChecked) {
        $pnlCustomUrls.Visibility = 'Visible'
        $pnlOfflinePath.Visibility = 'Collapsed'
        $pnlOnlineSelect.Visibility = 'Collapsed'
    } elseif ($rbOnlineSelect.IsChecked) {
        $pnlOnlineSelect.Visibility = 'Visible'
        $pnlCustomUrls.Visibility = 'Collapsed'
        $pnlOfflinePath.Visibility = 'Collapsed'
    } else {
        $pnlOnlineSelect.Visibility = 'Collapsed'
        $pnlCustomUrls.Visibility = 'Collapsed'
        $pnlOfflinePath.Visibility = 'Collapsed'
    }
}

function Get-HttpConfigUiSettings {
    return [ordered]@{
        hostList = $tbHttpHostList.Text.Trim()
        sqlIntegrated = [bool]$cbHttpSqlIntegrated.IsChecked
        sqlServer = $tbHttpSqlServer.Text.Trim()
        sqlCatalog = $tbHttpSqlCatalog.Text.Trim()
        sqlUser = $tbHttpSqlUser.Text.Trim()
        sqlPassword = $pbHttpSqlPassword.Password
        adminUser = $tbHttpAdminUser.Text.Trim()
        adminPassword = $pbHttpAdminPassword.Password
        reportingUser = $tbHttpReportingUser.Text.Trim()
        reportingPassword = $pbHttpReportingPassword.Password
        reportServerUrl = $tbHttpReportServerUrl.Text.Trim()
        addAgentReactivation = [bool]$cbHttpAddAgentReactivation.IsChecked
        agentReactivationValue = [bool]$cbHttpAgentReactivationValue.IsChecked
        addClientVirtualizationFilter = [bool]$cbHttpAddClientVirtualizationFilter.IsChecked
        clientVirtualizationFilter = $tbHttpClientVirtualizationFilter.Text.Trim()
        recordingHosts = $tbHttpRecordingHosts.Text.Trim()
        recordingUser = $tbHttpRecordingUser.Text.Trim()
        recordingPassword = $pbHttpRecordingPassword.Password
        recordingFolder = $tbHttpRecordingFolder.Text.Trim().Trim('/','\\')
        addSupervision = [bool]$cbHttpAddSupervision.IsChecked
    }
}

function Clear-HttpConfigFieldErrors {
    foreach ($target in $script:HttpConfigErrorTargets.Values) {
        $target.Text = ''
    }
    $script:HttpConfigValidationErrors = @{}
}

function Set-HttpConfigFieldError {
    param([string]$Field, [string]$Message)

    if (-not $script:HttpConfigValidationErrors.ContainsKey($Field)) {
        $script:HttpConfigValidationErrors[$Field] = @()
    }
    $script:HttpConfigValidationErrors[$Field] += $Message
}

function Show-HttpConfigFieldErrors {
    foreach ($key in $script:HttpConfigErrorTargets.Keys) {
        $messages = @()
        if ($script:HttpConfigValidationErrors.ContainsKey($key)) {
            $messages = $script:HttpConfigValidationErrors[$key]
        }
        $script:HttpConfigErrorTargets[$key].Text = ($messages -join ' ')
    }
}

function Build-RecordingUrlList {
    param($Settings)

    $hosts = @()
    if (-not [string]::IsNullOrWhiteSpace($Settings.recordingHosts)) {
        $hosts = @($Settings.recordingHosts -split ';' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    if ($hosts.Count -eq 0) { return '' }

    $folder = $Settings.recordingFolder.Trim('/','\\')
    $urls = @()
    foreach ($h in $hosts) {
        $urls += ('ftp://{0}:{1}@{2}/{3}/' -f $Settings.recordingUser, $Settings.recordingPassword, $h, $folder)
    }
    return ($urls -join ';')
}

function Validate-HttpConfigSettings {
    param([bool]$ShowInlineErrors = $true)

    $settings = Get-HttpConfigUiSettings
    Clear-HttpConfigFieldErrors

    if ([string]::IsNullOrWhiteSpace($settings.hostList)) {
        Set-HttpConfigFieldError 'host' 'Domain host list is required.'
    }

    if ([string]::IsNullOrWhiteSpace($settings.sqlServer)) {
        Set-HttpConfigFieldError 'sqlServer' 'SQL server/data source is required.'
    }
    if ([string]::IsNullOrWhiteSpace($settings.sqlCatalog)) {
        Set-HttpConfigFieldError 'sqlCatalog' 'SQL initial catalog pattern is required.'
    }
    if (-not $settings.sqlIntegrated) {
        if ([string]::IsNullOrWhiteSpace($settings.sqlUser)) {
            Set-HttpConfigFieldError 'sqlUser' 'SQL user is required when integrated security is off.'
        }
        if ([string]::IsNullOrWhiteSpace($settings.sqlPassword)) {
            Set-HttpConfigFieldError 'sqlPassword' 'SQL password is required when integrated security is off.'
        }
    }

    if ([string]::IsNullOrWhiteSpace($settings.adminUser)) {
        Set-HttpConfigFieldError 'adminUser' 'Admin credential user is required.'
    }
    if ([string]::IsNullOrWhiteSpace($settings.adminPassword)) {
        Set-HttpConfigFieldError 'adminPassword' 'Admin credential password is required.'
    }

    if ([string]::IsNullOrWhiteSpace($settings.reportingUser)) {
        Set-HttpConfigFieldError 'reportingUser' 'Reporting credential is required.'
    }
    if ([string]::IsNullOrWhiteSpace($settings.reportingPassword)) {
        Set-HttpConfigFieldError 'reportingPassword' 'Reporting password is required.'
    }

    $uriOk = $false
    if (-not [string]::IsNullOrWhiteSpace($settings.reportServerUrl)) {
        $uriRef = $null
        $uriOk = [Uri]::TryCreate($settings.reportServerUrl, [System.UriKind]::Absolute, [ref]$uriRef)
        if ($uriOk) {
            $uriOk = ($uriRef.Scheme -eq 'http' -or $uriRef.Scheme -eq 'https')
        }
    }
    if (-not $uriOk) {
        Set-HttpConfigFieldError 'reportServerUrl' 'Report server URL must be absolute http/https.'
    }

    if ($settings.addClientVirtualizationFilter -and [string]::IsNullOrWhiteSpace($settings.clientVirtualizationFilter)) {
        Set-HttpConfigFieldError 'clientVirtualizationFilter' 'ClientVirtualizationFilter value is required when enabled.'
    }

    $recordingHosts = @($settings.recordingHosts -split ';' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $recordingAny = ($recordingHosts.Count -gt 0) -or -not [string]::IsNullOrWhiteSpace($settings.recordingUser) -or -not [string]::IsNullOrWhiteSpace($settings.recordingPassword) -or -not [string]::IsNullOrWhiteSpace($settings.recordingFolder)
    if ($recordingAny) {
        if ($recordingHosts.Count -eq 0) { Set-HttpConfigFieldError 'recordingHosts' 'Recording hosts are required for recording configuration.' }
        if ([string]::IsNullOrWhiteSpace($settings.recordingUser)) { Set-HttpConfigFieldError 'recordingUser' 'Recording user is required.' }
        if ([string]::IsNullOrWhiteSpace($settings.recordingPassword)) { Set-HttpConfigFieldError 'recordingPassword' 'Recording password is required.' }
        if ([string]::IsNullOrWhiteSpace($settings.recordingFolder)) { Set-HttpConfigFieldError 'recordingFolder' 'Recording folder is required.' }
    }

    if ($ShowInlineErrors) {
        Show-HttpConfigFieldErrors
    }

    $errorCount = ($script:HttpConfigValidationErrors.Keys | Measure-Object).Count
    if ($errorCount -eq 0) {
        $tbHttpConfigSummary.Text = 'HTTP.Config wizard validation: OK.'
    } else {
        $tbHttpConfigSummary.Text = "HTTP.Config wizard validation: $errorCount field(s) require attention."
    }

    return [pscustomobject]@{
        IsValid = ($errorCount -eq 0)
        Errors = $script:HttpConfigValidationErrors
        Settings = $settings
    }
}

function New-HttpConfigPreviewText {
    $result = Validate-HttpConfigSettings -ShowInlineErrors $true
    $s = $result.Settings
    $recordingValue = Build-RecordingUrlList -Settings $s
    $connectionString = if ($s.sqlIntegrated) {
        "Integrated Security=Yes;Data Source=$($s.sqlServer);Initial Catalog=$($s.sqlCatalog)"
    } else {
        "Integrated Security=No;User ID=$($s.sqlUser);Password=$($s.sqlPassword);Data Source=$($s.sqlServer);Initial Catalog=$($s.sqlCatalog)"
    }

    $lines = @()
    $lines += '<domain host="' + $s.hostList + '" connectionString="' + $connectionString + '">'
    $lines += '  <credentials>'
    $lines += '    <add key="admin" credential="' + $s.adminUser + ':********" roles="NixxisAdminRole"/>'
    $lines += '    <add credential="' + $s.reportingUser + ':********" roles="NixxisReportingRole"/>'
    $lines += '  </credentials>'
    $lines += '  <application id="admin">'
    $lines += '    <add key="reportServerUrl" value="' + $s.reportServerUrl + '"/>'
    if ($s.addClientVirtualizationFilter) {
        $lines += '    <add key="ClientVirtualizationFilter" value="' + $s.clientVirtualizationFilter + '"/>'
    }
    $lines += '  </application>'
    if ($s.addAgentReactivation) {
        $lines += '  <application id="acd">'
        $lines += '    <add key="agentReactivationEnabled" value="' + ($s.agentReactivationValue.ToString().ToLowerInvariant()) + '"/>'
        $lines += '  </application>'
    }
    if (-not [string]::IsNullOrWhiteSpace($recordingValue)) {
        $lines += '  <application id="relay">'
        $lines += '    <add key="recording" value="' + ($recordingValue -replace [regex]::Escape($s.recordingPassword), '********') + '"/>'
        $lines += '  </application>'
    }
    if ($s.addSupervision) {
        $lines += '  <application id="supervision" name="supervision" type="SupervisionApp" preload="true" debug="false">'
        $lines += '  </application>'
    }
    $lines += '</domain>'

    return ($lines -join [Environment]::NewLine)
}

function Show-HttpConfigPreview {
    $previewText = New-HttpConfigPreviewText
    $form = [System.Windows.Forms.Form]::new()
    $form.Text = 'HTTP.Config Preview (Install)'
    $form.Width = 960
    $form.Height = 620
    $form.StartPosition = 'CenterParent'

    $tb = [System.Windows.Forms.TextBox]::new()
    $tb.Multiline = $true
    $tb.ReadOnly = $true
    $tb.ScrollBars = 'Both'
    $tb.WordWrap = $false
    $tb.Dock = 'Fill'
    $tb.Font = [System.Drawing.Font]::new('Consolas', 10)
    $tb.Text = $previewText
    $form.Controls.Add($tb)
    $form.ShowDialog() | Out-Null
}

function Save-HttpConfigInstallProfile {
    $result = Validate-HttpConfigSettings -ShowInlineErrors $true
    $settings = $result.Settings

    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title = 'Save HTTP.Config Install Profile'
    $dlg.Filter = 'Encrypted Profile (*.ncsprofile)|*.ncsprofile|All Files (*.*)|*.*'
    $dlg.FileName = "HttpConfigProfile_$(Get-Date -Format 'yyyyMMdd_HHmmss').ncsprofile"
    if ($dlg.ShowDialog() -ne $true) { return }

    $json = ($settings | ConvertTo-Json -Depth 5 -Compress)
    $plainBytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $entropy = [System.Text.Encoding]::UTF8.GetBytes('NCS-HttpConfig-Profile-v1')
    $protected = [System.Security.Cryptography.ProtectedData]::Protect($plainBytes, $entropy, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
    [System.IO.File]::WriteAllText($dlg.FileName, [Convert]::ToBase64String($protected), [System.Text.Encoding]::UTF8)
    Add-LogEntry "Saved encrypted HTTP.Config install profile: $($dlg.FileName)" 'OK'
    $tbHttpConfigSummary.Text = 'HTTP.Config install profile saved.'
}

function Update-HttpConfigUiState {
    $sqlCredEnabled = -not [bool]$cbHttpSqlIntegrated.IsChecked
    $tbHttpSqlUser.IsEnabled = $sqlCredEnabled
    $pbHttpSqlPassword.IsEnabled = $sqlCredEnabled

    $tbHttpClientVirtualizationFilter.IsEnabled = [bool]$cbHttpAddClientVirtualizationFilter.IsChecked
    $cbHttpAgentReactivationValue.IsEnabled = [bool]$cbHttpAddAgentReactivation.IsChecked
}

function Add-PlanItem {
    param([string]$Badge, [string]$Action, [string]$Path, [string]$Details)
    $script:PlanItems.Add([pscustomobject]@{
        Badge = $Badge
        Action = $Action
        Path = $Path
        Details = $Details
    })
}

function Build-ExecutionPlan {
    $script:PlanItems.Clear()

    $modeText = if ($rbModeInstall.IsChecked) { 'Fresh Install' } else { 'Update Existing' }
    $sourceMode = if ($rbOffline.IsChecked) { 'Offline' } elseif ($rbOnlineCustom.IsChecked) { 'Online Custom URLs' } elseif ($rbOnlineSelect.IsChecked) { 'Online Selected Versions' } else { 'Online Auto' }
    $workDir = $tbWorkDir.Text.Trim()
    $runDir = Join-Path $workDir (Get-Date -Format 'yyyyMMdd')
    $installPath = $tbInstallPath.Text.Trim().TrimEnd('\\')
    $serviceName = $tbServiceName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($serviceName)) { $serviceName = 'crappserver' }

    Add-PlanItem 'CREATE' 'Ensure run directory' $runDir "Mode: $modeText | Source: $sourceMode"

    if ($rbOffline.IsChecked) {
        Add-PlanItem 'SKIP' 'Download from internet' $runDir 'Offline mode selected.'
        Add-PlanItem 'OVERWRITE' 'Copy ZIP files' $tbOfflinePath.Text.Trim() 'ClientProvisioning.zip, ClientSoftware.zip, NCS.zip'
    } elseif ($rbOnlineSelect.IsChecked) {
        $clientV = if ($cbClientVersion.SelectedItem) { [string]$cbClientVersion.SelectedItem } else { '(not selected)' }
        $serverV = if ($cbServerVersion.SelectedItem) { [string]$cbServerVersion.SelectedItem } else { '(not selected)' }
        Add-PlanItem 'OVERWRITE' 'Download selected client ZIPs' $runDir "Client version: $clientV"
        Add-PlanItem 'OVERWRITE' 'Download selected server ZIP' $runDir "Server version: $serverV"
    } else {
        Add-PlanItem 'OVERWRITE' 'Download ZIP files' $runDir 'Auto/custom URL resolution for ClientProvisioning, ClientSoftware, NCS'
    }

    Add-PlanItem 'OVERWRITE' 'Extract staging files' (Join-Path $runDir 'NixxisApplicationServer') 'NCS, ClientSoftware, provisioning payloads'

    if ($rbModeInstall.IsChecked) {
        Add-PlanItem 'CHECK' 'Apply HTTP.Config wizard' (Join-Path $runDir 'NixxisApplicationServer\SampleConfigFiles') 'Updates staged http.config sample values before deploy'
    }

    if ($rbModeUpdate.IsChecked) {
        Add-PlanItem 'OVERWRITE' 'Stop service and backup existing' $installPath "Service: $serviceName"
        Add-PlanItem 'DELETE' 'Cleanup old files' (Join-Path $installPath 'CrAppServer') 'Deletes obsolete DLL/PDB/provisioning artifacts'
    } else {
        Add-PlanItem 'CREATE' 'Create install root' $installPath 'Fresh install path initialization'
    }

    Add-PlanItem 'OVERWRITE' 'Deploy application folders' $installPath 'ClientSoftware, CrAppServer, MediaServer, Reporting, SoundsSamples'
    if ($rbModeInstall.IsChecked) {
        Add-PlanItem 'OVERWRITE' 'Process SampleConfigFiles' (Join-Path $installPath 'CrAppServer') 'Copy all except NCC*, strip .sample extension'
    } else {
        Add-PlanItem 'CHECK' 'Review config manually' (Join-Path $installPath 'CrAppServer\CrAppserver.exe.config') 'Make sure CrAppserver.exe.config is up to date. Adapt as needed'
    }

    if ($rbModeInstall.IsChecked) {
        Add-PlanItem 'CREATE' 'Install Windows service' $serviceName "Path: $(Join-Path $installPath 'CrAppServer\\CrAppServer.exe')"
        if ([bool]$cbConfigureFirewall.IsChecked) {
            Add-PlanItem 'CREATE' 'Create firewall rules' 'Inbound TCP/UDP' "Rule base: Nixxis $serviceName"
        } else {
            Add-PlanItem 'SKIP' 'Configure firewall rules' 'Inbound TCP/UDP' 'Option disabled'
        }
        if ([bool]$cbInstallMoveFiles.IsChecked) {
            Add-PlanItem 'CREATE' 'Install MoveFiles service' (Join-Path $installPath 'Tools\\MoveFiles\\MoveFiles.exe') 'Will create/rename MoveFiles.xml'
        } else {
            Add-PlanItem 'SKIP' 'Install MoveFiles service' (Join-Path $installPath 'Tools\\MoveFiles') 'Option disabled'
        }
    }

    $tbPlanSummary.Text = "Plan generated: $($script:PlanItems.Count) actions | Service: $serviceName"
    $chkPlanApproved.IsChecked = $false
    $tcOps.SelectedIndex = 1
}

function Ensure-PlanApproved {
    if ($script:PlanItems.Count -eq 0) {
        Add-LogEntry 'Execution blocked: generate plan first.' 'WARN'
        $tcOps.SelectedIndex = 1
        return $false
    }
    if (-not [bool]$chkPlanApproved.IsChecked) {
        Add-LogEntry 'Execution blocked: approval gate not checked (I reviewed and approve this plan).' 'WARN'
        $tcOps.SelectedIndex = 1
        return $false
    }
    return $true
}

function Add-RiskItem {
    param([string]$Status, [string]$Risk, [string]$Details)
    $script:PreflightItems.Add([pscustomobject]@{
        Status = $Status
        Risk = $Risk
        Details = $Details
    })
}

function Run-PreflightChecks {
    $script:PreflightItems.Clear()

    $admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    Add-RiskItem ($(if ($admin) { 'GREEN' } else { 'RED' })) 'Administrator privileges' ($(if ($admin) { 'OK' } else { 'Run PowerShell as Administrator.' }))

    $installPath = $tbInstallPath.Text.Trim().TrimEnd('\\')
    Add-RiskItem ($(if (Test-Path $installPath) { 'GREEN' } else { 'YELLOW' })) 'Install path exists' ($(if (Test-Path $installPath) { $installPath } else { "$installPath will be created." }))

    $driveRoot = [System.IO.Path]::GetPathRoot($installPath)
    try {
        $drive = Get-PSDrive -Name ($driveRoot.TrimEnd(':\\')) -ErrorAction Stop
        $freeGb = [math]::Round($drive.Free / 1GB, 2)
        $status = if ($freeGb -lt 5) { 'RED' } elseif ($freeGb -lt 10) { 'YELLOW' } else { 'GREEN' }
        Add-RiskItem $status 'Free disk space' "$freeGb GB available on $driveRoot"
    } catch {
        Add-RiskItem 'YELLOW' 'Disk space check' "Could not evaluate free space on $driveRoot"
    }

    $dotNetRelease = 0
    $regPath = 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full'
    if (Test-Path $regPath) {
        $dotNetRelease = [int]((Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).Release)
    }
    Add-RiskItem ($(if ($dotNetRelease -ge 528040) { 'GREEN' } else { 'YELLOW' })) '.NET Framework 4.8' "Release key: $dotNetRelease"

    if ($rbOffline.IsChecked) {
        $off = $tbOfflinePath.Text.Trim()
        $allPresent = $true
        foreach ($z in @('ClientProvisioning.zip','ClientSoftware.zip','NCS.zip')) {
            if (-not (Test-Path (Join-Path $off $z))) { $allPresent = $false }
        }
        Add-RiskItem ($(if ($allPresent) { 'GREEN' } else { 'RED' })) 'Offline ZIP availability' ($(if ($allPresent) { 'All required ZIP files found.' } else { 'Missing one or more required ZIP files.' }))
    } else {
        try {
            Invoke-WebRequest -Uri 'http://update.nixxis.net' -UseBasicParsing -TimeoutSec 10 | Out-Null
            Add-RiskItem 'GREEN' 'Online source reachability' 'http://update.nixxis.net reachable.'
        } catch {
            Add-RiskItem 'RED' 'Online source reachability' 'Cannot reach http://update.nixxis.net'
        }
    }

    if ($rbOnlineSelect.IsChecked) {
        $selectedOk = $cbClientVersion.SelectedItem -and $cbServerVersion.SelectedItem
        Add-RiskItem ($(if ($selectedOk) { 'GREEN' } else { 'YELLOW' })) 'Selected online versions' ($(if ($selectedOk) { "Client=$($cbClientVersion.SelectedItem), Server=$($cbServerVersion.SelectedItem)" } else { 'Run Discover and select both versions.' }))
    }

    $reds = ($script:PreflightItems | Where-Object Status -eq 'RED').Count
    $yellows = ($script:PreflightItems | Where-Object Status -eq 'YELLOW').Count
    $greens = ($script:PreflightItems | Where-Object Status -eq 'GREEN').Count
    $tbPreflightSummary.Text = "GREEN=$greens  YELLOW=$yellows  RED=$reds"
    $tcOps.SelectedIndex = 2
}

function Initialize-Timeline {
    param([string]$JobName)

    $script:TimelineItems.Clear()
    $script:TimelineCurrentIndex = -1
    $script:TimelineStartedAt = Get-Date

    if ($JobName -eq 'Initial Setup') {
        $script:TimelineStepOrder = @('DOWNLOAD PHASE','PREPARE / EXTRACT PHASE','STOP SERVICE PHASE','BACKUP PHASE','CLEANUP PHASE','HTTP.CONFIG CONFIGURATION PHASE','DEPLOY PHASE','SERVICE INSTALLATION PHASE','TOOLS COPY PHASE','MOVEFILES INSTALL PHASE','REPORTING USER PHASE','FIREWALL PHASE','TRANSCRIPTION HELPERS PHASE')
    } elseif ($JobName -eq 'Full Update') {
        $script:TimelineStepOrder = @('DOWNLOAD PHASE','PREPARE / EXTRACT PHASE','STOP SERVICE PHASE','BACKUP PHASE','CLEANUP PHASE','DEPLOY PHASE')
    } else {
        $script:TimelineStepOrder = @($JobName.ToUpper())
    }

    foreach ($s in $script:TimelineStepOrder) {
        $script:TimelineItems.Add([pscustomobject]@{ Step = $s; State = 'Pending'; Started = ''; Finished = ''; Notes = '' })
    }

    $tbTimelineCurrent.Text = 'Current: Pending'
    $tbTimelineNext.Text = "Next: $($script:TimelineStepOrder[0])"
    $tbTimelineEta.Text = "ETA: ~$($script:TimelineStepOrder.Count * 45)s"
}

function Update-TimelineFromLog {
    param([string]$Message)

    if ([string]::IsNullOrWhiteSpace($Message)) { return }
    $normalizedMessage = $Message.Trim()

    $findStepIndex = {
        param([string]$Name)
        for ($j = 0; $j -lt $script:TimelineStepOrder.Count; $j++) {
            if ($script:TimelineStepOrder[$j] -eq $Name) { return $j }
        }
        return -1
    }

    if ($normalizedMessage -match '===\s*(.+?)\s*===') {
        $step = $Matches[1].Trim().ToUpper()

        $idx = & $findStepIndex $step
        if ($idx -lt 0 -and $step -notmatch '\s+PHASE$') {
            $idx = & $findStepIndex ($step + ' PHASE')
            if ($idx -ge 0) {
                $step = $step + ' PHASE'
            }
        }
        if ($idx -lt 0) { return }

        if ($script:TimelineCurrentIndex -ge 0 -and $script:TimelineCurrentIndex -lt $script:TimelineItems.Count) {
            $prev = $script:TimelineItems[$script:TimelineCurrentIndex]
            if ($prev.State -eq 'In Progress') {
                $prev.State = 'Done'
                $prev.Finished = (Get-Date).ToString('HH:mm:ss')
                $script:TimelineItems[$script:TimelineCurrentIndex] = $prev
            }
        }

        $current = $script:TimelineItems[$idx]
        $current.State = 'In Progress'
        $current.Started = (Get-Date).ToString('HH:mm:ss')
        $current.Notes = 'Running'
        $script:TimelineItems[$idx] = $current
        $script:TimelineCurrentIndex = $idx

        $tbTimelineCurrent.Text = "Current: $step"
        $nextIdx = $idx + 1
        if ($nextIdx -lt $script:TimelineStepOrder.Count) {
            $tbTimelineNext.Text = "Next: $($script:TimelineStepOrder[$nextIdx])"
        } else {
            $tbTimelineNext.Text = 'Next: Completed'
        }

        $remainingSteps = [Math]::Max(0, $script:TimelineStepOrder.Count - $idx - 1)
        $tbTimelineEta.Text = "ETA: ~$($remainingSteps * 45)s"
        return
    }

    if ($normalizedMessage -match '^(.+?)\s+phase complete\.$') {
        $completedStep = ($Matches[1].Trim().ToUpper() + ' PHASE')
        $doneIdx = & $findStepIndex $completedStep
        if ($doneIdx -lt 0) { return }

        $doneItem = $script:TimelineItems[$doneIdx]
        $doneItem.State = 'Done'
        if (-not $doneItem.Started) { $doneItem.Started = (Get-Date).ToString('HH:mm:ss') }
        $doneItem.Finished = (Get-Date).ToString('HH:mm:ss')
        $doneItem.Notes = 'Completed'
        $script:TimelineItems[$doneIdx] = $doneItem
        if ($doneIdx -gt $script:TimelineCurrentIndex) { $script:TimelineCurrentIndex = $doneIdx }

        $tbTimelineCurrent.Text = "Current: $completedStep"
        $nextDoneIdx = $doneIdx + 1
        if ($nextDoneIdx -lt $script:TimelineStepOrder.Count) {
            $tbTimelineNext.Text = "Next: $($script:TimelineStepOrder[$nextDoneIdx])"
        } else {
            $tbTimelineNext.Text = 'Next: Completed'
        }

        $remainingDoneSteps = [Math]::Max(0, $script:TimelineStepOrder.Count - $doneIdx - 1)
        $tbTimelineEta.Text = "ETA: ~$($remainingDoneSteps * 45)s"
        return
    }

    if ($normalizedMessage -match '^Service status:\s*Running\b') {
        $serviceIdx = & $findStepIndex 'START SERVICE'
        if ($serviceIdx -lt 0) { return }

        $svcItem = $script:TimelineItems[$serviceIdx]
        if (-not $svcItem.Started) { $svcItem.Started = (Get-Date).ToString('HH:mm:ss') }
        $svcItem.State = 'Done'
        $svcItem.Finished = (Get-Date).ToString('HH:mm:ss')
        $svcItem.Notes = 'Completed'
        $script:TimelineItems[$serviceIdx] = $svcItem
        $script:TimelineCurrentIndex = $serviceIdx
        $tbTimelineCurrent.Text = 'Current: START SERVICE'
        $tbTimelineNext.Text = 'Next: Completed'
        $tbTimelineEta.Text = 'ETA: 0s'
        return
    }
}

function Finalize-Timeline {
    if ($script:TimelineCurrentIndex -lt 0 -and $script:TimelineItems.Count -eq 1) {
        $single = $script:TimelineItems[0]
        if ($single.State -eq 'Pending') {
            $single.State = 'Done'
            $single.Started = (Get-Date).ToString('HH:mm:ss')
            $single.Finished = (Get-Date).ToString('HH:mm:ss')
            $single.Notes = 'Completed'
            $script:TimelineItems[0] = $single
            $script:TimelineCurrentIndex = 0
        }
    }

    if ($script:TimelineCurrentIndex -ge 0 -and $script:TimelineCurrentIndex -lt $script:TimelineItems.Count) {
        $last = $script:TimelineItems[$script:TimelineCurrentIndex]
        if ($last.State -eq 'In Progress') {
            $last.State = 'Done'
            $last.Finished = (Get-Date).ToString('HH:mm:ss')
            $last.Notes = 'Completed'
            $script:TimelineItems[$script:TimelineCurrentIndex] = $last
        }
    }
    $tbTimelineCurrent.Text = 'Current: Completed'
    $tbTimelineNext.Text = 'Next: -'
    $tbTimelineEta.Text = 'ETA: 0s'
}

function Add-ValidationItem {
    param([string]$Status, [string]$Check, [string]$Details)
    $script:ValidationItems.Add([pscustomobject]@{
        Status = $Status
        Check = $Check
        Details = $Details
    })
}

function Run-PostInstallValidation {
    $script:ValidationItems.Clear()

    $installPath = $tbInstallPath.Text.Trim().TrimEnd('\\')
    $serviceName = $tbServiceName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($serviceName)) { $serviceName = 'crappserver' }

    Add-ValidationItem ($(if (Test-Path $installPath) { 'PASS' } else { 'FAIL' })) 'Install path exists' $installPath
    Add-ValidationItem ($(if (Test-Path (Join-Path $installPath 'CrAppServer\\CrAppServer.exe')) { 'PASS' } else { 'FAIL' })) 'CrAppServer executable present' (Join-Path $installPath 'CrAppServer\\CrAppServer.exe')
    Add-ValidationItem ($(if (Test-Path (Join-Path $installPath 'ClientSoftware')) { 'PASS' } else { 'FAIL' })) 'ClientSoftware folder present' (Join-Path $installPath 'ClientSoftware')
    Add-ValidationItem ($(if (Test-Path (Join-Path $installPath 'Logs')) { 'PASS' } else { 'WARN' })) 'Logs folder present' (Join-Path $installPath 'Logs')

    $svc = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
    if ($svc) {
        Add-ValidationItem 'PASS' 'Service exists' "$serviceName ($($svc.Status))"
        Add-ValidationItem ($(if ($svc.Status -eq 'Running') { 'PASS' } else { 'WARN' })) 'Service running' "$serviceName status: $($svc.Status)"
    } else {
        Add-ValidationItem 'FAIL' 'Service exists' "$serviceName not found"
    }

    $tcpRule = Get-NetFirewallRule -DisplayName "Nixxis $serviceName TCP" -ErrorAction SilentlyContinue
    $udpRule = Get-NetFirewallRule -DisplayName "Nixxis $serviceName UDP" -ErrorAction SilentlyContinue
    Add-ValidationItem ($(if ($tcpRule -and $udpRule) { 'PASS' } else { 'WARN' })) 'Firewall rules' "TCP/UDP rules for $serviceName"

    $pass = ($script:ValidationItems | Where-Object Status -eq 'PASS').Count
    $warn = ($script:ValidationItems | Where-Object Status -eq 'WARN').Count
    $fail = ($script:ValidationItems | Where-Object Status -eq 'FAIL').Count
    $tbValidationSummary.Text = "PASS=$pass  WARN=$warn  FAIL=$fail"
    $tcOps.SelectedIndex = 4
}

function Export-SignedOperationReport {
    $workRoot = $tbWorkDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($workRoot)) { $workRoot = 'C:\NixxisMaintenance\Update' }
    $reportsDir = Join-Path $workRoot 'Reports'
    if (-not (Test-Path $reportsDir)) {
        New-Item -Path $reportsDir -ItemType Directory -Force | Out-Null
    }

    $payload = [ordered]@{
        appName = 'NCS Installer / Updater'
        appVersion = $script:AppVersion
        generatedAt = (Get-Date).ToString('s')
        user = $env:USERNAME
        host = $env:COMPUTERNAME
        mode = if ($rbModeInstall.IsChecked) { 'Fresh Install' } else { 'Update Existing' }
        sourceMode = if ($rbOffline.IsChecked) { 'Offline' } elseif ($rbOnlineCustom.IsChecked) { 'Online Custom' } elseif ($rbOnlineSelect.IsChecked) { 'Online Select' } else { 'Online Auto' }
        installPath = $tbInstallPath.Text.Trim().TrimEnd('\\')
        serviceName = $tbServiceName.Text.Trim()
        selectedVersions = [ordered]@{
            base = $sync.SelectedBaseVersion
            client = $sync.SelectedClientVersion
            server = $sync.SelectedServerVersion
        }
        plan = @($script:PlanItems)
        preflight = @($script:PreflightItems)
        timeline = @($script:TimelineItems)
        validation = @($script:ValidationItems)
        lastStatus = $tbStatus.Text
    }

    $payloadJson = $payload | ConvertTo-Json -Depth 8
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($payloadJson)
        $hash = [BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '').ToLowerInvariant()
    } finally {
        $sha.Dispose()
    }

    $report = [ordered]@{
        payload = $payload
        signature = [ordered]@{
            algorithm = 'SHA256'
            signedAt = (Get-Date).ToString('s')
            signedBy = "$env:USERNAME@$env:COMPUTERNAME"
            digest = $hash
        }
    }

    $reportFile = Join-Path $reportsDir ("NCS_OperationReport_{0}.json" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $report | ConvertTo-Json -Depth 10 | Set-Content -Path $reportFile -Encoding UTF8
    Add-LogEntry "Signed operation report exported: $reportFile" 'OK'
    return $reportFile
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
    Initialize-Timeline $JobName

    # Snapshot UI values for runspace
    $mode      = if ($rbOnlineAuto.IsChecked)    { 'online-auto' }
                 elseif ($rbOnlineSelect.IsChecked) { 'online-select' }
                 elseif ($rbOnlineCustom.IsChecked) { 'online-custom' }
                 else { 'offline' }
    $cpUrl     = $tbCPUrl.Text.Trim()
    $csUrl     = $tbCSUrl.Text.Trim()
    $ncsUrl    = $tbNCSUrl.Text.Trim()
    $offPath   = $tbOfflinePath.Text.Trim()
    $installPath = $tbInstallPath.Text.Trim().TrimEnd('\\')
    $serviceName = $tbServiceName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($serviceName)) { $serviceName = 'crappserver' }
    $selectedClientVersion = if ($cbClientVersion.SelectedItem) { [string]$cbClientVersion.SelectedItem } else { '' }
    $selectedServerVersion = if ($cbServerVersion.SelectedItem) { [string]$cbServerVersion.SelectedItem } else { '' }
    $selectedBaseVersion = if ($sync.SelectedBaseVersion) { [string]$sync.SelectedBaseVersion } else { '' }
    $operationMode = if ($rbModeInstall.IsChecked) { 'install' } else { 'update' }
    $httpConfigHostList = $tbHttpHostList.Text.Trim()
    $httpConfigSqlIntegrated = [bool]$cbHttpSqlIntegrated.IsChecked
    $httpConfigSqlServer = $tbHttpSqlServer.Text.Trim()
    $httpConfigSqlCatalog = $tbHttpSqlCatalog.Text.Trim()
    $httpConfigSqlUser = $tbHttpSqlUser.Text.Trim()
    $httpConfigSqlPassword = $pbHttpSqlPassword.Password
    $httpConfigAdminUser = $tbHttpAdminUser.Text.Trim()
    $httpConfigAdminPassword = $pbHttpAdminPassword.Password
    $httpConfigReportingUser = $tbHttpReportingUser.Text.Trim()
    $httpConfigReportingPassword = $pbHttpReportingPassword.Password
    $httpConfigReportServerUrl = $tbHttpReportServerUrl.Text.Trim()
    $httpConfigAddAgentReactivation = [bool]$cbHttpAddAgentReactivation.IsChecked
    $httpConfigAgentReactivationValue = [bool]$cbHttpAgentReactivationValue.IsChecked
    $httpConfigAddClientVirtualizationFilter = [bool]$cbHttpAddClientVirtualizationFilter.IsChecked
    $httpConfigClientVirtualizationFilter = $tbHttpClientVirtualizationFilter.Text.Trim()
    $httpConfigRecordingHosts = $tbHttpRecordingHosts.Text.Trim()
    $httpConfigRecordingUser = $tbHttpRecordingUser.Text.Trim()
    $httpConfigRecordingPassword = $pbHttpRecordingPassword.Password
    $httpConfigRecordingFolder = $tbHttpRecordingFolder.Text.Trim().Trim('/','\\')
    $httpConfigAddSupervision = [bool]$cbHttpAddSupervision.IsChecked
    $optEnsureDotNet48 = [bool]$cbEnsureDotNet48.IsChecked
    $optInstallMoveFiles = [bool]$cbInstallMoveFiles.IsChecked
    $optCreateReportingUser = [bool]$cbCreateReportingUser.IsChecked
    $optConfigureFirewall = [bool]$cbConfigureFirewall.IsChecked
    $optDeployTranscription = [bool]$cbDeployTranscription.IsChecked
    $runDir    = $sync.RunDir
    $sync.InstallPath = $installPath
    $sync.ServiceName = $serviceName
    $sync.SelectedClientVersion = $selectedClientVersion
    $sync.SelectedServerVersion = $selectedServerVersion

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
    $rs.SessionStateProxy.SetVariable('serviceName', $serviceName)
    $rs.SessionStateProxy.SetVariable('selectedClientVersion', $selectedClientVersion)
    $rs.SessionStateProxy.SetVariable('selectedServerVersion', $selectedServerVersion)
    $rs.SessionStateProxy.SetVariable('selectedBaseVersion', $selectedBaseVersion)
    $rs.SessionStateProxy.SetVariable('operationMode', $operationMode)
    $rs.SessionStateProxy.SetVariable('httpConfigHostList', $httpConfigHostList)
    $rs.SessionStateProxy.SetVariable('httpConfigSqlIntegrated', $httpConfigSqlIntegrated)
    $rs.SessionStateProxy.SetVariable('httpConfigSqlServer', $httpConfigSqlServer)
    $rs.SessionStateProxy.SetVariable('httpConfigSqlCatalog', $httpConfigSqlCatalog)
    $rs.SessionStateProxy.SetVariable('httpConfigSqlUser', $httpConfigSqlUser)
    $rs.SessionStateProxy.SetVariable('httpConfigSqlPassword', $httpConfigSqlPassword)
    $rs.SessionStateProxy.SetVariable('httpConfigAdminUser', $httpConfigAdminUser)
    $rs.SessionStateProxy.SetVariable('httpConfigAdminPassword', $httpConfigAdminPassword)
    $rs.SessionStateProxy.SetVariable('httpConfigReportingUser', $httpConfigReportingUser)
    $rs.SessionStateProxy.SetVariable('httpConfigReportingPassword', $httpConfigReportingPassword)
    $rs.SessionStateProxy.SetVariable('httpConfigReportServerUrl', $httpConfigReportServerUrl)
    $rs.SessionStateProxy.SetVariable('httpConfigAddAgentReactivation', $httpConfigAddAgentReactivation)
    $rs.SessionStateProxy.SetVariable('httpConfigAgentReactivationValue', $httpConfigAgentReactivationValue)
    $rs.SessionStateProxy.SetVariable('httpConfigAddClientVirtualizationFilter', $httpConfigAddClientVirtualizationFilter)
    $rs.SessionStateProxy.SetVariable('httpConfigClientVirtualizationFilter', $httpConfigClientVirtualizationFilter)
    $rs.SessionStateProxy.SetVariable('httpConfigRecordingHosts', $httpConfigRecordingHosts)
    $rs.SessionStateProxy.SetVariable('httpConfigRecordingUser', $httpConfigRecordingUser)
    $rs.SessionStateProxy.SetVariable('httpConfigRecordingPassword', $httpConfigRecordingPassword)
    $rs.SessionStateProxy.SetVariable('httpConfigRecordingFolder', $httpConfigRecordingFolder)
    $rs.SessionStateProxy.SetVariable('httpConfigAddSupervision', $httpConfigAddSupervision)
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

    # DispatcherTimer - runs on UI thread, safe to touch controls directly
    $timer          = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [timespan]::FromMilliseconds(200)

    # Use a script-level variable to hold closure state
    $script:_job = @{ ps = $ps; rs = $rs; handle = $handle; timer = $timer; start = $startTime; name = $JobName }

    $timer.Add_Tick({
        $j = $script:_job

        # Drain log queue - already on UI thread, call Add-LogEntry directly
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

            Finalize-Timeline
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
        Write-BgLog "Offline mode - source: $src" 'MAGENTA'
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

        if ($mode -eq 'online-select') {
            if (-not $selectedClientVersion -or -not $selectedServerVersion) {
                throw 'No client/server version selected. Use Discover Available Versions first.'
            }

            $selectedBase = if ($selectedBaseVersion) { $selectedBaseVersion } else { Get-LatestFolder $baseUrl 'version' }
            $versionUrl = "$baseUrl/$selectedBase"
            $clientUrl = "$versionUrl/Client"
            $serverUrl = "$versionUrl/Server"

            if (-not $resolvedCP)  { $resolvedCP  = "$clientUrl/$selectedClientVersion/ClientProvisioning.zip" }
            if (-not $resolvedCS)  { $resolvedCS  = "$clientUrl/$selectedClientVersion/ClientSoftware.zip" }
            if (-not $resolvedNCS) { $resolvedNCS = "$serverUrl/$selectedServerVersion/NCS.zip" }
        }

        if ($mode -ne 'online-select' -and (-not $resolvedCP -or -not $resolvedCS -or -not $resolvedNCS)) {
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

    Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
    $max = 60; $i = 0
    do {
        $svc  = Get-Service $serviceName -ErrorAction SilentlyContinue
        $proc = Get-Process 'crappserver' -ErrorAction SilentlyContinue
        if ((-not $svc -or $svc.Status -eq 'Stopped') -and -not $proc) {
            Start-Sleep 3
            if (-not (Get-Process 'crappserver' -ErrorAction SilentlyContinue)) {
                Write-BgLog "$serviceName fully stopped." 'OK'; break
            }
        }
        $i++
        Write-BgLog "Waiting for service to stop... attempt $i / $max" 'WARN'
        Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
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
        } else { Write-BgLog "$src not found - skipping." 'WARN' }
    }
    Write-BgLog 'Backup phase complete.' 'OK'
}

$sbCleanup = {
    Write-BgLog '=== CLEANUP PHASE ===' 'HEADER'
    $base = Join-Path $installPath 'CrAppServer'
    if (-not (Test-Path $base)) {
        Write-BgLog "CrAppServer not found at $base - nothing to clean." 'WARN'
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
    } else { Write-BgLog "ClientSoftware not found - skipping." 'WARN' }
    Write-BgLog 'Cleanup phase complete.' 'OK'
}

$sbConfigureHttpConfig = {
    Write-BgLog '=== HTTP.CONFIG CONFIGURATION PHASE ===' 'HEADER'

    if ($operationMode -ne 'install') {
        Write-BgLog 'Update mode detected - skipped HTTP.Config wizard apply.' 'GRAY'
        return
    }

    try {
        $appServer = Join-Path $runDir 'NixxisApplicationServer'
        $sampleRoot = Join-Path $appServer 'SampleConfigFiles'
        if (-not (Test-Path $sampleRoot)) {
            Write-BgLog "SampleConfigFiles folder not found: $sampleRoot" 'WARN'
            return
        }

        $candidates = @(Get-ChildItem -Path $sampleRoot -File -Recurse | Where-Object {
            $_.Name -ieq 'http.config.sample' -or $_.Name -ieq 'http.config'
        })
        if ($candidates.Count -eq 0) {
            Write-BgLog 'No http.config sample file found in SampleConfigFiles - skipping.' 'WARN'
            return
        }

        $target = $null
        $sampleExact = $candidates | Where-Object { $_.Name -ieq 'http.config.sample' } | Select-Object -First 1
        if ($sampleExact) { $target = $sampleExact } else { $target = $candidates[0] }
        Write-BgLog "Applying HTTP.Config wizard values to: $($target.FullName)" 'CYAN'

        [xml]$doc = Get-Content -Path $target.FullName -Raw
        $httpNode = $doc.SelectSingleNode('/http')
        if (-not $httpNode) {
            throw 'Invalid http.config: /http node not found.'
        }
        $domain = $doc.SelectSingleNode('/http/domain')
        if (-not $domain) {
            $domain = $doc.CreateElement('domain')
            $domain.SetAttribute('id', 'Default')
            $domain.SetAttribute('domainType', 'http')
            $domain.SetAttribute('type', 'HttpProcessor')
            $httpNode.AppendChild($domain) | Out-Null
        }

        function Set-OrCreateDomainAttribute {
            param([string]$Name, [string]$Value)
            if (-not [string]::IsNullOrWhiteSpace($Value)) {
                $domain.SetAttribute($Name, $Value)
            }
        }

        function Ensure-ApplicationNode {
            param([string]$Id, [string]$Name, [string]$Type)

            $node = $domain.SelectSingleNode("application[@id='$Id']")
            if (-not $node) {
                $node = $doc.CreateElement('application')
                $node.SetAttribute('id', $Id)
                if ($Name) { $node.SetAttribute('name', $Name) }
                if ($Type) { $node.SetAttribute('type', $Type) }
                $domain.AppendChild($node) | Out-Null
            }
            return $node
        }

        function Set-ApplicationKey {
            param([System.Xml.XmlElement]$AppNode, [string]$Key, [string]$Value)

            if (-not $AppNode -or [string]::IsNullOrWhiteSpace($Key) -or $null -eq $Value) { return }
            $entry = $AppNode.SelectSingleNode("add[@key='$Key']")
            if (-not $entry) {
                $entry = $doc.CreateElement('add')
                $entry.SetAttribute('key', $Key)
                $AppNode.AppendChild($entry) | Out-Null
            }
            $entry.SetAttribute('value', $Value)
        }

        $hostValue = $httpConfigHostList
        if (-not [string]::IsNullOrWhiteSpace($hostValue)) {
            Set-OrCreateDomainAttribute -Name 'host' -Value $hostValue
        }

        if (-not [string]::IsNullOrWhiteSpace($httpConfigSqlServer) -and -not [string]::IsNullOrWhiteSpace($httpConfigSqlCatalog)) {
            $connectionString = if ($httpConfigSqlIntegrated) {
                "Integrated Security=Yes;Data Source=$httpConfigSqlServer;Initial Catalog=$httpConfigSqlCatalog"
            } else {
                "Integrated Security=No;User ID=$httpConfigSqlUser;Password=$httpConfigSqlPassword;Data Source=$httpConfigSqlServer;Initial Catalog=$httpConfigSqlCatalog"
            }
            Set-OrCreateDomainAttribute -Name 'connectionString' -Value $connectionString
        } else {
            Write-BgLog 'HTTP.Config: SQL server/catalog not fully provided - connectionString left unchanged.' 'WARN'
        }

        $credentialsNode = $domain.SelectSingleNode('credentials')
        if (-not $credentialsNode) {
            $credentialsNode = $doc.CreateElement('credentials')
            $domain.AppendChild($credentialsNode) | Out-Null
        }

        if (-not [string]::IsNullOrWhiteSpace($httpConfigAdminUser) -and -not [string]::IsNullOrWhiteSpace($httpConfigAdminPassword)) {
            $adminNode = $credentialsNode.SelectSingleNode("add[@key='admin']")
            if (-not $adminNode) {
                $adminNode = $doc.CreateElement('add')
                $adminNode.SetAttribute('key', 'admin')
                $credentialsNode.AppendChild($adminNode) | Out-Null
            }
            $adminNode.SetAttribute('credential', "$httpConfigAdminUser`:$httpConfigAdminPassword")
            $adminNode.SetAttribute('roles', 'NixxisAdminRole')
        } else {
            Write-BgLog 'HTTP.Config: admin credential incomplete - existing admin credential left unchanged.' 'WARN'
        }

        if (-not [string]::IsNullOrWhiteSpace($httpConfigReportingUser) -and -not [string]::IsNullOrWhiteSpace($httpConfigReportingPassword)) {
            $reportingNode = $credentialsNode.SelectSingleNode("add[@roles='NixxisReportingRole']")
            if (-not $reportingNode) {
                $reportingNode = $doc.CreateElement('add')
                $credentialsNode.AppendChild($reportingNode) | Out-Null
            }
            $reportingNode.SetAttribute('credential', "$httpConfigReportingUser`:$httpConfigReportingPassword")
            $reportingNode.SetAttribute('roles', 'NixxisReportingRole')
        } else {
            Write-BgLog 'HTTP.Config: reporting credential incomplete - existing reporting credential left unchanged.' 'WARN'
        }

        $adminApp = Ensure-ApplicationNode -Id 'admin' -Name 'CrAdmin' -Type 'NixxisAdminApp'
        if (-not [string]::IsNullOrWhiteSpace($httpConfigReportServerUrl)) {
            Set-ApplicationKey -AppNode $adminApp -Key 'reportServerUrl' -Value $httpConfigReportServerUrl
        }
        if ($httpConfigAddClientVirtualizationFilter -and -not [string]::IsNullOrWhiteSpace($httpConfigClientVirtualizationFilter)) {
            Set-ApplicationKey -AppNode $adminApp -Key 'ClientVirtualizationFilter' -Value $httpConfigClientVirtualizationFilter
        }

        if ($httpConfigAddAgentReactivation) {
            $acdApp = Ensure-ApplicationNode -Id 'acd' -Name 'CrAcd' -Type 'NixxisAcdApp'
            $agentReactivation = if ($httpConfigAgentReactivationValue) { 'true' } else { 'false' }
            Set-ApplicationKey -AppNode $acdApp -Key 'agentReactivationEnabled' -Value $agentReactivation
        }

        $recordingHosts = @($httpConfigRecordingHosts -split ';' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $recordingConfigured = ($recordingHosts.Count -gt 0) -and -not [string]::IsNullOrWhiteSpace($httpConfigRecordingUser) -and -not [string]::IsNullOrWhiteSpace($httpConfigRecordingPassword) -and -not [string]::IsNullOrWhiteSpace($httpConfigRecordingFolder)
        if ($recordingConfigured) {
            $relayApp = $domain.SelectSingleNode("application[@id='relay']")
            if (-not $relayApp) {
                $relayApp = Ensure-ApplicationNode -Id 'relay' -Name 'relay' -Type 'HttpRelay'
                $relayApp.SetAttribute('sessionLess', 'true')
                $relayApp.SetAttribute('serviceId', 'recording')
                $relayApp.SetAttribute('debug', 'true')
            }
            $folder = $httpConfigRecordingFolder.Trim('/','\\')
            $recUrls = @($recordingHosts | ForEach-Object {
                ('ftp://{0}:{1}@{2}/{3}/' -f $httpConfigRecordingUser, $httpConfigRecordingPassword, $_, $folder)
            })
            Set-ApplicationKey -AppNode $relayApp -Key 'recording' -Value ($recUrls -join ';')
        } elseif (($recordingHosts.Count -gt 0) -or -not [string]::IsNullOrWhiteSpace($httpConfigRecordingFolder)) {
            Write-BgLog 'HTTP.Config: recording parameters incomplete - recording key left unchanged.' 'WARN'
        }

        if ($httpConfigAddSupervision) {
            $supervisionApp = $domain.SelectSingleNode("application[@id='supervision']")
            if (-not $supervisionApp) {
                $supervisionApp = $doc.CreateElement('application')
                $supervisionApp.SetAttribute('id', 'supervision')
                $supervisionApp.SetAttribute('name', 'supervision')
                $supervisionApp.SetAttribute('type', 'SupervisionApp')
                $supervisionApp.SetAttribute('preload', 'true')
                $supervisionApp.SetAttribute('debug', 'false')
                $domain.AppendChild($supervisionApp) | Out-Null
            }
        }

        $settings = [System.Xml.XmlWriterSettings]::new()
        $settings.Indent = $true
        $settings.Encoding = [System.Text.UTF8Encoding]::new($false)
        $writer = [System.Xml.XmlWriter]::Create($target.FullName, $settings)
        try {
            $doc.Save($writer)
        }
        finally {
            $writer.Dispose()
        }

        Write-BgLog 'HTTP.Config wizard values applied to staged sample file.' 'OK'
    } catch {
        Write-BgLog "HTTP.Config wizard apply failed: $_" 'ERROR'
        Write-BgLog 'Fresh install will continue with existing staged config values.' 'WARN'
    }
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
        } else { Write-BgLog "Source not found in staging: $($f.S) - skipping." 'WARN' }
    }

    if ($operationMode -eq 'install') {
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
    } else {
        Write-BgLog 'Make sure CrAppserver.exe.config is up to date. Adapt as needed' 'WARN'
    }

    Write-BgLog 'Deploy phase complete.' 'OK'
}

$sbStartService = {
    Write-BgLog '=== START SERVICE ===' 'HEADER'
    $svc = Get-Service $serviceName -ErrorAction SilentlyContinue
    if (-not $svc) { throw "Service '$serviceName' not found on this machine." }
    if ($svc.Status -ne 'Running') {
        Start-Service $serviceName
    }

    $deadline = (Get-Date).AddMinutes(4)
    do {
        Start-Sleep -Seconds 2
        $svc.Refresh()
        if ($svc.Status -eq 'Running') { break }
    } while ((Get-Date) -lt $deadline)

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

    $existingService = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
    if ($existingService) {
        Write-BgLog "Service '$serviceName' already exists - skipping install." 'WARN'
        return
    }

    if ($serviceName -ieq 'crappserver') {
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
        Write-BgLog "Installed service using CrAppServer installer: $serviceName" 'OK'
    } else {
        $binPath = "`"$crAppServerExe`""
        New-Service -Name $serviceName -BinaryPathName $binPath -DisplayName $serviceName -StartupType Automatic | Out-Null
        Write-BgLog "Installed custom service '$serviceName' using New-Service." 'OK'
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
        $ruleName = "Nixxis $serviceName $proto"
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

function New-InitialSetupScriptBlock {
    $initialSetupText = @(
        'if ($optEnsureDotNet48) {'
        $sbEnsureDotNet48.ToString()
        '} else { Write-BgLog ''Skipped .NET 4.8 check by option.'' ''GRAY'' }'
        $sbDownload.ToString()
        $sbPrepare.ToString()
        $sbStopService.ToString()
        $sbBackup.ToString()
        $sbCleanup.ToString()
        $sbConfigureHttpConfig.ToString()
        $sbDeploy.ToString()
        $sbInstallService.ToString()
        $sbCopyTools.ToString()
        'if ($optInstallMoveFiles) {'
        $sbInstallMoveFiles.ToString()
        '} else { Write-BgLog ''Skipped MoveFiles installation by option.'' ''GRAY'' }'
        'if ($optCreateReportingUser) {'
        $sbCreateReportingUser.ToString()
        '} else { Write-BgLog ''Skipped Reporting user creation by option.'' ''GRAY'' }'
        'if ($optConfigureFirewall) {'
        $sbConfigFirewall.ToString()
        '} else { Write-BgLog ''Skipped firewall configuration by option.'' ''GRAY'' }'
        'if ($optDeployTranscription) {'
        $sbDeployTranscription.ToString()
        '} else { Write-BgLog ''Skipped transcription helpers by option.'' ''GRAY'' }'
        'Write-BgLog ''=== INITIAL SETUP COMPLETE ==='' ''OK'''
    ) -join "`n"

    return [scriptblock]::Create($initialSetupText)
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
$rbModeInstall.Add_Checked({ Update-OperationModeUI })
$rbModeUpdate.Add_Checked({ Update-OperationModeUI })
$rbOnlineSelect.Add_Checked({ Update-SourceModeUI })
$rbOnlineCustom.Add_Checked({ Update-SourceModeUI })
$rbOffline.Add_Checked({ Update-SourceModeUI })
$rbOnlineAuto.Add_Checked({ Update-SourceModeUI })

$btnDiscoverVersions.Add_Click({
    try {
        $btnDiscoverVersions.IsEnabled = $false
        Set-Status 'Discovering online versions...' 12
        Add-LogEntry 'Discovering available online client/server versions...' 'CYAN'

        $baseUrl = 'http://update.nixxis.net'
        $baseFolders = Get-DiscoveredWebFolders -Url $baseUrl
        $selectedBase = Get-LatestVersionValue -Values $baseFolders
        if (-not $selectedBase) {
            throw 'Could not discover a base version from update.nixxis.net'
        }

        $sync.SelectedBaseVersion = $selectedBase
        $tbDiscoveredBaseVersion.Text = "Base version: $selectedBase"

        $clientFolders = Get-DiscoveredWebFolders -Url "$baseUrl/$selectedBase/Client"
        $serverFolders = Get-DiscoveredWebFolders -Url "$baseUrl/$selectedBase/Server"

        $clientSorted = $clientFolders | Sort-Object {
            $v = $_ -replace '[^\d\.]', ''
            try { [version]$v } catch { [version]'0.0' }
        }
        $serverSorted = $serverFolders | Sort-Object {
            $v = $_ -replace '[^\d\.]', ''
            try { [version]$v } catch { [version]'0.0' }
        }

        $cbClientVersion.Items.Clear()
        foreach ($v in $clientSorted) { [void]$cbClientVersion.Items.Add($v) }
        $cbServerVersion.Items.Clear()
        foreach ($v in $serverSorted) { [void]$cbServerVersion.Items.Add($v) }

        if ($cbClientVersion.Items.Count -gt 0) { $cbClientVersion.SelectedIndex = $cbClientVersion.Items.Count - 1 }
        if ($cbServerVersion.Items.Count -gt 0) { $cbServerVersion.SelectedIndex = $cbServerVersion.Items.Count - 1 }

        Add-LogEntry "Discovered base version: $selectedBase" 'OK'
        Add-LogEntry "Client versions found: $($cbClientVersion.Items.Count)" 'OK'
        Add-LogEntry "Server versions found: $($cbServerVersion.Items.Count)" 'OK'
        Set-Status 'Version discovery complete.' 100
    }
    catch {
        Add-LogEntry "Version discovery failed: $($_.Exception.Message)" 'ERROR'
        Set-Status 'Version discovery failed.' 0
    }
    finally {
        $btnDiscoverVersions.IsEnabled = $true
    }
})

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

# WorkDir text change - keep sync live
$tbWorkDir.Add_TextChanged({
    $sync.WorkDir = $tbWorkDir.Text.Trim()
    Update-RunDirLabel
})

$tbInstallPath.Add_TextChanged({
    $sync.InstallPath = $tbInstallPath.Text.Trim().TrimEnd('\\')
})

$cbClientVersion.Add_SelectionChanged({
    if ($cbClientVersion.SelectedItem) {
        $sync.SelectedClientVersion = [string]$cbClientVersion.SelectedItem
    }
})

$cbServerVersion.Add_SelectionChanged({
    if ($cbServerVersion.SelectedItem) {
        $sync.SelectedServerVersion = [string]$cbServerVersion.SelectedItem
    }
})

$tbServiceName.Add_TextChanged({
    $name = $tbServiceName.Text.Trim()
    if (-not [string]::IsNullOrWhiteSpace($name)) {
        $sync.ServiceName = $name
    }
    Update-ServiceStatus
})

$btnHttpValidate.Add_Click({
    $validation = Validate-HttpConfigSettings -ShowInlineErrors $true
    if ($validation.IsValid) {
        Add-LogEntry 'HTTP.Config wizard validation completed: no field errors.' 'OK'
    } else {
        Add-LogEntry "HTTP.Config wizard validation completed: $($validation.Errors.Count) field(s) need attention." 'WARN'
    }
})

$btnHttpPreview.Add_Click({ Show-HttpConfigPreview })
$btnHttpSaveProfile.Add_Click({ Save-HttpConfigInstallProfile })

foreach ($tb in @($tbHttpHostList,$tbHttpSqlServer,$tbHttpSqlCatalog,$tbHttpSqlUser,$tbHttpAdminUser,$tbHttpReportingUser,$tbHttpReportServerUrl,$tbHttpClientVirtualizationFilter,$tbHttpRecordingHosts,$tbHttpRecordingUser,$tbHttpRecordingFolder)) {
    $tb.Add_TextChanged({ Validate-HttpConfigSettings -ShowInlineErrors $true | Out-Null })
}
foreach ($pb in @($pbHttpSqlPassword,$pbHttpAdminPassword,$pbHttpReportingPassword,$pbHttpRecordingPassword)) {
    $pb.Add_PasswordChanged({ Validate-HttpConfigSettings -ShowInlineErrors $true | Out-Null })
}
foreach ($cb in @($cbHttpSqlIntegrated,$cbHttpAddAgentReactivation,$cbHttpAgentReactivationValue,$cbHttpAddClientVirtualizationFilter,$cbHttpAddSupervision)) {
    $cb.Add_Checked({ Update-HttpConfigUiState; Validate-HttpConfigSettings -ShowInlineErrors $true | Out-Null })
    $cb.Add_Unchecked({ Update-HttpConfigUiState; Validate-HttpConfigSettings -ShowInlineErrors $true | Out-Null })
}

# Offline folder picker
$btnBrowseOffline.Add_Click({
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description = 'Select folder containing the 3 Nixxis ZIP files'
    if ($dlg.ShowDialog() -eq 'OK') { $tbOfflinePath.Text = $dlg.SelectedPath }
})

$btnRefreshStatus.Add_Click({ Update-ServiceStatus })

$btnGeneratePlan.Add_Click({ Build-ExecutionPlan })
$btnRunPreflight.Add_Click({ Run-PreflightChecks })
$btnRunValidation.Add_Click({ Run-PostInstallValidation })
$btnExportSignedReport.Add_Click({
    try {
        $reportPath = Export-SignedOperationReport
        Set-Status 'Signed operation report exported.' 100
        [System.Windows.MessageBox]::Show("Signed operation report created:`n$reportPath", 'Export Signed Report', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
    } catch {
        Add-LogEntry "Signed report export failed: $_" 'ERROR'
        Set-Status 'Signed report export failed.' 0
        [System.Windows.MessageBox]::Show("Signed report export failed.`n$_", 'Export Signed Report', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
})

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
    else { Add-LogEntry "Working directory not found: $p - it will be created on first run." 'WARN' }
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
$btnRunInitialSetup.Add_Click({
    if (-not (Ensure-PlanApproved)) { return }
    $httpValidation = Validate-HttpConfigSettings -ShowInlineErrors $true
    if (-not $httpValidation.IsValid) {
        Add-LogEntry "HTTP.Config wizard has $($httpValidation.Errors.Count) field(s) requiring attention. Fresh Install is not blocked; existing or partial values will be used." 'WARN'
    }
    Add-LogEntry '==== INITIAL SETUP ====' 'HEADER'
    Set-Status 'Running initial setup...' 0
    $initialSetupScript = New-InitialSetupScriptBlock
    Start-NixxisJob $initialSetupScript 'Initial Setup'
})
$btnEnsureDotNet48.Add_Click({     Add-LogEntry '==== ENSURE .NET 4.8 ====' 'HEADER'; Set-Status 'Checking .NET...' 0; Start-NixxisJob $sbEnsureDotNet48 'Ensure .NET 4.8' })
$btnRunFull.Add_Click({
    if (-not (Ensure-PlanApproved)) { return }
    Add-LogEntry '==== FULL UPDATE ====' 'HEADER'
    Set-Status 'Running full update...' 0
    Start-NixxisJob $sbFullUpdate 'Full Update'
})
$btnDownload.Add_Click({    Add-LogEntry '==== DOWNLOAD ====' 'HEADER';      Set-Status 'Downloading ZIPs...' 0;    Start-NixxisJob $sbDownload      'Download'      })
$btnPrepare.Add_Click({     Add-LogEntry '==== PREPARE ====' 'HEADER';       Set-Status 'Preparing files...' 0;     Start-NixxisJob $sbPrepare       'Prepare'       })
$btnStopService.Add_Click({ if (-not (Ensure-PlanApproved)) { return }; Add-LogEntry '==== STOP SERVICE ====' 'HEADER';  Set-Status 'Stopping service...' 0;    Start-NixxisJob $sbStopService   'Stop Service'  })
$btnBackup.Add_Click({      if (-not (Ensure-PlanApproved)) { return }; Add-LogEntry '==== BACKUP ====' 'HEADER';        Set-Status 'Backing up...' 0;         Start-NixxisJob $sbBackup        'Backup'        })
$btnCleanup.Add_Click({     if (-not (Ensure-PlanApproved)) { return }; Add-LogEntry '==== CLEANUP ====' 'HEADER';       Set-Status 'Cleaning up...' 0;        Start-NixxisJob $sbCleanup       'Cleanup'       })
$btnDeploy.Add_Click({      if (-not (Ensure-PlanApproved)) { return }; Add-LogEntry '==== DEPLOY ====' 'HEADER';        Set-Status 'Deploying...' 0;          Start-NixxisJob $sbDeploy        'Deploy'        })
$btnStartService.Add_Click({Add-LogEntry '==== START SERVICE ====' 'HEADER'; Set-Status 'Starting service...' 0;   Start-NixxisJob $sbStartService  'Start Service' })
$btnInstallService.Add_Click({     if (-not (Ensure-PlanApproved)) { return }; Add-LogEntry '==== INSTALL SERVICE ====' 'HEADER'; Set-Status 'Installing service...' 0; Start-NixxisJob $sbInstallService 'Install Service' })
$btnCopyTools.Add_Click({          if (-not (Ensure-PlanApproved)) { return }; Add-LogEntry '==== COPY TOOLS ====' 'HEADER';      Set-Status 'Copying tools...' 0;      Start-NixxisJob $sbCopyTools 'Copy Tools' })
$btnInstallMoveFiles.Add_Click({   if (-not (Ensure-PlanApproved)) { return }; Add-LogEntry '==== INSTALL MOVEFILES ====' 'HEADER'; Set-Status 'Installing MoveFiles...' 0; Start-NixxisJob $sbInstallMoveFiles 'Install MoveFiles' })
$btnCreateReportingUser.Add_Click({if (-not (Ensure-PlanApproved)) { return }; Add-LogEntry '==== CREATE REPORTING USER ====' 'HEADER'; Set-Status 'Creating Reporting user...' 0; Start-NixxisJob $sbCreateReportingUser 'Create Reporting User' })
$btnConfigFirewall.Add_Click({     if (-not (Ensure-PlanApproved)) { return }; Add-LogEntry '==== CONFIG FIREWALL ====' 'HEADER'; Set-Status 'Configuring firewall...' 0; Start-NixxisJob $sbConfigFirewall 'Configure Firewall' })
$btnDeployTranscription.Add_Click({if (-not (Ensure-PlanApproved)) { return }; Add-LogEntry '==== DEPLOY TRANSCRIPTION ====' 'HEADER'; Set-Status 'Deploying transcription helpers...' 0; Start-NixxisJob $sbDeployTranscription 'Deploy Transcription Helpers' })
$btnLaunchDeployReports.Add_Click({Add-LogEntry '==== LAUNCH DEPLOY REPORTS ====' 'HEADER'; Set-Status 'Launching DeployReports...' 0; Start-NixxisJob $sbLaunchDeployReports 'Launch DeployReports' })
#endregion

#region --- Startup ---
$window.Title = "NCS Installer / Updater v$script:AppVersion"
$tbHeaderTitle.Text = 'NCS Installer / Updater'
$tbHeaderSubtitle.Text = "Explicit flows for Fresh Install and Existing Update | v$script:AppVersion"
Update-RunDirLabel
Update-OperationModeUI
Update-SourceModeUI
Update-HttpConfigUiState
Validate-HttpConfigSettings -ShowInlineErrors $true | Out-Null
Add-LogEntry "NCS Installer / Updater ready (v$script:AppVersion)." 'HEADER'
Add-LogEntry "User: $env:USERNAME  |  Host: $env:COMPUTERNAME" 'GRAY'
Add-LogEntry "Log file: $logFile" 'GRAY'
Add-LogEntry "Default working directory: $($sync.WorkDir)" 'CYAN'
Add-LogEntry "Install path: $($sync.InstallPath)" 'CYAN'
Add-LogEntry "Service name: $($sync.ServiceName)" 'CYAN'
Add-LogEntry "A dated subfolder (YYYYMMDD) will be created in that directory on each run." 'GRAY'
Add-LogEntry '' 'INFO'
Update-ServiceStatus

$window.ShowDialog() | Out-Null
#endregion
