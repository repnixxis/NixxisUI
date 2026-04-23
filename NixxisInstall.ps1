# Nixxis Initial Setup Script
# Version 1.0 - FP - Complete initial setup and configuration

# Script Parameters
param(
    [string]$ClientProvisioningUrl = "",
    [string]$ClientSoftwareUrl = "",
    [string]$NCSUrl = "",
    [switch]$OfflineMode,
    [string]$OfflinePath = "",
    [string]$LicenseKey = "",
    [switch]$SkipFirewall,
    [switch]$SkipMoveFiles,
    [switch]$SkipLicenseJobs,
    [switch]$SkipEScriptRunner,
    [switch]$Help
)

# Display help information
if ($Help) {
    Write-Host "Nixxis Initial Setup Script v1.0" -ForegroundColor Cyan
    Write-Host "=================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "DESCRIPTION:" -ForegroundColor Yellow
    Write-Host "  Performs complete initial setup of Nixxis application including:"
    Write-Host "  - Create folder structure (C:\NixxisMaintenance\Update, Backup, Install)"
    Write-Host "  - Download/copy installation files"
    Write-Host "  - Extract and deploy application"
    Write-Host "  - Configure firewall rules"
    Write-Host "  - Install MoveFiles utility"
    Write-Host "  - Configure NCS license jobs (Event 903)"
    Write-Host "  - Install eScript Runner"
    Write-Host ""
    Write-Host "USAGE:" -ForegroundColor Yellow
    Write-Host "  .\NixxisSetup.ps1 [parameters]"
    Write-Host ""
    Write-Host "PARAMETERS:" -ForegroundColor Yellow
    Write-Host "  -ClientProvisioningUrl <string>   Custom URL for ClientProvisioning.zip"
    Write-Host "  -ClientSoftwareUrl <string>       Custom URL for ClientSoftware.zip"
    Write-Host "  -NCSUrl <string>                  Custom URL for NCS.zip"
    Write-Host "  -OfflineMode                      Use ZIP files from local folder"
    Write-Host "  -OfflinePath <string>             Folder path containing ZIP files"
    Write-Host "  -LicenseKey <string>              Nixxis license key"
    Write-Host "  -SkipFirewall                     Skip firewall configuration"
    Write-Host "  -SkipMoveFiles                    Skip MoveFiles installation"
    Write-Host "  -SkipLicenseJobs                  Skip license job configuration"
    Write-Host "  -SkipEScriptRunner                Skip eScript Runner installation"
    Write-Host "  -Help                             Show this help message"
    Write-Host ""
    Write-Host "EXAMPLES:" -ForegroundColor Yellow
    Write-Host "  # Full installation (default)"
    Write-Host "  .\NixxisSetup.ps1 -LicenseKey 'YOUR-LICENSE-KEY'"
    Write-Host ""
    Write-Host "  # Offline installation"
    Write-Host "  .\NixxisSetup.ps1 -OfflineMode -OfflinePath 'C:\NixxisZips' -LicenseKey 'KEY'"
    Write-Host ""
    Write-Host "FOLDER STRUCTURE:" -ForegroundColor Yellow
    Write-Host "  C:\NixxisMaintenance\"
    Write-Host "    ├── Update\       (Future updates)"
    Write-Host "    ├── Backup\       (System backups)"
    Write-Host "    ├── Install\      (Installation files)"
    Write-Host "    │   └── YYYYMMDD\ (Today's installation)"
    Write-Host "    └── Logs\         (Log files)"
    Write-Host ""
    exit 0
}

# Set error action preference
$ErrorActionPreference = "Stop"

# Check if running as Administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Administrator)) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Please right-click on PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-Host "Administrator privileges confirmed - Proceeding with setup..." -ForegroundColor Green

# =============================================================================
# PRE-CHECK: .NET FRAMEWORK 4.8 RUNTIME
# =============================================================================

function Get-DotNetRelease {
    $regPath = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
    if (Test-Path $regPath) {
        $release = (Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).Release
        return [int]$release
    }
    return 0
}

# .NET 4.8 release key >= 528040
$dotNetRelease    = Get-DotNetRelease
$dotNet48Required = 528040

if ($dotNetRelease -lt $dotNet48Required) {
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Red
    Write-Host "  WARNING: .NET Framework 4.8 is NOT installed on this machine." -ForegroundColor Red
    if ($dotNetRelease -eq 0) {
        Write-Host "  Detected:  .NET Framework 4.x not found" -ForegroundColor Yellow
    } else {
        Write-Host "  Detected release key: $dotNetRelease  (4.8 requires >= $dotNet48Required)" -ForegroundColor Yellow
    }
    Write-Host "  Nixxis requires .NET Framework 4.8 to operate correctly." -ForegroundColor Yellow
    Write-Host "================================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "  The installer will be downloaded from Microsoft and installed." -ForegroundColor White
    Write-Host "  NOTE: A system reboot will be required after installation." -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  [I]  Install .NET 4.8 now and continue" -ForegroundColor White
    Write-Host "  [A]  Abort setup" -ForegroundColor White
    Write-Host ""
    do {
        $dotNetChoice = (Read-Host "  Enter choice (I/A)").Trim()
    } while ($dotNetChoice -notmatch '^[IiAa]$')

    if ($dotNetChoice -match '^[Aa]$') {
        Write-Host "Setup aborted by user. Please install .NET Framework 4.8 manually and re-run." -ForegroundColor Red
        exit 1
    }

    # Download and install .NET 4.8
    $dotNet48Url       = "https://go.microsoft.com/fwlink/?LinkId=2085155"
    $dotNet48Installer = Join-Path -Path $env:TEMP -ChildPath "ndp48-x86-x64-allos-enu.exe"

    Write-Host ""
    Write-Host "Downloading .NET Framework 4.8 installer..." -ForegroundColor Cyan
    try {
        $webClient48 = New-Object System.Net.WebClient
        $webClient48.DownloadFile($dotNet48Url, $dotNet48Installer)
        $webClient48.Dispose()
        Write-Host "[OK] Download complete: $dotNet48Installer" -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to download .NET 4.8 installer: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please download and install manually from: https://dotnet.microsoft.com/download/dotnet-framework/net48" -ForegroundColor Yellow
        exit 1
    }

    Write-Host "Installing .NET Framework 4.8 (this may take several minutes)..." -ForegroundColor Cyan
    try {
        # /quiet = silent, /norestart = defer reboot so we can prompt the user
        $process = Start-Process -FilePath $dotNet48Installer `
                                 -ArgumentList "/quiet /norestart /log `"$env:TEMP\dotnet48_install.log`"" `
                                 -Wait -PassThru
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
            Write-Host "[OK] .NET Framework 4.8 installed successfully (exit code: $($process.ExitCode))" -ForegroundColor Green
        } else {
            throw "Installer exited with code $($process.ExitCode). Check log: $env:TEMP\dotnet48_install.log"
        }
    }
    catch {
        Write-Host "ERROR: .NET 4.8 installation failed: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
    finally {
        Remove-Item -Path $dotNet48Installer -Force -ErrorAction SilentlyContinue
    }

    # Verify installation succeeded
    $dotNetRelease = Get-DotNetRelease
    if ($dotNetRelease -ge $dotNet48Required) {
        Write-Host ".NET Framework 4.8 is now installed. A reboot is recommended." -ForegroundColor Green
        Write-Host ""
        Write-Host "  [C]  Continue setup without rebooting (not recommended)" -ForegroundColor White
        Write-Host "  [R]  Reboot now (setup must be re-run after reboot)" -ForegroundColor White
        Write-Host ""
        do {
            $rebootChoice = (Read-Host "  Enter choice (C/R)").Trim()
        } while ($rebootChoice -notmatch '^[CcRr]$')

        if ($rebootChoice -match '^[Rr]$') {
            Write-Host "Rebooting in 15 seconds. Re-run this script after restart." -ForegroundColor Yellow
            Start-Sleep -Seconds 15
            Restart-Computer -Force
            exit 0
        }
        Write-Host "Continuing without reboot..." -ForegroundColor Yellow
    } else {
        Write-Host "ERROR: .NET 4.8 installation could not be verified. Please reboot and try again." -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host ".NET Framework 4.8 verified (release key: $dotNetRelease)" -ForegroundColor Green
}

# Load .NET compression assemblies required for robust ZIP extraction
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# =============================================================================
# PHASE 0: CREATE FOLDER STRUCTURE
# =============================================================================

Write-Host "`n=== PHASE 0: CREATING FOLDER STRUCTURE ===" -ForegroundColor Cyan

# Define base paths
$basePath = "C:\NixxisMaintenance"
$updatePath = Join-Path -Path $basePath -ChildPath "Update"
$backupPath = Join-Path -Path $basePath -ChildPath "Backup"
$installPath = Join-Path -Path $basePath -ChildPath "Install"
$logsPath = Join-Path -Path $basePath -ChildPath "Logs"

# Create today's installation folder
$dateFolder = Get-Date -Format "yyyyMMdd"
$todayInstallPath = Join-Path -Path $installPath -ChildPath $dateFolder

# Create all required directories
$foldersToCreate = @(
    @{Path = $basePath; Description = "Base maintenance folder"},
    @{Path = $updatePath; Description = "Updates folder"},
    @{Path = $backupPath; Description = "Backups folder"},
    @{Path = $installPath; Description = "Installation folder"},
    @{Path = $todayInstallPath; Description = "Today's installation folder ($dateFolder)"},
    @{Path = $logsPath; Description = "Logs folder"}
)

foreach ($folder in $foldersToCreate) {
    if (-not (Test-Path $folder.Path)) {
        Write-Host "Creating $($folder.Description): $($folder.Path)" -ForegroundColor Yellow
        New-Item -Path $folder.Path -ItemType Directory -Force | Out-Null
        Write-Host "[OK] Created: $($folder.Path)" -ForegroundColor Green
    } else {
        Write-Host "[EXISTS] $($folder.Description): $($folder.Path)" -ForegroundColor Gray
    }
}

Write-Host "`nFolder structure created successfully!" -ForegroundColor Green
Write-Host "Installation files will be stored in: $todayInstallPath" -ForegroundColor White

# Initialize logging (now that Logs folder exists)
$logDate = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path -Path $logsPath -ChildPath "NixxisSetup_$logDate.log"

# Function to write to both log and console
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [string]$Color = "White"
    )
    
    $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timeStamp - $Message"
    
    Add-Content -Path $logFile -Value $logMessage
    Write-Host $logMessage -ForegroundColor $Color
}

Write-Log "=== NIXXIS INITIAL SETUP SCRIPT ===" -Color "Cyan"
Write-Log "Starting Nixxis setup process..." -Color "Cyan"
Write-Log "Installation directory: $todayInstallPath" -Color "White"

# Prompt user for Nixxis installation path
Write-Host ""
Write-Host "=== NIXXIS INSTALLATION PATH ==" -ForegroundColor Cyan
Write-Host "  Where should Nixxis be installed?" -ForegroundColor White
Write-Host "  Press ENTER to accept the default [C:\Nixxis]" -ForegroundColor Gray
Write-Host ""
$nixxisInstallInput = (Read-Host "  Install path").Trim()
if ([string]::IsNullOrEmpty($nixxisInstallInput)) {
    $nixxisInstallPath = "C:\Nixxis"
} else {
    $nixxisInstallPath = $nixxisInstallInput.TrimEnd('\')
}
Write-Log "Nixxis will be installed to: $nixxisInstallPath" -Color "Cyan"

# Function to get latest folder from web directory
function Get-LatestWebFolder {
    param(
        [Parameter(Mandatory=$true)]
        [string]$BaseUrl,
        [string]$Description = "folder"
    )
    
    try {
        Write-Log "Fetching directory listing from: $BaseUrl" -Color "Yellow"
        $response = Invoke-WebRequest -Uri $BaseUrl -UseBasicParsing -TimeoutSec 30
        
        $folders = @()
        
        # Try multiple parsing patterns for different web server formats
        
        # Pattern 1: Standard href links (href="folder/")
        $folderPattern1 = 'href="([^"]+/)"'
        $matches1 = [regex]::Matches($response.Content, $folderPattern1)
        
        foreach ($match in $matches1) {
            $folderName = $match.Groups[1].Value.TrimEnd('/')
            if ($folderName -ne ".." -and $folderName -ne "." -and $folderName -ne "icons" -and $folderName -ne "cgi-bin") {
                $folders += $folderName
                Write-Log "  Found $description (pattern 1): $folderName" -Color "Gray"
            }
        }
        
        # Pattern 2: Directory listing format like "v3.2" where &lt;dir&gt; indicates directory
        $folderPattern2 = '(\w+\.\d+|\w+\d+\.\d+\.\d+)\s+&lt;dir&gt;'
        $matches2 = [regex]::Matches($response.Content, $folderPattern2)
        
        foreach ($match in $matches2) {
            $folderName = $match.Groups[1].Value
            if ($folderName -notin $folders) {
                $folders += $folderName
                Write-Log "  Found $description (pattern 2): $folderName" -Color "Gray"
            }
        }
        
        # Pattern 3: Look for version patterns in text (v3.2, v3.1, etc.)
        $folderPattern3 = '(v\d+\.\d+)'
        $matches3 = [regex]::Matches($response.Content, $folderPattern3)
        
        foreach ($match in $matches3) {
            $folderName = $match.Groups[1].Value
            if ($folderName -notin $folders) {
                $folders += $folderName
                Write-Log "  Found $description (pattern 3): $folderName" -Color "Gray"
            }
        }
        
        # Pattern 4: Look for build numbers (3.2.1055, 3.2.1056, etc.)
        $folderPattern4 = '(\d+\.\d+\.\d+)'
        $matches4 = [regex]::Matches($response.Content, $folderPattern4)
        
        foreach ($match in $matches4) {
            $folderName = $match.Groups[1].Value
            if ($folderName -notin $folders) {
                $folders += $folderName
                Write-Log "  Found $description (pattern 4): $folderName" -Color "Gray"
            }
        }
        
        if ($folders.Count -eq 0) {
            Write-Log "Debug: Response content preview:" -Color "Red"
            Write-Log $response.Content.Substring(0, [Math]::Min(500, $response.Content.Length)) -Color "Red"
            throw "No valid folders found at $BaseUrl"
        }
        
        # Sort folders and get the latest
        $sortedFolders = $folders | Sort-Object { 
            $version = $_ -replace '[^\d\.]', ''
            if ($version -and $version -match '^\d+\.\d+') {
                try {
                    [version]$version
                } catch {
                    [version]"0.0"
                }
            } else {
                [version]"0.0"
            }
        }
        
        $latestFolder = $sortedFolders | Select-Object -Last 1
        Write-Log "Latest $description found: $latestFolder" -Color "Green"
        
        return $latestFolder
    }
    catch {
        Write-Log "Error getting latest $description from $BaseUrl`: $($_.Exception.Message)" -Color "Red"
        throw
    }
}

# Function to download file with progress
function Download-FileWithProgress {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Url,
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )
    
    try {
        Write-Log "Downloading from: $Url" -Color "Yellow"
        Write-Log "Saving to: $OutputPath" -Color "Yellow"
        
        # Create directory if it doesn't exist
        $outputDir = Split-Path -Parent $OutputPath
        if (-not (Test-Path $outputDir)) {
            New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
        }
        
        # Download with progress
        $webClient = New-Object System.Net.WebClient
        
        # Add progress tracking
        Register-ObjectEvent -InputObject $webClient -EventName DownloadProgressChanged -Action {
            $Global:DownloadProgress = $Event.SourceEventArgs.ProgressPercentage
            Write-Progress -Activity "Downloading file" -Status "Progress: $($Event.SourceEventArgs.ProgressPercentage)%" -PercentComplete $Event.SourceEventArgs.ProgressPercentage
        } | Out-Null
        
        Register-ObjectEvent -InputObject $webClient -EventName DownloadFileCompleted -Action {
            Write-Progress -Activity "Downloading file" -Completed
            $Global:DownloadComplete = $true
        } | Out-Null
        
        $Global:DownloadComplete = $false
        $webClient.DownloadFileAsync($Url, $OutputPath)
        
        # Wait for download to complete
        while (-not $Global:DownloadComplete) {
            Start-Sleep -Milliseconds 100
        }
        
        $webClient.Dispose()
        Get-Event | Remove-Event  # Clean up events
        
        if (Test-Path $OutputPath) {
            $fileSize = (Get-Item $OutputPath).Length
            Write-Log "Download completed successfully. File size: $([math]::Round($fileSize/1MB, 2)) MB" -Color "Green"
        } else {
            throw "File was not downloaded successfully"
        }
    }
    catch {
        Write-Log "Error downloading file: $($_.Exception.Message)" -Color "Red"
        throw
    }
}

# Function to verify offline ZIP files
function Test-OfflineZipFiles {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderPath
    )
    
    try {
        Write-Log "Verifying offline ZIP files in folder: $FolderPath" -Color "Yellow"
        
        if (-not (Test-Path $FolderPath)) {
            throw "Offline folder does not exist: $FolderPath"
        }
        
        $requiredZips = @("ClientProvisioning.zip", "ClientSoftware.zip", "NCS.zip")
        $missingFiles = @()
        
        foreach ($zip in $requiredZips) {
            $zipPath = Join-Path -Path $FolderPath -ChildPath $zip
            if (Test-Path $zipPath) {
                $fileSize = (Get-Item $zipPath).Length
                Write-Log "[OK] Found $zip - Size: $([math]::Round($fileSize/1MB, 2)) MB" -Color "Green"
            } else {
                $missingFiles += $zip
                Write-Log "[MISSING] $zip not found in offline folder" -Color "Red"
            }
        }
        
        if ($missingFiles.Count -gt 0) {
            throw "Missing required ZIP files: $($missingFiles -join ', ')"
        }
        
        Write-Log "All required ZIP files found in offline folder" -Color "Green"
        return $true
    }
    catch {
        Write-Log "Offline ZIP verification failed: $($_.Exception.Message)" -Color "Red"
        throw
    }
}

# Function to copy offline ZIP files
function Copy-OfflineZipFiles {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SourcePath,
        [Parameter(Mandatory=$true)]
        [string]$DestinationPath
    )
    
    try {
        Write-Log "Copying offline ZIP files from $SourcePath to $DestinationPath" -Color "Yellow"
        
        $requiredZips = @("ClientProvisioning.zip", "ClientSoftware.zip", "NCS.zip")
        
        foreach ($zip in $requiredZips) {
            $sourcePath = Join-Path -Path $SourcePath -ChildPath $zip
            $destPath = Join-Path -Path $DestinationPath -ChildPath $zip
            
            Write-Log "Copying $zip..." -Color "Yellow"
            Copy-Item -Path $sourcePath -Destination $destPath -Force
            
            # Verify copy
            if (Test-Path $destPath) {
                $fileSize = (Get-Item $destPath).Length
                Write-Log "[OK] Copied $zip successfully - Size: $([math]::Round($fileSize/1MB, 2)) MB" -Color "Green"
            } else {
                throw "Failed to copy $zip to destination"
            }
        }
        
        Write-Log "All offline ZIP files copied successfully" -Color "Green"
        return $true
    }
    catch {
        Write-Log "Error copying offline ZIP files: $($_.Exception.Message)" -Color "Red"
        throw
    }
}

# Function to fetch ZIP files from web
function Get-NixxisZipFiles {
    param(
        [Parameter(Mandatory=$true)]
        [string]$DownloadDirectory,
        [string]$CustomClientProvisioningUrl = "",
        [string]$CustomClientSoftwareUrl = "",
        [string]$CustomNCSUrl = ""
    )
    
    try {
        Write-Log "Starting web download process..." -Color "Cyan"
        
        # Check if custom URLs are provided
        $useCustomUrls = $CustomClientProvisioningUrl -or $CustomClientSoftwareUrl -or $CustomNCSUrl
        
        if ($useCustomUrls) {
            Write-Log "Custom URLs detected - Using provided URLs where specified" -Color "Magenta"
        }
        
        # Initialize variables for final URLs
        $clientProvisioningUrl = ""
        $clientSoftwareUrl = ""
        $ncsUrl = ""
        
        # Auto-discovery logic (only if custom URLs are not fully provided)
        if (-not $CustomClientProvisioningUrl -or -not $CustomClientSoftwareUrl -or -not $CustomNCSUrl) {
            Write-Log "Auto-discovering URLs for missing parameters..." -Color "Yellow"
            
            # Base URL configuration
            $baseUrl = "http://update.nixxis.net"
            
            # Auto-detect the latest version folder
            Write-Log "Auto-detecting latest version folder..." -Color "Yellow"
            $latestVersion = Get-LatestWebFolder -BaseUrl $baseUrl -Description "version"
            Write-Log "Detected latest version: $latestVersion" -Color "Green"
            
            # Construct version URL
            $versionUrl = "$baseUrl/$latestVersion"
            Write-Log "Using version URL: $versionUrl" -Color "White"
            
            # Auto-discover client URLs if not provided
            if (-not $CustomClientProvisioningUrl -or -not $CustomClientSoftwareUrl) {
                $clientUrl = "$versionUrl/Client"
                $latestClientFolder = Get-LatestWebFolder -BaseUrl $clientUrl -Description "client version"
                
                if (-not $CustomClientProvisioningUrl) {
                    $clientProvisioningUrl = "$clientUrl/$latestClientFolder/ClientProvisioning.zip"
                }
                
                if (-not $CustomClientSoftwareUrl) {
                    $clientSoftwareUrl = "$clientUrl/$latestClientFolder/ClientSoftware.zip"
                }
            }
            
            # Auto-discover NCS URL if not provided
            if (-not $CustomNCSUrl) {
                $serverUrl = "$versionUrl/Server"
                $latestServerFolder = Get-LatestWebFolder -BaseUrl $serverUrl -Description "server version"
                $ncsUrl = "$serverUrl/$latestServerFolder/NCS.zip"
            }
        }
        
        # Use custom URLs where provided
        if ($CustomClientProvisioningUrl) {
            $clientProvisioningUrl = $CustomClientProvisioningUrl
            Write-Log "Using custom ClientProvisioning URL: $clientProvisioningUrl" -Color "Magenta"
        }
        
        if ($CustomClientSoftwareUrl) {
            $clientSoftwareUrl = $CustomClientSoftwareUrl
            Write-Log "Using custom ClientSoftware URL: $clientSoftwareUrl" -Color "Magenta"
        }
        
        if ($CustomNCSUrl) {
            $ncsUrl = $CustomNCSUrl
            Write-Log "Using custom NCS URL: $ncsUrl" -Color "Magenta"
        }
        
        # Define local file paths in today's installation folder
        $clientProvisioningPath = Join-Path -Path $DownloadDirectory -ChildPath "ClientProvisioning.zip"
        $clientSoftwarePath = Join-Path -Path $DownloadDirectory -ChildPath "ClientSoftware.zip"
        $ncsPath = Join-Path -Path $DownloadDirectory -ChildPath "NCS.zip"
        
        # Display final URLs
        Write-Log "=== FINAL DOWNLOAD URLS ===" -Color "Magenta"
        Write-Log "ClientProvisioning.zip: $clientProvisioningUrl" -Color "White"
        Write-Log "ClientSoftware.zip: $clientSoftwareUrl" -Color "White"
        Write-Log "NCS.zip: $ncsUrl" -Color "White"
        Write-Log "============================" -Color "Magenta"
        
        # Download files
        Write-Log "Downloading ClientProvisioning.zip..." -Color "Cyan"
        Download-FileWithProgress -Url $clientProvisioningUrl -OutputPath $clientProvisioningPath
        
        Write-Log "Downloading ClientSoftware.zip..." -Color "Cyan"
        Download-FileWithProgress -Url $clientSoftwareUrl -OutputPath $clientSoftwarePath
        
        Write-Log "Downloading NCS.zip..." -Color "Cyan"
        Download-FileWithProgress -Url $ncsUrl -OutputPath $ncsPath
        
        Write-Log "All ZIP files downloaded successfully!" -Color "Green"
        
        # Verify all files exist
        $requiredZips = @("NCS.zip", "ClientProvisioning.zip", "ClientSoftware.zip")
        foreach ($zip in $requiredZips) {
            $zipPath = Join-Path -Path $DownloadDirectory -ChildPath $zip
            if (-not (Test-Path $zipPath)) {
                throw "Downloaded file verification failed: $zip not found"
            }
            $fileSize = (Get-Item $zipPath).Length
            Write-Log "Verified $zip - Size: $([math]::Round($fileSize/1MB, 2)) MB" -Color "Green"
        }
    }
    catch {
        Write-Log "Error during web download process: $($_.Exception.Message)" -Color "Red"
        throw
    }
}

# Function to extract ZIP files robustly, bypassing Expand-Archive's Central Directory issues
function Expand-ZipRobust {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ZipPath,
        [Parameter(Mandatory=$true)]
        [string]$DestinationPath
    )

    try {
        Write-Log "Extracting ZIP (robust stream method): $(Split-Path -Leaf $ZipPath)" -Color "Yellow"

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

                    # Directory entry
                    if ($entry.FullName.EndsWith('/') -or $entry.FullName.EndsWith('\')) {
                        New-Item -Path $entryDestPath -ItemType Directory -Force | Out-Null
                        continue
                    }

                    # Ensure parent directory exists
                    $parentDir = Split-Path -Parent $entryDestPath
                    if (-not (Test-Path $parentDir)) {
                        New-Item -Path $parentDir -ItemType Directory -Force | Out-Null
                    }

                    # Stream-copy the entry to disk
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
                Write-Log "Extraction completed: $(Split-Path -Leaf $ZipPath)" -Color "Green"
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
        Write-Log "Robust extraction failed ($($_.Exception.Message)) - Trying Expand-Archive fallback..." -Color "Yellow"
        try {
            Expand-Archive -Path $ZipPath -DestinationPath $DestinationPath -Force
            Write-Log "Extraction completed via Expand-Archive fallback: $(Split-Path -Leaf $ZipPath)" -Color "Green"
        }
        catch {
            Write-Log "All extraction methods failed for: $ZipPath" -Color "Red"
            throw
        }
    }
}

# =============================================================================
# PHASE 1: DOWNLOAD/COPY INSTALLATION FILES
# =============================================================================

Write-Log "`n=== PHASE 1: OBTAINING INSTALLATION FILES ===" -Color "Cyan"

# Check if any ZIP files already exist in today's installation folder
$zipNames = @("NCS.zip", "ClientProvisioning.zip", "ClientSoftware.zip")
$existingZips = $zipNames | Where-Object { Test-Path (Join-Path -Path $todayInstallPath -ChildPath $_) }

$skipAcquisition = $false
if ($existingZips.Count -gt 0) {
    Write-Log "The following ZIP file(s) already exist in: $todayInstallPath" -Color "Yellow"
    foreach ($z in $existingZips) {
        $zPath = Join-Path -Path $todayInstallPath -ChildPath $z
        $zSize = [math]::Round((Get-Item $zPath).Length / 1MB, 2)
        Write-Log "  [EXISTS] $z  ($zSize MB)" -Color "Yellow"
    }
    Write-Host ""
    Write-Host "  Choose an option:" -ForegroundColor Cyan
    Write-Host "  [R]  Re-download / overwrite existing files" -ForegroundColor White
    Write-Host "  [S]  Skip download and use existing files" -ForegroundColor White
    Write-Host ""
    do {
        $choice = (Read-Host "  Enter choice (R/S)").Trim()
    } while ($choice -notmatch '^[RrSs]$')

    if ($choice -match '^[Ss]$') {
        Write-Log "User chose to use existing ZIP files - Skipping acquisition phase" -Color "Cyan"
        $skipAcquisition = $true
    } else {
        Write-Log "User chose to re-download/overwrite existing files" -Color "Cyan"
    }
}

if (-not $skipAcquisition) {
    if ($OfflineMode) {
        Write-Log "OFFLINE MODE ACTIVATED" -Color "Magenta"
        
        # Determine offline folder path
        $offlineFolder = if ($OfflinePath) {
            Write-Log "Using specified offline path: $OfflinePath" -Color "White"
            $OfflinePath
        } else {
            Write-Log "No offline path specified - Please provide -OfflinePath parameter" -Color "Red"
            exit 1
        }
        
        try {
            # Verify offline ZIP files exist
            Test-OfflineZipFiles -FolderPath $offlineFolder
            
            # Copy files to today's installation folder
            Write-Log "Copying ZIP files to installation directory..." -Color "Yellow"
            Copy-OfflineZipFiles -SourcePath $offlineFolder -DestinationPath $todayInstallPath
            
            Write-Log "Offline mode preparation completed successfully!" -Color "Green"
        }
        catch {
            Write-Log "Error during offline mode preparation: $($_.Exception.Message)" -Color "Red"
            Write-Log "Stack Trace: $($_.Exception.StackTrace)" -Color "Red"
            exit 1
        }
    } else {
        # Online mode - Download files to today's installation folder
        Write-Log "ONLINE MODE - DOWNLOADING FILES" -Color "Cyan"
        try {
            Get-NixxisZipFiles -DownloadDirectory $todayInstallPath `
                               -CustomClientProvisioningUrl $ClientProvisioningUrl `
                               -CustomClientSoftwareUrl $ClientSoftwareUrl `
                               -CustomNCSUrl $NCSUrl
            
            Write-Log "Web download phase completed successfully!" -Color "Green"
        }
        catch {
            Write-Log "Error during web download phase: $($_.Exception.Message)" -Color "Red"
            Write-Log "Stack Trace: $($_.Exception.StackTrace)" -Color "Red"
            exit 1
        }
    }
}

# =============================================================================
# PHASE 2: EXTRACT AND DEPLOY APPLICATION
# =============================================================================

Write-Log "`n=== PHASE 2: EXTRACTING AND DEPLOYING APPLICATION ===" -Color "Cyan"

try {
    # Verify zip files exist in today's installation folder
    $requiredZips = @("NCS.zip", "ClientProvisioning.zip", "ClientSoftware.zip")
    foreach ($zip in $requiredZips) {
        $zipPath = Join-Path -Path $todayInstallPath -ChildPath $zip
        if (-not (Test-Path $zipPath)) {
            throw "Required zip file not found: $zip"
        }
    }

    # Create/Clean NixxisApplicationServer folder in today's install directory
    $nixxisAppServer = Join-Path -Path $todayInstallPath -ChildPath "NixxisApplicationServer"
    if (Test-Path $nixxisAppServer) {
        Write-Log "Removing existing NixxisApplicationServer folder..." -Color "Yellow"
        Remove-Item -Path $nixxisAppServer -Recurse -Force
    }
    New-Item -Path $nixxisAppServer -ItemType Directory -Force | Out-Null
    Write-Log "Created NixxisApplicationServer folder" -Color "Green"

    # Extract NCS.zip
    Write-Log "Extracting NCS.zip..." -Color "Yellow"
    Expand-ZipRobust -ZipPath (Join-Path -Path $todayInstallPath -ChildPath "NCS.zip") -DestinationPath $nixxisAppServer
    Write-Log "Extracted NCS.zip successfully" -Color "Green"

    # Create and extract ClientSoftware
    $clientSoftwarePath = Join-Path -Path $nixxisAppServer -ChildPath "ClientSoftware"
    New-Item -Path $clientSoftwarePath -ItemType Directory -Force | Out-Null
    Write-Log "Created ClientSoftware folder" -Color "Green"

    Write-Log "Extracting ClientSoftware.zip..." -Color "Yellow"
    Expand-ZipRobust -ZipPath (Join-Path -Path $todayInstallPath -ChildPath "ClientSoftware.zip") -DestinationPath $clientSoftwarePath
    Write-Log "Extracted ClientSoftware.zip successfully" -Color "Green"

    # Create provisioning folder
    $provisioningPath = Join-Path -Path $nixxisAppServer -ChildPath "CrAppServer\provisioning"
    New-Item -Path $provisioningPath -ItemType Directory -Force | Out-Null
    Write-Log "Created provisioning folder" -Color "Green"

    # Copy ClientSoftware.zip to provisioning
    Copy-Item -Path (Join-Path -Path $todayInstallPath -ChildPath "ClientSoftware.zip") -Destination $provisioningPath -Force
    Write-Log "Copied ClientSoftware.zip to provisioning folder" -Color "Green"

    # Extract ClientProvisioning.zip
    $provisioningClientPath = Join-Path -Path $provisioningPath -ChildPath "client"
    New-Item -Path $provisioningClientPath -ItemType Directory -Force | Out-Null
    Write-Log "Extracting ClientProvisioning.zip..." -Color "Yellow"
    Expand-ZipRobust -ZipPath (Join-Path -Path $todayInstallPath -ChildPath "ClientProvisioning.zip") -DestinationPath $provisioningClientPath
    Write-Log "Extracted ClientProvisioning.zip successfully" -Color "Green"

    # Move settings folder
    $settingsSource = Join-Path -Path $provisioningClientPath -ChildPath "settings"
    if (Test-Path $settingsSource) {
        Write-Log "Moving settings folder..." -Color "Yellow"
        Move-Item -Path $settingsSource -Destination $provisioningPath -Force
        Write-Log "Moved settings folder successfully" -Color "Green"
    }
    else {
        Write-Log "Settings folder not found in ClientProvisioning extract - Skipping (this is optional)" -Color "Yellow"
    }

    # Now deploy to C:\Nixxis
    Write-Log "Starting deployment to $nixxisInstallPath..." -Color "Cyan"
    
    $sourceBaseDir = $nixxisAppServer
    
    # Define folders to copy
    $foldersToCopy = @(
        @{Source="ClientSoftware"; Destination=(Join-Path $nixxisInstallPath "ClientSoftware")},
        @{Source="CrAppServer";    Destination=(Join-Path $nixxisInstallPath "CrAppServer")},
        @{Source="MediaServer";    Destination=(Join-Path $nixxisInstallPath "MediaServer")},
        @{Source="Reporting";      Destination=(Join-Path $nixxisInstallPath "Reporting")}
        # SampleConfigFiles is handled separately below (rename + NCC* exclusion)
    )

    # Create Nixxis install root if it doesn't exist
    if (-not (Test-Path $nixxisInstallPath)) {
        Write-Log "Creating Nixxis directory at $nixxisInstallPath" -Color "Yellow"
        New-Item -Path $nixxisInstallPath -ItemType Directory -Force | Out-Null
    }

    # Copy each folder
    foreach ($folder in $foldersToCopy) {
        $sourcePath = Join-Path -Path $sourceBaseDir -ChildPath $folder.Source
        $destinationPath = $folder.Destination

        if (Test-Path $sourcePath) {
            Write-Log "Copying $($folder.Source) to $($folder.Destination)..." -Color "Yellow"
            try {
                # Create destination directory if it doesn't exist
                if (-not (Test-Path $destinationPath)) {
                    New-Item -Path $destinationPath -ItemType Directory -Force | Out-Null
                    Write-Log "Created destination directory: $destinationPath" -Color "Green"
                }

                # Copy with overwrite
                Copy-Item -Path "$sourcePath\*" -Destination $destinationPath -Recurse -Force
                Write-Log "Successfully copied $($folder.Source) to $($folder.Destination)" -Color "Green"
            }
            catch {
                Write-Log "Error copying $($folder.Source): $($_.Exception.Message)" -Color "Red"
                throw
            }
        }
        else {
            Write-Log "Source folder not found: $sourcePath - Skipping" -Color "Yellow"
        }
    }

    Write-Log "Deployment process completed successfully" -Color "Green"

    # -------------------------------------------------------------------------
    # Deploy SampleConfigFiles -> CrAppServer  (skip NCC*, strip .sample)
    # -------------------------------------------------------------------------
    $sampleSource      = Join-Path -Path $sourceBaseDir -ChildPath "SampleConfigFiles"
    $sampleDestination = Join-Path -Path $nixxisInstallPath -ChildPath "CrAppServer"

    if (Test-Path $sampleSource) {
        Write-Log "Processing SampleConfigFiles -> $sampleDestination" -Color "Cyan"

        $sampleFiles = Get-ChildItem -Path $sampleSource -File -Recurse
        $copied  = 0
        $skipped = 0

        foreach ($file in $sampleFiles) {
            # Skip files whose name starts with NCC
            if ($file.Name -like 'NCC*') {
                Write-Log "  [SKIP] $($file.Name)  (NCC* exclusion rule)" -Color "Gray"
                $skipped++
                continue
            }

            # Compute relative path inside SampleConfigFiles to preserve sub-folders
            $relativePath = $file.FullName.Substring($sampleSource.Length).TrimStart('\', '/')

            # Strip .sample extension if present
            $destRelative = if ($relativePath -match '\.sample$') {
                $relativePath -replace '\.sample$', ''
            } else {
                $relativePath
            }

            $destFile = Join-Path -Path $sampleDestination -ChildPath $destRelative
            $destDir  = Split-Path -Parent $destFile

            if (-not (Test-Path $destDir)) {
                New-Item -Path $destDir -ItemType Directory -Force | Out-Null
            }

            Copy-Item -Path $file.FullName -Destination $destFile -Force
            Write-Log "  [OK] $($file.Name) -> $(Split-Path -Leaf $destFile)" -Color "Green"
            $copied++
        }

        Write-Log "SampleConfigFiles processing complete: $copied file(s) copied, $skipped skipped" -Color "Green"
    }
    else {
        Write-Log "SampleConfigFiles folder not found in source - Skipping" -Color "Yellow"
    }
}
catch {
    Write-Log "Error during extraction/deployment phase: $($_.Exception.Message)" -Color "Red"
    Write-Log "Stack Trace: $($_.Exception.StackTrace)" -Color "Red"
    exit 1
}

# =============================================================================
# PHASE 3: CREATE LOGS FOLDER + INSTALL CRAPPSERVER SERVICE
# =============================================================================

Write-Log "`n=== PHASE 3: SERVICE INSTALLATION ===" -Color "Cyan"

try {
    # Create Logs folder inside Nixxis install path
    $nixxisLogsPath = Join-Path -Path $nixxisInstallPath -ChildPath "Logs"
    if (-not (Test-Path $nixxisLogsPath)) {
        New-Item -Path $nixxisLogsPath -ItemType Directory -Force | Out-Null
        Write-Log "Created Logs folder: $nixxisLogsPath" -Color "Green"
    } else {
        Write-Log "Logs folder already exists: $nixxisLogsPath" -Color "Gray"
    }

    # Install CrAppServer Windows service
    $crAppServerDir = Join-Path -Path $nixxisInstallPath -ChildPath "CrAppServer"
    $crAppServerExe = Join-Path -Path $crAppServerDir -ChildPath "CrAppServer.exe"

    if (Test-Path $crAppServerExe) {
        Write-Log "Installing CrAppServer service..." -Color "Cyan"
        Push-Location -Path $crAppServerDir
        try {
            $svcOutput = & $crAppServerExe -install 2>&1
            foreach ($line in $svcOutput) {
                Write-Log "  [CrAppServer] $line" -Color "White"
            }
            Write-Log "CrAppServer service installation completed" -Color "Green"
        }
        finally {
            Pop-Location
        }
    } else {
        Write-Log "WARNING: CrAppServer.exe not found at $crAppServerExe - Skipping service install" -Color "Yellow"
    }
}
catch {
    Write-Log "Error during service installation: $($_.Exception.Message)" -Color "Red"
    Write-Log "Stack Trace: $($_.Exception.StackTrace)" -Color "Red"
    exit 1
}

# =============================================================================
# PHASE 4: COPY TOOLS FOLDER
# =============================================================================

Write-Log "`n=== PHASE 4: COPYING TOOLS FOLDER ===" -Color "Cyan"

try {
    $toolsSource      = Join-Path -Path $nixxisAppServer -ChildPath "Tools"
    $toolsDestination = Join-Path -Path $nixxisInstallPath -ChildPath "Tools"

    if (Test-Path $toolsSource) {
        Write-Log "Copying Tools: $toolsSource -> $toolsDestination" -Color "Yellow"
        if (-not (Test-Path $toolsDestination)) {
            New-Item -Path $toolsDestination -ItemType Directory -Force | Out-Null
        }
        Copy-Item -Path "$toolsSource\*" -Destination $toolsDestination -Recurse -Force
        Write-Log "Tools folder copied successfully" -Color "Green"
    } else {
        Write-Log "Tools folder not found at $toolsSource - Skipping" -Color "Yellow"
    }
}
catch {
    Write-Log "Error copying Tools folder: $($_.Exception.Message)" -Color "Red"
    Write-Log "Stack Trace: $($_.Exception.StackTrace)" -Color "Red"
    exit 1
}

# =============================================================================
# PHASE 5: INSTALL MOVEFILES SERVICE
# =============================================================================

Write-Log "`n=== PHASE 5: INSTALLING MOVEFILES SERVICE ===" -Color "Cyan"

try {
    # Derive drive letter from the Nixxis install path (e.g. C:)
    $installDrive    = Split-Path -Qualifier $nixxisInstallPath
    $moveFilesExe    = Join-Path -Path $nixxisInstallPath -ChildPath "Tools\MoveFiles\MoveFiles.exe"
    $installUtilExe  = "$installDrive\Windows\Microsoft.NET\Framework64\v4.0.30319\installutil.exe"

    if (-not (Test-Path $installUtilExe)) {
        Write-Log "WARNING: installutil.exe not found at $installUtilExe" -Color "Yellow"
    }

    if (Test-Path $moveFilesExe) {
        Write-Log "Installing MoveFiles service..." -Color "Cyan"
        Write-Log "  installutil : $installUtilExe" -Color "Gray"
        Write-Log "  MoveFiles   : $moveFilesExe" -Color "Gray"

        Push-Location -Path "$installDrive\Windows\Microsoft.NET\Framework64\v4.0.30319"
        try {
            $mfOutput = & $installUtilExe $moveFilesExe 2>&1
            foreach ($line in $mfOutput) {
                Write-Log "  [installutil] $line" -Color "White"
            }
            Write-Log "MoveFiles service installation completed" -Color "Green"
        }
        finally {
            Pop-Location
        }

        # Rename SampleMoveFiles.xml -> MoveFiles.xml
        $sampleXml = Join-Path -Path $nixxisInstallPath -ChildPath "Tools\MoveFiles\SampleMoveFiles.xml"
        $targetXml = Join-Path -Path $nixxisInstallPath -ChildPath "Tools\MoveFiles\MoveFiles.xml"

        if (Test-Path $sampleXml) {
            if (Test-Path $targetXml) {
                # Target already exists — Rename-Item would fail; just remove the sample file
                Remove-Item -Path $sampleXml -Force
                Write-Log "MoveFiles.xml already exists - Removed SampleMoveFiles.xml" -Color "Gray"
            } else {
                Rename-Item -Path $sampleXml -NewName "MoveFiles.xml" -Force
                Write-Log "Renamed SampleMoveFiles.xml -> MoveFiles.xml" -Color "Green"
            }
        } elseif (Test-Path $targetXml) {
            Write-Log "MoveFiles.xml already exists - Skipping rename" -Color "Gray"
        } else {
            Write-Log "WARNING: SampleMoveFiles.xml not found at $sampleXml" -Color "Yellow"
        }

        Write-Host ""
        Write-Host "  *** ACTION REQUIRED - MoveFiles ***" -ForegroundColor Yellow
        Write-Host "  Please edit the MoveFiles configuration file:" -ForegroundColor White
        Write-Host "  $targetXml" -ForegroundColor Cyan
        Write-Host "  Then start the MoveFiles service manually via services.msc." -ForegroundColor White
        Write-Host ""
    } else {
        Write-Log "WARNING: MoveFiles.exe not found at $moveFilesExe - Skipping" -Color "Yellow"
    }
}
catch {
    Write-Log "Error during MoveFiles installation: $($_.Exception.Message)" -Color "Red"
    Write-Log "Stack Trace: $($_.Exception.StackTrace)" -Color "Red"
    exit 1
}

# =============================================================================
# PHASE 6: CREATE REPORTING LOCAL USER
# =============================================================================

Write-Log "`n=== PHASE 6: REPORTING USER SETUP ===" -Color "Cyan"

Write-Host ""
Write-Host "  A local user account 'Reporting' can be created for SQL Reporting Services." -ForegroundColor White
Write-Host "  Default credentials: Username=Reporting  Password=Rep0rting" -ForegroundColor Gray
Write-Host "  Settings: password never expires, user cannot change password" -ForegroundColor Gray
Write-Host ""
Write-Host "  [Y]  Create the Reporting user now" -ForegroundColor White
Write-Host "  [N]  Skip user creation" -ForegroundColor White
Write-Host ""
do {
    $userChoice = (Read-Host "  Create Reporting user? (Y/N)").Trim()
} while ($userChoice -notmatch '^[YyNn]$')

if ($userChoice -match '^[Yy]$') {
    try {
        $reportingUsername = "Reporting"
        $reportingPassword = ConvertTo-SecureString "Rep0rting" -AsPlainText -Force

        $existingUser = Get-LocalUser -Name $reportingUsername -ErrorAction SilentlyContinue
        if ($existingUser) {
            Write-Log "Local user '$reportingUsername' already exists - Skipping creation" -Color "Yellow"
        } else {
            New-LocalUser -Name $reportingUsername `
                          -Password $reportingPassword `
                          -PasswordNeverExpires `
                          -UserMayNotChangePassword `
                          -Description "Nixxis Reporting Services account" | Out-Null
            Write-Log "Local user '$reportingUsername' created (pwd never expires, user cannot change pwd)" -Color "Green"
        }
    }
    catch {
        Write-Log "Error creating Reporting user: $($_.Exception.Message)" -Color "Red"
        Write-Log "You may need to create this user manually." -Color "Yellow"
    }
} else {
    Write-Log "Skipped Reporting user creation" -Color "Gray"
}

# =============================================================================
# PHASE 7: CONFIGURE WINDOWS FIREWALL
# =============================================================================

Write-Log "`n=== PHASE 7: CONFIGURING WINDOWS FIREWALL ===" -Color "Cyan"

try {
    $crAppServerExeFw  = Join-Path -Path $nixxisInstallPath -ChildPath "CrAppServer\CrAppServer.exe"
    $fwRuleBaseName    = "Nixxis CrAppServer"

    foreach ($proto in @("TCP", "UDP")) {
        $ruleName = "$fwRuleBaseName $proto"
        # Remove any pre-existing rule with the same name
        Remove-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue

        New-NetFirewallRule -DisplayName $ruleName `
                            -Direction Inbound `
                            -Program $crAppServerExeFw `
                            -Protocol $proto `
                            -Action Allow `
                            -Profile Any | Out-Null
        Write-Log "Firewall rule created: '$ruleName' (Inbound, $proto, Any port, Any profile)" -Color "Green"
    }
}
catch {
    Write-Log "Error configuring firewall: $($_.Exception.Message)" -Color "Red"
    Write-Log "You may need to add the firewall rules manually." -Color "Yellow"
}

# =============================================================================
# PHASE 8: TRANSCRIPTION CLIENT HELPERS
# =============================================================================

Write-Log "`n=== PHASE 8: TRANSCRIPTION CLIENT HELPERS ===" -Color "Cyan"

Write-Host ""
Write-Host "  Download and deploy Transcription Client Helper files from GitHub?" -ForegroundColor White
Write-Host "  Target: $nixxisInstallPath\CrAppServer\provisioning\client" -ForegroundColor Gray
Write-Host ""
Write-Host "  [Y]  Download and deploy" -ForegroundColor White
Write-Host "  [N]  Skip" -ForegroundColor White
Write-Host ""
do {
    $transcriptChoice = (Read-Host "  Deploy Transcription Helpers? (Y/N)").Trim()
} while ($transcriptChoice -notmatch '^[YyNn]$')

if ($transcriptChoice -match '^[Yy]$') {
    try {
        $transcriptDestination = Join-Path -Path $nixxisInstallPath -ChildPath "CrAppServer\provisioning\client"
        $transcriptTempPath    = Join-Path -Path $todayInstallPath  -ChildPath "TranscriptClientHelpers"

        Write-Log "Transcription helpers destination: $transcriptDestination" -Color "White"
        Write-Log "Temp folder: $transcriptTempPath" -Color "White"

        # Create / clean temp folder
        if (Test-Path $transcriptTempPath) {
            Remove-Item -Path $transcriptTempPath -Recurse -Force
        }
        New-Item -Path $transcriptTempPath -ItemType Directory -Force | Out-Null
        Write-Log "Temp folder created" -Color "Green"

        # Ensure destination exists
        if (-not (Test-Path $transcriptDestination)) {
            Write-Log "Destination folder not found - creating: $transcriptDestination" -Color "Yellow"
            New-Item -Path $transcriptDestination -ItemType Directory -Force | Out-Null
        }

        # Enforce TLS 1.2 for GitHub raw content
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $baseUrl = "https://raw.githubusercontent.com/NixxisIntegration/TranscriptionClientHelpers/refs/heads/main"

        $transcriptFiles = @(
            "DarkTranscriptions.css",
            "LightTranscriptions.css",
            "defaultTranscriptions.css",
            "handleTranscriptions.js",
            "handleTranscriptions_en.js",
            "handleTranscriptions_fr.js"
        )

        Write-Log "--- Downloading Transcription Helper files ---" -Color "Cyan"

        $tcDownloaded = @()
        $tcFailed     = @()

        foreach ($fileName in $transcriptFiles) {
            $fileUrl      = "$baseUrl/$fileName"
            $tempFilePath = Join-Path -Path $transcriptTempPath -ChildPath $fileName

            Write-Log "Downloading: $fileName" -Color "Yellow"
            try {
                try {
                    Invoke-WebRequest -Uri $fileUrl -OutFile $tempFilePath -UseBasicParsing -TimeoutSec 60
                }
                catch [System.Net.WebException] {
                    Write-Log "  Retrying with WebClient..." -Color "Yellow"
                    $wc = New-Object System.Net.WebClient
                    $wc.DownloadFile($fileUrl, $tempFilePath)
                    $wc.Dispose()
                }

                if (Test-Path $tempFilePath) {
                    $sz = (Get-Item $tempFilePath).Length
                    Write-Log "  [OK] $fileName ($sz bytes)" -Color "Green"
                    $tcDownloaded += $fileName
                } else {
                    throw "File not found after download"
                }
            }
            catch {
                Write-Log "  [FAILED] $fileName`: $($_.Exception.Message)" -Color "Red"
                $tcFailed += $fileName
            }
        }

        Write-Log "Download complete: $($tcDownloaded.Count) succeeded, $($tcFailed.Count) failed" -Color $(if ($tcFailed.Count -gt 0) { "Yellow" } else { "Green" })

        if ($tcDownloaded.Count -gt 0) {
            Write-Log "--- Deploying to $transcriptDestination ---" -Color "Cyan"

            $tcCopied      = @()
            $tcCopyFailed  = @()

            foreach ($fileName in $tcDownloaded) {
                $src  = Join-Path -Path $transcriptTempPath   -ChildPath $fileName
                $dest = Join-Path -Path $transcriptDestination -ChildPath $fileName
                try {
                    Copy-Item -Path $src -Destination $dest -Force
                    $srcSz  = (Get-Item $src).Length
                    $dstSz  = (Get-Item $dest).Length
                    if ($srcSz -ne $dstSz) { throw "Size mismatch (src=$srcSz, dst=$dstSz)" }
                    Write-Log "  [OK] $fileName deployed" -Color "Green"
                    $tcCopied += $fileName
                }
                catch {
                    Write-Log "  [FAILED] $fileName`: $($_.Exception.Message)" -Color "Red"
                    $tcCopyFailed += $fileName
                }
            }

            Write-Log "Deployment complete: $($tcCopied.Count) deployed, $($tcCopyFailed.Count) failed" -Color $(if ($tcCopyFailed.Count -gt 0) { "Yellow" } else { "Green" })
        }

        # Cleanup temp folder
        Write-Host ""
        $tcCleanup = (Read-Host "  Remove temporary download folder? [Y/N, default Y]").Trim()
        if ($tcCleanup -eq "" -or $tcCleanup -match '^[Yy]$') {
            Remove-Item -Path $transcriptTempPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "Temporary folder removed" -Color "Green"
        } else {
            Write-Log "Temporary folder retained: $transcriptTempPath" -Color "Yellow"
        }

        Write-Log "Transcription Client Helpers phase completed" -Color "Green"
    }
    catch {
        Write-Log "Error during Transcription Helpers phase: $($_.Exception.Message)" -Color "Red"
        Write-Log "Stack Trace: $($_.Exception.StackTrace)" -Color "Red"
        Write-Log "Continuing with remaining installation steps..." -Color "Yellow"
    }
} else {
    Write-Log "Skipped Transcription Client Helpers" -Color "Gray"
}

# FINAL SUMMARY
Write-Log " " -Color "White"
Write-Log "=== SETUP SUMMARY ===" -Color "Cyan"
Write-Log "Installation files location: $todayInstallPath" -Color "White"
Write-Log "Nixxis installation directory: $nixxisInstallPath" -Color "White"
Write-Log "Maintenance folder structure:" -Color "White"
Write-Log "  - Update folder: $updatePath" -Color "White"
Write-Log "  - Backup folder: $backupPath" -Color "White"
Write-Log "  - Install folder: $installPath" -Color "White"
Write-Log "  - Logs folder: $logsPath" -Color "White"
Write-Log " " -Color "White"

Write-Log "=== SETUP COMPLETED SUCCESSFULLY ===" -Color "Green"
Write-Log "Log file saved to: $logFile" -Color "White"
Write-Log " " -Color "White"

# =============================================================================
# POST-INSTALL MANUAL ACTIONS
# =============================================================================

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  POST-INSTALL ACTION REQUIRED - REPORTING SERVER" -ForegroundColor Yellow
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Copy the CONTENTS of the folder:" -ForegroundColor White
Write-Host "  $nixxisInstallPath\Reporting\ToCopyReportingServer" -ForegroundColor Cyan
Write-Host ""
Write-Host "  INTO (accept folder merge when prompted):" -ForegroundColor White
Write-Host "  [Drive]:\Program Files\Microsoft SQL Server\MSRS.?\" -ForegroundColor Cyan
Write-Host "          Reporting Services\ReportServer\bin" -ForegroundColor Cyan
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# Launch DeployReports.exe
$deployReportsExe = Join-Path -Path $nixxisInstallPath -ChildPath "Reporting\Deploy\DeployReports.exe"
if (Test-Path $deployReportsExe) {
    Write-Log "Launching DeployReports.exe..." -Color "Cyan"
    Start-Process -FilePath $deployReportsExe
    Write-Log "DeployReports.exe launched" -Color "Green"
} else {
    Write-Log "WARNING: DeployReports.exe not found at $deployReportsExe" -Color "Yellow"
    Write-Log "Please launch it manually once Reporting Services is configured." -Color "Yellow"
}

Write-Host "Press any key to close..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
