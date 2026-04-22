# Script to prepare files, stop Nixxis service, perform backup, cleanup files & deploy.
# Version 1.4 - FP - Added Offline Mode Support - 2025-06-18

# Script Parameters - Allow users to specify custom URLs for ZIP files or use offline mode
param(
    [string]$ClientProvisioningUrl = "",
    [string]$ClientSoftwareUrl = "",
    [string]$NCSUrl = "",
    [switch]$OfflineMode,
    [string]$OfflinePath = "",
    [switch]$Help
)

# Display help information
if ($Help) {
    Write-Host "Nixxis Maintenance Script v1.4" -ForegroundColor Cyan
    Write-Host "================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "DESCRIPTION:" -ForegroundColor Yellow
    Write-Host "  Automatically downloads, backs up, and deploys Nixxis application updates."
    Write-Host "  By default, the script auto-discovers the latest versions from update.nixxis.net"
    Write-Host "  Can also operate in offline mode using pre-downloaded ZIP files."
    Write-Host ""
    Write-Host "USAGE:" -ForegroundColor Yellow
    Write-Host "  .\NixxisMaintenance.ps1 [parameters]"
    Write-Host ""
    Write-Host "PARAMETERS:" -ForegroundColor Yellow
    Write-Host "  -ClientProvisioningUrl <string>   Custom URL for ClientProvisioning.zip"
    Write-Host "  -ClientSoftwareUrl <string>       Custom URL for ClientSoftware.zip"
    Write-Host "  -NCSUrl <string>                  Custom URL for NCS.zip"
    Write-Host "  -OfflineMode                      Use ZIP files from local folder (no download)"
    Write-Host "  -OfflinePath <string>             Folder path containing ZIP files (for offline mode)"
    Write-Host "  -Help                             Show this help message"
    Write-Host ""
    Write-Host "EXAMPLES:" -ForegroundColor Yellow
    Write-Host "  # Auto-discover latest versions (default behavior)"
    Write-Host "  .\NixxisMaintenance.ps1"
    Write-Host ""
    Write-Host "  # Use offline mode with ZIP files in current folder"
    Write-Host "  .\NixxisMaintenance.ps1 -OfflineMode"
    Write-Host ""
    Write-Host "  # Use offline mode with ZIP files in specific folder"
    Write-Host "  .\NixxisMaintenance.ps1 -OfflineMode -OfflinePath 'C:\NixxisZips'"
    Write-Host ""
    Write-Host "  # Use custom URL for NCS.zip only"
    Write-Host "  .\NixxisMaintenance.ps1 -NCSUrl 'http://custom.server.com/NCS.zip'"
    Write-Host ""
    Write-Host "  # Use custom URLs for all files"
    Write-Host "  .\NixxisMaintenance.ps1 ``"
    Write-Host "    -ClientProvisioningUrl 'http://server.com/ClientProvisioning.zip' ``"
    Write-Host "    -ClientSoftwareUrl 'http://server.com/ClientSoftware.zip' ``"
    Write-Host "    -NCSUrl 'http://server.com/NCS.zip'"
    Write-Host ""
    Write-Host "OFFLINE MODE REQUIREMENTS:" -ForegroundColor Yellow
    Write-Host "  The following files must be present in the offline folder:"
    Write-Host "  - ClientProvisioning.zip"
    Write-Host "  - ClientSoftware.zip"
    Write-Host "  - NCS.zip"
    Write-Host ""
    Write-Host "REQUIREMENTS:" -ForegroundColor Yellow
    Write-Host "  - Must run as Administrator"
    Write-Host "  - PowerShell 5.1 or higher"
    Write-Host "  - Internet connectivity for downloads (not needed in offline mode)"
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

# Administrator privilege check
if (-not (Test-Administrator)) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Please right-click on PowerShell and select 'Run as Administrator', then run this script again." -ForegroundColor Yellow
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-Host "Administrator privileges confirmed - Proceeding with script execution..." -ForegroundColor Green

# Initialize logging
$logDate = Get-Date -Format "yyyyMMdd_HHmmss"
$logPath = "C:\NixxisMaintenance\Logs"
$logFile = Join-Path -Path $logPath -ChildPath "NixxisMaintenance_$logDate.log"

# Create log directory if it doesn't exist
if (-not (Test-Path $logPath)) {
    New-Item -Path $logPath -ItemType Directory -Force | Out-Null
}

# Function to write to both log and console
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [string]$Color = "White"
    )
    
    $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timeStamp - $Message"
    
    # Write to log file
    Add-Content -Path $logFile -Value $logMessage
    
    # Write to console with color
    Write-Host $logMessage -ForegroundColor $Color
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

# Function to copy offline ZIP files to script directory
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
        # Handle both version formats (v3.2) and numeric formats (3.2.1086)
        $sortedFolders = $folders | Sort-Object { 
            # Extract numeric parts for proper version sorting
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
        
        # Download with progress (for PowerShell 5.1 compatibility)
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

# Function to fetch ZIP files from web
function Get-NixxisZipFiles {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ScriptDirectory,
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
        
        # Use custom URLs where provided, otherwise use auto-discovered URLs
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
        
        # Define local file paths
        $clientProvisioningPath = Join-Path -Path $ScriptDirectory -ChildPath "ClientProvisioning.zip"
        $clientSoftwarePath = Join-Path -Path $ScriptDirectory -ChildPath "ClientSoftware.zip"
        $ncsPath = Join-Path -Path $ScriptDirectory -ChildPath "NCS.zip"
        
        # Display final URLs that will be used
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
            $zipPath = Join-Path -Path $ScriptDirectory -ChildPath $zip
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

# Function to handle file deletion with retry
function Remove-FileWithRetry {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        [int]$MaxAttempts = 3,
        [int]$RetryDelaySeconds = 2
    )

    $attempts = 0
    while ($attempts -lt $MaxAttempts) {
        try {
            if (Test-Path $FilePath) {
                Remove-Item -Path $FilePath -Force -ErrorAction Stop
                Write-Log "Successfully deleted: $FilePath" -Color "Green"
                return $true
            }
            return $true  # File doesn't exist, consider it a success
        }
        catch {
            $attempts++
            if ($attempts -lt $MaxAttempts) {
                Write-Log "Attempt $attempts failed to delete $FilePath - Retrying in $RetryDelaySeconds seconds..." -Color "Yellow"
                Start-Sleep -Seconds $RetryDelaySeconds
            }
            else {
                Write-Log "Failed to delete $FilePath after $MaxAttempts attempts. Error: $($_.Exception.Message)" -Color "Red"
                return $false
            }
        }
    }
    return $false
}

# Get script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Determine operation mode and handle ZIP files
if ($OfflineMode) {
    Write-Log "=== OFFLINE MODE ACTIVATED ===" -Color "Magenta"
    
    # Determine offline folder path
    $offlineFolder = if ($OfflinePath) {
        Write-Log "Using specified offline path: $OfflinePath" -Color "White"
        $OfflinePath
    } else {
        Write-Log "No offline path specified - Using script directory: $scriptDir" -Color "White"
        $scriptDir
    }
    
    try {
        # Verify offline ZIP files exist
        Test-OfflineZipFiles -FolderPath $offlineFolder
        
        # If offline folder is different from script directory, copy files
        if ($offlineFolder -ne $scriptDir) {
            Write-Log "Offline folder differs from script directory - Copying ZIP files..." -Color "Yellow"
            Copy-OfflineZipFiles -SourcePath $offlineFolder -DestinationPath $scriptDir
        } else {
            Write-Log "Offline ZIP files are already in script directory - No copy needed" -Color "Green"
        }
        
        Write-Log "Offline mode preparation completed successfully!" -Color "Green"
    }
    catch {
        Write-Log "Error during offline mode preparation: $($_.Exception.Message)" -Color "Red"
        Write-Log "Stack Trace: $($_.Exception.StackTrace)" -Color "Red"
        exit 1
    }
} else {
    # Online mode - Download files
    Write-Log "=== ONLINE MODE - DOWNLOADING FILES ===" -Color "Cyan"
    try {
        # Download ZIP files from web (with custom URL support)
        Get-NixxisZipFiles -ScriptDirectory $scriptDir -CustomClientProvisioningUrl $ClientProvisioningUrl -CustomClientSoftwareUrl $ClientSoftwareUrl -CustomNCSUrl $NCSUrl
        
        Write-Log "Web download phase completed successfully!" -Color "Green"
    }
    catch {
        Write-Log "Error during web download phase: $($_.Exception.Message)" -Color "Red"
        Write-Log "Stack Trace: $($_.Exception.StackTrace)" -Color "Red"
        exit 1
    }
}

# Preparation Phase
Write-Log "Starting file preparation phase..." -Color "Cyan"
try {
    # Verify zip files exist (they should now be downloaded or copied)
    $requiredZips = @("NCS.zip", "ClientProvisioning.zip", "ClientSoftware.zip")
    foreach ($zip in $requiredZips) {
        $zipPath = Join-Path -Path $scriptDir -ChildPath $zip
        if (-not (Test-Path $zipPath)) {
            throw "Required zip file not found: $zip"
        }
    }

    # Create/Clean NixxisApplicationServer folder
    $nixxisAppServer = Join-Path -Path $scriptDir -ChildPath "NixxisApplicationServer"
    if (Test-Path $nixxisAppServer) {
        Write-Log "Removing existing NixxisApplicationServer folder..." -Color "Yellow"
        Remove-Item -Path $nixxisAppServer -Recurse -Force
    }
    New-Item -Path $nixxisAppServer -ItemType Directory -Force | Out-Null
    Write-Log "Created NixxisApplicationServer folder" -Color "Green"

    # Extract NCS.zip
    Write-Log "Extracting NCS.zip..." -Color "Yellow"
    Expand-Archive -Path (Join-Path -Path $scriptDir -ChildPath "NCS.zip") -DestinationPath $nixxisAppServer -Force
    Write-Log "Extracted NCS.zip successfully" -Color "Green"

    # Create and extract ClientSoftware
    $clientSoftwarePath = Join-Path -Path $nixxisAppServer -ChildPath "ClientSoftware"
    New-Item -Path $clientSoftwarePath -ItemType Directory -Force | Out-Null
    Write-Log "Created ClientSoftware folder" -Color "Green"

    Write-Log "Extracting ClientSoftware.zip..." -Color "Yellow"
    Expand-Archive -Path (Join-Path -Path $scriptDir -ChildPath "ClientSoftware.zip") -DestinationPath $clientSoftwarePath -Force
    Write-Log "Extracted ClientSoftware.zip successfully" -Color "Green"

    # Create provisioning folder
    $provisioningPath = Join-Path -Path $nixxisAppServer -ChildPath "CrAppServer\provisioning"
    New-Item -Path $provisioningPath -ItemType Directory -Force | Out-Null
    Write-Log "Created provisioning folder" -Color "Green"

    # Copy ClientSoftware.zip to provisioning
    Copy-Item -Path (Join-Path -Path $scriptDir -ChildPath "ClientSoftware.zip") -Destination $provisioningPath -Force
    Write-Log "Copied ClientSoftware.zip to provisioning folder" -Color "Green"

    # Extract ClientProvisioning.zip
    $provisioningClientPath = Join-Path -Path $provisioningPath -ChildPath "client"
    New-Item -Path $provisioningClientPath -ItemType Directory -Force | Out-Null
    Write-Log "Extracting ClientProvisioning.zip..." -Color "Yellow"
    Expand-Archive -Path (Join-Path -Path $scriptDir -ChildPath "ClientProvisioning.zip") -DestinationPath $provisioningClientPath -Force
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

    Write-Log "File preparation phase completed successfully!" -Color "Green"
}
catch {
    Write-Log "Error during preparation phase: $($_.Exception.Message)" -Color "Red"
    Write-Log "Stack Trace: $($_.Exception.StackTrace)" -Color "Red"
    exit 1
}

# Initialize variables
$serviceName = "crappserver"
$processName = "crappserver"
Write-Log "Starting Nixxis maintenance script" -Color "Cyan"

# Kill Nixxis Client Desktop process before stopping service
Write-Log "Attempting to kill Nixxis Client Desktop process..." -Color "Cyan"
try {
    # Check if nixxisclientdesktop.exe is running
    $clientProcess = Get-Process -Name "nixxisclientdesktop" -ErrorAction SilentlyContinue
    
    if ($clientProcess) {
        Write-Log "Found nixxisclientdesktop.exe process(es) running with PID(s): $($clientProcess.Id -join ', ')" -Color "Yellow"
        
        # Use taskkill to forcefully terminate the process
        $taskkillOutput = & taskkill /IM nixxisclientdesktop.exe /F 2>&1
        $taskkillExitCode = $LASTEXITCODE
        
        Write-Log "Taskkill command output: $taskkillOutput" -Color "Gray"
        Write-Log "Taskkill exit code: $taskkillExitCode" -Color "Gray"
        
        if ($taskkillExitCode -eq 0) {
            Write-Log "Successfully killed nixxisclientdesktop.exe process" -Color "Green"
        } else {
            Write-Log "Taskkill command completed with exit code $taskkillExitCode" -Color "Yellow"
        }
        
        # Verify the process is terminated
        Start-Sleep -Seconds 2
        $remainingProcess = Get-Process -Name "nixxisclientdesktop" -ErrorAction SilentlyContinue
        if ($remainingProcess) {
            Write-Log "WARNING: nixxisclientdesktop.exe process still running after taskkill attempt" -Color "Red"
        } else {
            Write-Log "Confirmed: nixxisclientdesktop.exe process has been terminated" -Color "Green"
        }
    } else {
        Write-Log "No nixxisclientdesktop.exe process found running" -Color "Green"
    }
}
catch {
    Write-Log "Error during client process termination: $($_.Exception.Message)" -Color "Red"
    Write-Log "Continuing with service stop process..." -Color "Yellow"
}

Write-Log "Attempting to stop Nixxis Crappserver service..." -Color "Cyan"

# Initial service stop
try {
    Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
}
catch {
    Write-Log "Initial service stop attempt failed: $($_.Exception.Message)" -Color "Yellow"
}

$attempts = 0
$maxAttempts = 60  # 120 seconds total

Do {
    $service = Get-Service -Name $serviceName
    $process = Get-Process -Name $processName -ErrorAction SilentlyContinue
    
    if ($service.Status -eq 'Stopped' -and $null -eq $process) {
        Start-Sleep -Seconds 5  # Double verification wait
        $process = Get-Process -Name $processName -ErrorAction SilentlyContinue
        
        if ($null -eq $process) {
            Write-Log "Nixxis Crappserver is fully stopped and process is terminated!" -Color "Green"
            break
        }
    }
    
    $attempts++
    Write-Log "Waiting for service and process to fully stop... Attempt $attempts of $maxAttempts" -Color "Yellow"
    
    if ($null -ne $process) {
        Write-Log "Process still running with PID: $($process.Id)" -Color "Yellow"
        Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
    }
    
    Start-Sleep -Seconds 2
    
} Until ($attempts -ge $maxAttempts)

if ($attempts -ge $maxAttempts) {
    Write-Log "Failed to fully stop Nixxis Crappserver after 120 seconds!" -Color "Red"
    $remainingProcess = Get-Process -Name $processName -ErrorAction SilentlyContinue
    if ($null -ne $remainingProcess) {
        Write-Log "Process still running with PID: $($remainingProcess.Id)" -Color "Red"
    }
    exit 1
}

Write-Log "Service stop verification complete - Proceeding with next actions" -Color "Green"

# Backup Process
if ($?) {
    Write-Log "Starting backup process..." -Color "Cyan"
    
    try {
        # Define base paths
        $baseBackupPath = "C:\NixxisMaintenance\BackUp"
        $yearFolder = (Get-Date).Year.ToString()
        $dateFolder = Get-Date -Format "yyyyMMdd"
        
        # Create folder structure
        $yearPath = Join-Path -Path $baseBackupPath -ChildPath $yearFolder
        $datePath = Join-Path -Path $yearPath -ChildPath $dateFolder
        $nmsPath = Join-Path -Path $datePath -ChildPath "NMS"
        $SQLPath = Join-Path -Path $datePath -ChildPath "SQL"
        
        # Create main folders
        New-Item -Path $yearPath -ItemType Directory -Force | Out-Null
        New-Item -Path $datePath -ItemType Directory -Force | Out-Null
        New-Item -Path $nmsPath -ItemType Directory -Force | Out-Null
        New-Item -Path $SQLPath -ItemType Directory -Force | Out-Null
        
        Write-Log "Created backup directory structure" -Color "Green"
        
        # Create NMS subfolders
        $nmsSubfolders = @("NMS1", "NMS2", "NMS3", "NMS4")
        foreach ($folder in $nmsSubfolders) {
            $subfolderPath = Join-Path -Path $nmsPath -ChildPath $folder
            New-Item -Path $subfolderPath -ItemType Directory -Force | Out-Null
            Write-Log "Created NMS subfolder: $folder" -Color "Green"
        }
		
		# Source paths for copying
        $crAppServerSource = "C:\Nixxis\CrAppServer"
        $clientSoftwareSource = "C:\Nixxis\ClientSoftware"
        
        # Copy CrAppServer folder if it exists
        if (Test-Path $crAppServerSource) {
            Write-Log "Copying CrAppServer folder..." -Color "Yellow"
            Copy-Item -Path $crAppServerSource -Destination $datePath -Recurse -Force
            Write-Log "CrAppServer folder copied successfully" -Color "Green"
        }
        else {
            Write-Log "CrAppServer folder not found at $crAppServerSource - Skipping backup" -Color "Yellow"
        }
        
        # Copy ClientSoftware folder if it exists
        if (Test-Path $clientSoftwareSource) {
            Write-Log "Copying ClientSoftware folder..." -Color "Yellow"
            Copy-Item -Path $clientSoftwareSource -Destination $datePath -Recurse -Force
            Write-Log "ClientSoftware folder copied successfully" -Color "Green"
        }
        else {
            Write-Log "ClientSoftware folder not found at $clientSoftwareSource - Skipping backup" -Color "Yellow"
        }
        
        # Cleanup Process
        Write-Log "Starting file cleanup process..." -Color "Cyan"
        
        # Define array of specific files to delete
        $filesToDelete = @(
            "C:\Nixxis\CrAppServer\AdminLink.dll",
            "C:\Nixxis\CrAppServer\agsXMPP.dll",
            "C:\Nixxis\CrAppServer\CRAppServerInterfaces.dll",
            "C:\Nixxis\CrAppServer\CrAppServerReferences.dll",
            "C:\Nixxis\CrAppServer\CRShared.dll",
            "C:\Nixxis\CrAppServer\Microsoft.ReportViewer.Common.dll",
            "C:\Nixxis\CrAppServer\Microsoft.ReportViewer.WinForms.dll",
            "C:\Nixxis\CrAppServer\Nixxis.Install.dll",
            "C:\Nixxis\CrAppServer\NixxisClientReferences.dll",
            "C:\Nixxis\CrAppServer\RestSharp.dll",
            "C:\Nixxis\CrAppServer\SampleConfigurationFiles.dll",
            "C:\Nixxis\CrAppServer\SocialMediaElements.dll",
            "C:\Nixxis\CrAppServer\SFAgent.dll",
            "C:\Nixxis\CrAppServer\System.Windows.Controls.DataVisualization.Toolkit.dll",
            "C:\Nixxis\CrAppServer\Twitterizer2.dll",
            "C:\Nixxis\CrAppServer\UscNetTools.dll",
            "C:\Nixxis\CrAppServer\CRReportingServer.exe",
            #"C:\Nixxis\CrAppServer\CrAppServer.exe.lic",
            "C:\Nixxis\CrAppServer\ManitermSettings.dll",
            "C:\Nixxis\CrAppServer\NCS.TranscriberSettings.dll",
            "C:\Nixxis\CrAppServer\NixxisSalesForceClient.dll",
            "C:\Nixxis\CrAppServer\TicketControllerSettings.dll",
            "C:\Nixxis\CrAppServer\WebChatSettings.dll",
            "C:\Nixxis\CrAppServer\Plugins\ManitermSettings.dll",
            "C:\Nixxis\CrAppServer\Plugins\TicketsControllerSettings.dll",
            "C:\Nixxis\CrAppServer\Default_reporting\StateMachineContactList.bin",
            "C:\Nixxis\CrAppServer\TicketsControllerSettings.dll",
            "C:\Nixxis\CrAppServer\TicketsController.dll"
        )
# Verify CrAppServer directory exists before proceeding
        $crAppServerPath = "C:\Nixxis\CrAppServer"
        if (-not (Test-Path $crAppServerPath)) {
            Write-Log "CrAppServer directory not found at $crAppServerPath" -Color "Yellow"
        }
        else {
            # Delete specific files with retry
            $failedDeletions = @()
            foreach ($file in $filesToDelete) {
                if (Test-Path $file) {
                    if (-not (Remove-FileWithRetry -FilePath $file)) {
                        $failedDeletions += $file
                    }
                }
                else {
                    Write-Log "File not found (skipping): $file" -Color "Yellow"
                }
            }

            # Delete .pdb files (only in root directory)
            Write-Log "Checking for .pdb files in root directory..." -Color "Yellow"
            $pdbFiles = Get-ChildItem -Path $crAppServerPath -Filter "*.pdb" -File -ErrorAction SilentlyContinue
            if ($pdbFiles) {
                Write-Log "Found $($pdbFiles.Count) .pdb files to remove" -Color "Yellow"
                foreach ($pdbFile in $pdbFiles) {
                    if (-not (Remove-FileWithRetry -FilePath $pdbFile.FullName)) {
                        $failedDeletions += $pdbFile.FullName
                    }
                }
            }
            else {
                Write-Log "No .pdb files found in root directory" -Color "Yellow"
            }

            # Check and clean Provisioning folder
            $provisioningPath = Join-Path -Path $crAppServerPath -ChildPath "Provisioning"
            Write-Log "Checking Provisioning folder..." -Color "Yellow"
            if (Test-Path $provisioningPath) {
                $provisioningItems = Get-ChildItem -Path $provisioningPath -ErrorAction SilentlyContinue | 
                                   Where-Object { $_.Name -ne "Settings" }
                
                if ($provisioningItems) {
                    Write-Log "Found $($provisioningItems.Count) items to remove from Provisioning" -Color "Yellow"
                    foreach ($item in $provisioningItems) {
                        Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
                        if ($?) {
                            Write-Log "Deleted from Provisioning: $($item.Name)" -Color "Green"
                        }
                        else {
                            $failedDeletions += $item.FullName
                            Write-Log "Failed to delete from Provisioning: $($item.Name)" -Color "Red"
                        }
                    }
                }
                else {
                    Write-Log "No items to remove from Provisioning folder (excluding Settings)" -Color "Yellow"
                }
            }
            else {
                Write-Log "Provisioning folder not found at $provisioningPath" -Color "Yellow"
            }
        }
		
		# Check and delete ClientSoftware folder
        Write-Log "Checking ClientSoftware folder..." -Color "Yellow"
        $clientSoftwarePath = "C:\Nixxis\ClientSoftware"
        if (Test-Path $clientSoftwarePath) {
            Write-Log "Found ClientSoftware folder, attempting to delete..." -Color "Yellow"
            Remove-Item -Path $clientSoftwarePath -Recurse -Force -ErrorAction SilentlyContinue
            if ($?) {
                Write-Log "Deleted ClientSoftware folder successfully" -Color "Green"
            }
            else {
                $failedDeletions += $clientSoftwarePath
                Write-Log "Failed to delete ClientSoftware folder" -Color "Red"
            }
        }
        else {
            Write-Log "ClientSoftware folder not found at $clientSoftwarePath" -Color "Yellow"
        }

        # Report any failed deletions
        if ($failedDeletions.Count -gt 0) {
            Write-Log "WARNING: The following files/folders could not be deleted:" -Color "Red"
            foreach ($failedFile in $failedDeletions) {
                Write-Log "- $failedFile" -Color "Red"
            }
        }
        else {
            Write-Log "All requested deletions completed successfully" -Color "Green"
        }

        Write-Log "Maintenance process completed!" -Color "Green"

        # Start deployment phase after maintenance is complete
        Write-Log "Starting deployment phase..." -Color "Cyan"
        try {
            # Get script directory and construct source base directory
            $sourceBaseDir = Join-Path -Path $scriptDir -ChildPath "NixxisApplicationServer"

            # Define folders to copy
            $foldersToCopy = @(
                @{Source="ClientSoftware"; Destination="C:\Nixxis\ClientSoftware"},
                @{Source="CrAppServer"; Destination="C:\Nixxis\CrAppServer"},
                @{Source="MediaServer"; Destination="C:\Nixxis\MediaServer"},
                @{Source="Reporting"; Destination="C:\Nixxis\Reporting"},
                @{Source="SampleConfigFiles"; Destination="C:\Nixxis\SampleConfigFiles"},
                @{Source="SoundsSamples"; Destination="C:\Nixxis\SoundsSamples"}
            )

            # Create C:\Nixxis if it doesn't exist
            $nixxisPath = "C:\Nixxis"
            if (-not (Test-Path $nixxisPath)) {
                Write-Log "Creating Nixxis directory at $nixxisPath" -Color "Yellow"
                New-Item -Path $nixxisPath -ItemType Directory -Force | Out-Null
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
                        throw  # Re-throw to be caught by outer try-catch
                    }
                }
                else {
                    Write-Log "Source folder not found: $sourcePath" -Color "Red"
                    throw "Source folder not found: $sourcePath"
                }
            }

            Write-Log "Deployment process completed successfully" -Color "Green"
        }
        catch {
            Write-Log "Error during deployment process: $($_.Exception.Message)" -Color "Red"
            Write-Log "Stack Trace: $($_.Exception.StackTrace)" -Color "Red"
            exit 1
        }
    }
    catch {
        Write-Log "Error during maintenance process: $($_.Exception.Message)" -Color "Red"
        Write-Log "Stack Trace: $($_.Exception.StackTrace)" -Color "Red"
        exit 1
    }
}

Write-Log "Script execution completed" -Color "Cyan"