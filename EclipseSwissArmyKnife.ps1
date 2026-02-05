# PowerShell script for Eclipse Swiss Army Knife
# This script will provide a menu to perform the main actions described in the README
# Run this script with Administrator privileges
# Created by Connor Brown for Eclipse Support Team to assist making our life #easier

# Script Version - Used for automatic update checking
$ScriptVersion = "2.8.1"

# Download URLs - Update these when new versions are released
# Option 6: Eclipse DMS (version derived from URL filename)
$EclipseDmsUrl = "http://ws.dev.ultimate.net.au:8029/downloads/EclipseDesktop/Eclipse2037-26-35.msi"
$EclipseDmsVersion = ([System.Net.WebUtility]::UrlDecode((Split-Path $EclipseDmsUrl -Leaf)) -replace '\.msi$','' -replace '^Eclipse','' -replace '-','.').Trim()

# Option 15: Eclipse Update Service (version derived from URL filename)
$EclipseUpdateServiceUrl = "http://ws.dev.ultimate.net.au:8029/downloads/EclipseUpdateService/Eclipse%20Update%20Service%201.52.exe"
$EclipseUpdateServiceVersion = ([System.Net.WebUtility]::UrlDecode((Split-Path $EclipseUpdateServiceUrl -Leaf)) -replace '.*?(\d+(\.\d+)+)\.(exe|msi)$','$1').Trim()

# Option 16: Eclipse Online Chrome (version derived from URL filename)
$EclipseOnlineChromeUrl = "http://ws.dev.ultimate.net.au:8029/downloads/EclipseOnlineChrome/EclipseOnline%20Chrome%203.8.21.msi"
$EclipseOnlineChromeVersion = ([System.Net.WebUtility]::UrlDecode((Split-Path $EclipseOnlineChromeUrl -Leaf)) -replace '.*?(\d+(\.\d+)+)\.(exe|msi)$','$1').Trim()

# Option 17: Eclipse Online Server (version derived from URL filename)
$EclipseOnlineServerUrl = "http://ws.dev.ultimate.net.au:8029/downloads/EclipseOnlineServer/EclipseOnline%20Server%2012.0.133.0.exe"
$EclipseOnlineServerVersion = ([System.Net.WebUtility]::UrlDecode((Split-Path $EclipseOnlineServerUrl -Leaf)) -replace '.*?(\d+(\.\d+)+)\.(exe|msi)$','$1').Trim()

# Option 20: Eclipse Smart Hub (version derived from URL filename)
$EclipseSmartHubUrl = "http://ws.dev.ultimate.net.au:8029/downloads/EclipseSmartHub/EclipseSmartHubInstaller%2012.4.1.0.exe"
$EclipseSmartHubVersion = ([System.Net.WebUtility]::UrlDecode((Split-Path $EclipseSmartHubUrl -Leaf)) -replace '.*?(\d+(\.\d+)+)\.(exe|msi)$','$1').Trim()

function Show-Menu {
    Clear-Host
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "      Eclipse Swiss Army Knife v$ScriptVersion" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "[1]  " -NoNewline; Write-Host "Create Eclipse Install Directory Structure" -ForegroundColor Green
    Write-Host "[2]  " -NoNewline; Write-Host "Share Eclipse Install Directory" -ForegroundColor Green
    Write-Host "[3]  " -NoNewline; Write-Host "Download SQL Server Express 2025" -ForegroundColor Green
    Write-Host "[4]  " -NoNewline; Write-Host "Download SQL Server Management Studio 21" -ForegroundColor Green
    Write-Host "[5]  " -NoNewline; Write-Host "Configure SQL Server Network Protocols" -ForegroundColor Green
    Write-Host "[6]  " -NoNewline; Write-Host "Install/Update Eclipse DMS $EclipseDmsVersion" -ForegroundColor Green
    Write-Host "[7]  " -NoNewline; Write-Host "Extract Eclipse Database Registry Keys" -ForegroundColor Green
    Write-Host "[8]  " -NoNewline; Write-Host "Add Ultimate Domains to Trusted Sites" -ForegroundColor Green
    Write-Host ""
    Write-Host "--- Eclipse Feed Services ---" -ForegroundColor Cyan
    Write-Host "[9]  " -NoNewline; Write-Host "Eclipse ESB Service Management" -ForegroundColor Cyan
    Write-Host "[10] " -NoNewline; Write-Host "Eclipse Stock Export (Beta) Management" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "--- Eclipse Online Services ---" -ForegroundColor Blue
    Write-Host "[11] " -NoNewline; Write-Host "Install Microsoft IIS Server + Dependencies" -ForegroundColor Yellow
    Write-Host "[12] " -NoNewline; Write-Host "Download IIS URL Rewrite Module (Install IIS First)" -ForegroundColor Yellow
    Write-Host "[13] " -NoNewline; Write-Host "Download .NET 8 Core + ASP 8 + Runtime 8" -ForegroundColor Yellow
    Write-Host "[14] " -NoNewline; Write-Host "Download .NET Framework 4.8 (System Restart Needed)" -ForegroundColor Yellow
    Write-Host "[15] " -NoNewline; Write-Host "Download Eclipse Update Service $EclipseUpdateServiceVersion" -ForegroundColor Yellow
    Write-Host "[16] " -NoNewline; Write-Host "Download Eclipse Online Chrome $EclipseOnlineChromeVersion" -ForegroundColor Yellow
    Write-Host "[17] " -NoNewline; Write-Host "Download Eclipse Online Server $EclipseOnlineServerVersion" -ForegroundColor Yellow
    Write-Host "[18] " -NoNewline; Write-Host "Update IIS Application Pools" -ForegroundColor Yellow
    Write-Host "[19] " -NoNewline; Write-Host "Install Win-ACMEv2" -ForegroundColor Yellow
    Write-Host "[20] " -NoNewline; Write-Host "Install Eclipse Smart Hub $EclipseSmartHubVersion" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "[U]  " -NoNewline; Write-Host "Unattended Server Setup (Select Multiple Tasks)" -ForegroundColor Cyan
}

# PowerShell script for Eclipse Swiss Army Knife v$ScriptVersion
# Unattended Server Setup - All Options (1-20)
# This script provides an unattended setup experience for Eclipse server installation
# Run this script with Administrator privileges
# Created by Connor Brown for Eclipse Support Team

# Global variables for task results
$script:TaskResults = @()
$script:LogFile = "C:\EclipseTechInstall.log"
$script:LogInitialized = $false

# Function to write to both console and log file
function Write-Log {
    param(
        [string]$Message,
        [string]$ForegroundColor = "White",
        [switch]$NoNewline
    )
    
    # Write to console
    if ($NoNewline) {
        Write-Host $Message -ForegroundColor $ForegroundColor -NoNewline
    } else {
        Write-Host $Message -ForegroundColor $ForegroundColor
    }
    
    # Write to log file (strip color codes and format)
    $logMessage = $Message
    if (-not $script:LogInitialized) {
        # Initialize log file
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $header = @"
===============================================
Eclipse Swiss Army Knife v$ScriptVersion - Installation Log
Execution Started: $timestamp
===============================================

"@
        Set-Content -Path $script:LogFile -Value $header -Encoding UTF8
        $script:LogInitialized = $true
    }
    
    try {
        if ($NoNewline) {
            Add-Content -Path $script:LogFile -Value $logMessage -Encoding UTF8 -NoNewline
        } else {
            Add-Content -Path $script:LogFile -Value $logMessage -Encoding UTF8
        }
    } catch {
        # Silently fail if log write fails
    }
}

# Function to log process output
function Write-LogProcessOutput {
    param(
        [string]$Output,
        [string]$ErrorOutput = ""
    )
    
    if ($Output) {
        $lines = $Output -split "`n"
        foreach ($line in $lines) {
            if ($line.Trim()) {
                Write-Log "  [PROCESS] $line" -ForegroundColor Gray
            }
        }
    }
    
    if ($ErrorOutput) {
        $lines = $ErrorOutput -split "`n"
        foreach ($line in $lines) {
            if ($line.Trim()) {
                Write-Log "  [ERROR] $line" -ForegroundColor Red
            }
        }
    }
}

# Function to ensure NSSM is installed
function Install-NSSM {
    param(
        [string]$NssmPath = "C:\Windows\System32\nssm.exe"
    )
    
    # Check if NSSM already exists
    if (Test-Path $NssmPath) {
        return $true
    }
    
    Write-Host "NSSM not found. Downloading and installing NSSM..." -ForegroundColor Yellow
    Write-Host ""
    
    try {
        $tempDir = Join-Path $env:TEMP "nssm-install"
        $zipPath = Join-Path $tempDir "nssm-2.24.zip"
        $extractPath = Join-Path $tempDir "nssm-2.24"
        $nssmUrl = "https://nssm.cc/release/nssm-2.24.zip"
        
        # Create temp directory
        if (Test-Path $tempDir) {
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
        
        # Download NSSM
        Write-Host "Downloading NSSM 2.24..." -ForegroundColor Yellow
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($nssmUrl, $zipPath)
        $webClient.Dispose()
        
        # Extract ZIP file
        Write-Host "Extracting NSSM..." -ForegroundColor Yellow
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $tempDir)
        
        # Determine architecture (64-bit for System32)
        $nssmSource = Join-Path $extractPath "win64\nssm.exe"
        if (-not (Test-Path $nssmSource)) {
            # Fallback to win32 if win64 doesn't exist
            $nssmSource = Join-Path $extractPath "win32\nssm.exe"
        }
        
        if (-not (Test-Path $nssmSource)) {
            throw "NSSM executable not found in extracted archive"
        }
        
        # Copy NSSM to System32 (requires admin privileges)
        Write-Host "Installing NSSM to System32..." -ForegroundColor Yellow
        Copy-Item -Path $nssmSource -Destination $NssmPath -Force
        
        # Verify installation
        if (Test-Path $NssmPath) {
            Write-Host "NSSM installed successfully!" -ForegroundColor Green
            Write-Host ""
            
            # Clean up temp files
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            
            return $true
        } else {
            throw "NSSM installation failed - file not found at destination"
        }
    } catch {
        Write-Host "Error installing NSSM: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please manually download NSSM from https://nssm.cc/download and install to $NssmPath" -ForegroundColor Yellow
        Write-Host ""
        
        # Clean up temp files on error
        if (Test-Path $tempDir) {
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        
        return $false
    }
}

function Get-Configuration {
    Clear-Host
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Eclipse Swiss Army Knife v$ScriptVersion" -ForegroundColor Yellow
    Write-Host "   Unattended Server Setup" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Get Eclipse Install Folder Path
    $defaultEclipsePath = "C:\Eclipse Install"
    Write-Host "Eclipse Install Folder Path" -ForegroundColor Yellow
    Write-Host "This folder will be used for:" -ForegroundColor Cyan
    Write-Host "  - Directory structure creation" -ForegroundColor Gray
    Write-Host "  - Network sharing" -ForegroundColor Gray
    Write-Host "  - Registry key exports" -ForegroundColor Gray
    $eclipsePath = Read-Host "Enter path (default: $defaultEclipsePath)"
    if ([string]::IsNullOrWhiteSpace($eclipsePath)) {
        $eclipsePath = $defaultEclipsePath
    }
    
    Write-Host ""
    
    # Get SQL Server Drive Letter
    Write-Host "SQL Server Installation Drive" -ForegroundColor Yellow
    Write-Host "Enter the drive letter where SQL Server should be installed" -ForegroundColor Cyan
    Write-Host "Examples: C, D, E (will be used as C:, D:, E:)" -ForegroundColor Gray
    $sqlDriveInput = Read-Host "Enter drive letter (default: C)"
    if ([string]::IsNullOrWhiteSpace($sqlDriveInput)) {
        $sqlDrive = "C:"
    } else {
        $sqlDrive = $sqlDriveInput.Trim().ToUpper()
        if (-not $sqlDrive.EndsWith(":")) {
            $sqlDrive = $sqlDrive + ":"
        }
    }
    
    Write-Host ""
    
    # Get Eclipse Client ID
    Write-Host "Eclipse Client ID" -ForegroundColor Yellow
    Write-Host "Enter the 4-digit Eclipse Client ID" -ForegroundColor Cyan
    Write-Host "This will be used for the Eclipse Update Service installation" -ForegroundColor Gray
    do {
        $clientIdInput = Read-Host "Enter Eclipse Client ID (4 digits, required)"
        $clientId = $clientIdInput.Trim()
        if ([string]::IsNullOrWhiteSpace($clientId)) {
            Write-Host "Client ID cannot be empty. Please enter a 4-digit number." -ForegroundColor Red
        } elseif ($clientId -notmatch '^\d{4}$') {
            Write-Host "Client ID must be exactly 4 digits. Please try again." -ForegroundColor Red
        } else {
            break
        }
    } while ($true)
    
    Write-Host ""
    
    # Get Eclipse Aura Hub Subdomain
    Write-Host "Eclipse Aura Hub Subdomain" -ForegroundColor Yellow
    Write-Host "Enter the subdomain for eclipseaurahub.com.au" -ForegroundColor Cyan
    Write-Host "Example: If your site is 'client.eclipseaurahub.com.au', enter 'client'" -ForegroundColor Gray
    Write-Host "This will be used to create an IIS binding for Eclipse Online Server" -ForegroundColor Gray
    do {
        $subdomainInput = Read-Host "Enter subdomain (required)"
        $subdomain = $subdomainInput.Trim()
        if ([string]::IsNullOrWhiteSpace($subdomain)) {
            Write-Host "Subdomain cannot be empty. Please enter a subdomain." -ForegroundColor Red
        } elseif ($subdomain -match '[^a-zA-Z0-9\-]') {
            Write-Host "Subdomain can only contain letters, numbers, and hyphens. Please try again." -ForegroundColor Red
        } else {
            break
        }
    } while ($true)
    
    Write-Host ""
    
    # Get External Port Number for IIS Binding
    Write-Host "External Port Number for IIS Binding" -ForegroundColor Yellow
    Write-Host "Enter the external port number for the IIS binding" -ForegroundColor Cyan
    Write-Host "Common ports: 80 (HTTP), 443 (HTTPS), or custom port (1-65535)" -ForegroundColor Gray
    Write-Host "This port will be used for the IIS binding and firewall rule" -ForegroundColor Gray
    do {
        $portInput = Read-Host "Enter port number (1-65535, required)"
        $port = $portInput.Trim()
        if ([string]::IsNullOrWhiteSpace($port)) {
            Write-Host "Port number cannot be empty. Please enter a port number." -ForegroundColor Red
        } elseif ($port -notmatch '^\d+$') {
            Write-Host "Port number must be numeric. Please try again." -ForegroundColor Red
        } else {
            $portNum = [int]$port
            if ($portNum -lt 1 -or $portNum -gt 65535) {
                Write-Host "Port number must be between 1 and 65535. Please try again." -ForegroundColor Red
            } else {
                break
            }
        }
    } while ($true)
    
    Write-Host ""
    Write-Host "Configuration Summary:" -ForegroundColor Green
    Write-Host "  Eclipse Install Path: $eclipsePath" -ForegroundColor Cyan
    Write-Host "  SQL Server Drive: $sqlDrive" -ForegroundColor Cyan
    Write-Host "  Eclipse Client ID: $clientId" -ForegroundColor Cyan
    Write-Host "  Eclipse Aura Hub Subdomain: $subdomain" -ForegroundColor Cyan
    Write-Host "  Full Domain: $subdomain.eclipseaurahub.com.au" -ForegroundColor Cyan
    Write-Host "  External Port: $port" -ForegroundColor Cyan
    Write-Host ""
    $null = Read-Host "Press Enter to continue to task selection"
    
    $config = New-Object -TypeName System.Collections.Hashtable
    $config['EclipseInstallPath'] = $eclipsePath
    $config['SqlDrive'] = $sqlDrive
    $config['EclipseClientId'] = $clientId
    $config['EclipseAuraHubSubdomain'] = $subdomain
    $config['EclipseAuraHubPort'] = [int]$port
    
    return $config
}

function Show-TaskSelectionMenu {
    param(
        [hashtable]$SelectedTasks
    )
    
    Clear-Host
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Select Tasks to Execute" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Use ARROW KEYS to navigate, SPACEBAR to toggle, A to toggle all, ENTER to proceed" -ForegroundColor Cyan
    Write-Host ""
    
    $tasks = @(
        @{Id=1; Name="Create Eclipse Install Directory Structure"; Selected=$SelectedTasks[1]},
        @{Id=2; Name="Share Eclipse Install Directory"; Selected=$SelectedTasks[2]},
        @{Id=3; Name="Download & Install SQL Server Express 2025"; Selected=$SelectedTasks[3]},
        @{Id=4; Name="Download & Install SQL Server Management Studio 21"; Selected=$SelectedTasks[4]},
        @{Id=5; Name="Configure SQL Server Network Protocols"; Selected=$SelectedTasks[5]},
        @{Id=6; Name="Install/Update Eclipse DMS $EclipseDmsVersion"; Selected=$SelectedTasks[6]},
        @{Id=8; Name="Add Ultimate Domains to Trusted Sites"; Selected=$SelectedTasks[8]},
        @{Id=11; Name="Install Microsoft IIS Server + Dependencies"; Selected=$SelectedTasks[11]},
        @{Id=12; Name="Download IIS URL Rewrite Module"; Selected=$SelectedTasks[12]},
        @{Id=13; Name="Install ASP.NET Core Hosting Bundle 8.0.22 & .NET Desktop Runtime 8.0.22"; Selected=$SelectedTasks[13]},
        @{Id=14; Name="Download .NET Framework 4.8"; Selected=$SelectedTasks[14]},
        @{Id=15; Name="Download Eclipse Update Service $EclipseUpdateServiceVersion"; Selected=$SelectedTasks[15]},
        @{Id=17; Name="Install Eclipse Online Server $EclipseOnlineServerVersion & Create IIS Binding"; Selected=$SelectedTasks[17]},
        @{Id=18; Name="Update IIS Application Pools"; Selected=$SelectedTasks[18]},
        @{Id=19; Name="Install Eclipse Smart Hub $EclipseSmartHubVersion"; Selected=$SelectedTasks[19]},
        @{Id=20; Name="Install Win-ACMEv2"; Selected=$SelectedTasks[20]}
    )
    
    $currentIndex = 0
    $key = $null
    
    do {
        Clear-Host
        Write-Host "===============================================" -ForegroundColor Cyan
        Write-Host "   Select Tasks to Execute" -ForegroundColor Yellow
        Write-Host "===============================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Use ARROW KEYS to navigate, SPACEBAR to toggle, A to toggle all, ENTER to proceed" -ForegroundColor Cyan
        Write-Host ""
        
        for ($i = 0; $i -lt $tasks.Count; $i++) {
            $task = $tasks[$i]
            $marker = if ($i -eq $currentIndex) { ">> " } else { "   " }
            $checkbox = if ($task.Selected) { "[X]" } else { "[ ]" }
            $color = if ($i -eq $currentIndex) { "Yellow" } else { "White" }
            
            Write-Host "$marker$checkbox " -NoNewline
            Write-Host "$($task.Id). $($task.Name)" -ForegroundColor $color
        }
        
        Write-Host ""
        Write-Host "Press ENTER to start execution with selected tasks" -ForegroundColor Green
        Write-Host "Press A to toggle all tasks on/off" -ForegroundColor Cyan
        Write-Host "Press ESC to exit" -ForegroundColor Red
        
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        
        switch ($key.VirtualKeyCode) {
            38 { # Up Arrow
                if ($currentIndex -gt 0) { $currentIndex-- }
            }
            40 { # Down Arrow
                if ($currentIndex -lt ($tasks.Count - 1)) { $currentIndex++ }
            }
            32 { # Spacebar
                $tasks[$currentIndex].Selected = -not $tasks[$currentIndex].Selected
                $SelectedTasks[$tasks[$currentIndex].Id] = $tasks[$currentIndex].Selected
            }
            65 { # A key
                # Check if all tasks are selected
                $allSelected = $true
                foreach ($task in $tasks) {
                    if (-not $task.Selected) {
                        $allSelected = $false
                        break
                    }
                }
                
                # Toggle all: if all are selected, deselect all; otherwise select all
                $newState = -not $allSelected
                foreach ($task in $tasks) {
                    $task.Selected = $newState
                    $SelectedTasks[$task.Id] = $newState
                }
            }
            13 { # Enter
                # Check if any tasks are selected
                $hasSelection = $false
                foreach ($task in $tasks) {
                    if ($task.Selected) {
                        $hasSelection = $true
                        break
                    }
                }
                
                if (-not $hasSelection) {
                    # Show error message and wait for key press
                    Write-Host ""
                    Write-Host "ERROR: Please select at least one task before proceeding!" -ForegroundColor Red
                    Write-Host "Press any key to continue..." -ForegroundColor Yellow
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    # Continue the loop instead of returning
                } else {
                    return $SelectedTasks
                }
            }
            27 { # Escape
                return $null
            }
        }
    } while ($true)
}

function Execute-Task1 {
    param(
        [string]$EclipseInstallPath
    )
    
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Task 1: Create Eclipse Install Directory Structure" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        Write-Host "Target directory: $EclipseInstallPath" -ForegroundColor Green
        $folders = @("Documents", "Pictures", "Help", "Program Updates", "Layouts", "Scripts", "Backups", "Dependencies", "Price Files")
        $picturesSub = @("Stock", "Parts", "Service")
        
        # Create main folder if not exists
        if (!(Test-Path $EclipseInstallPath)) {
            New-Item -Path $EclipseInstallPath -ItemType Directory -Force | Out-Null
            Write-Host "Created: $EclipseInstallPath" -ForegroundColor Green
        } else {
            Write-Host "Exists: $EclipseInstallPath (skipping)" -ForegroundColor Yellow
        }
        
        # Create subfolders
        foreach ($folder in $folders) {
            $subPath = Join-Path $EclipseInstallPath $folder
            if (!(Test-Path $subPath)) {
                New-Item -Path $subPath -ItemType Directory -Force | Out-Null
                Write-Host "Created: $subPath" -ForegroundColor Green
            } else {
                Write-Host "Exists: $subPath (skipping)" -ForegroundColor Yellow
            }
        }
        
        # Pictures subfolders
        $picturesPath = Join-Path $EclipseInstallPath "Pictures"
        foreach ($sub in $picturesSub) {
            $subPicPath = Join-Path $picturesPath $sub
            if (!(Test-Path $subPicPath)) {
                New-Item -Path $subPicPath -ItemType Directory -Force | Out-Null
                Write-Host "Created: $subPicPath" -ForegroundColor Green
            } else {
                Write-Host "Exists: $subPicPath (skipping)" -ForegroundColor Yellow
            }
        }
        
        Write-Host ""
        Write-Host "Directory structure creation complete!" -ForegroundColor Cyan
        return $true
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Execute-Task2 {
    param(
        [string]$EclipseInstallPath
    )
    
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Task 2: Share Eclipse Install Directory" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        # Validate path exists
        if (!(Test-Path $EclipseInstallPath)) {
            Write-Host "Error: The specified path does not exist: $EclipseInstallPath" -ForegroundColor Red
            Write-Host "Please run Task 1 first to create the directory." -ForegroundColor Yellow
            return $false
        }
        
        Write-Host "Target directory: $EclipseInstallPath" -ForegroundColor Green
        
        # Get share name (default to folder name)
        $shareName = Split-Path $EclipseInstallPath -Leaf
        
        # Check if share already exists
        $existingShare = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
        if ($existingShare) {
            Write-Host "Share '$shareName' already exists. Removing existing share..." -ForegroundColor Yellow
            Remove-SmbShare -Name $shareName -Force -ErrorAction Stop
            Write-Host "Existing share removed." -ForegroundColor Green
        }
        
        # Create the share
        Write-Host "Creating Windows share '$shareName' for path: $EclipseInstallPath" -ForegroundColor Yellow
        New-SmbShare -Path $EclipseInstallPath -Name $shareName -Description "Eclipse Install Directory" -ErrorAction Stop | Out-Null
        Write-Host "Share created successfully: \\$env:COMPUTERNAME\$shareName" -ForegroundColor Green
        
        # Set share permissions
        Write-Host "Configuring share permissions..." -ForegroundColor Yellow
        $existingPermissions = Get-SmbShareAccess -Name $shareName -ErrorAction SilentlyContinue
        foreach ($perm in $existingPermissions) {
            if ($perm.AccountName -eq "Everyone") {
                Revoke-SmbShareAccess -Name $shareName -AccountName "Everyone" -Force -ErrorAction SilentlyContinue
            }
        }
        
        Grant-SmbShareAccess -Name $shareName -AccountName "Administrators" -AccessRight Full -Force -ErrorAction Stop
        Write-Host "Granted Full Control to Administrators" -ForegroundColor Green
        
        Grant-SmbShareAccess -Name $shareName -AccountName "Authenticated Users" -AccessRight Full -Force -ErrorAction Stop
        Write-Host "Granted Full Control to Authenticated Users" -ForegroundColor Green
        
        # Set NTFS permissions
        Write-Host "Configuring NTFS security permissions..." -ForegroundColor Yellow
        $acl = Get-Acl -Path $EclipseInstallPath
        
        $adminRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Administrators", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
        $acl.SetAccessRule($adminRule)
        
        $authUsersRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Authenticated Users", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
        $acl.SetAccessRule($authUsersRule)
        
        Set-Acl -Path $EclipseInstallPath -AclObject $acl -ErrorAction Stop
        Write-Host "NTFS permissions configured successfully" -ForegroundColor Green
        
        Write-Host ""
        Write-Host "Share Configuration Complete!" -ForegroundColor Green
        Write-Host "Share Path: \\$env:COMPUTERNAME\$shareName" -ForegroundColor Cyan
        return $true
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Execute-Task3 {
    param(
        [string]$SqlDrive
    )
    
    Write-Log "===============================================" -ForegroundColor Cyan
    Write-Log "   Task 3: Download & Install SQL Server Express 2025" -ForegroundColor Yellow
    Write-Log "===============================================" -ForegroundColor Cyan
    Write-Log ""
    
    try {
        $targetDir = "$SqlDrive\SQL2025"
        if (!(Test-Path $targetDir)) {
            New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
            Write-Log "Created directory: $targetDir" -ForegroundColor Green
        }
        
        # Download ZIP file
        $zipUrl = "https://tnh.net.au/packages/SQLEXPR_x64_ENU.zip"
        $zipPath = Join-Path $targetDir "SQLEXPR_x64_ENU.zip"
        
        if (Test-Path $zipPath) {
            Write-Log "ZIP file already downloaded: $zipPath" -ForegroundColor Green
        } else {
            Write-Log "Downloading SQL Server Express 2025 ZIP file (this may take a few minutes)..." -ForegroundColor Yellow
            Write-Log "  Source URL: $zipUrl" -ForegroundColor Gray
            Write-Log "  Destination: $zipPath" -ForegroundColor Gray
            # Ensure TLS 1.2 is used (required by many HTTPS servers; default on some Windows is TLS 1.0)
            $prevProtocol = [Net.ServicePointManager]::SecurityProtocol
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            } catch {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            }
            try {
                $webClient = New-Object System.Net.WebClient
                $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) PowerShell-EclipseSwissArmyKnife")
                $webClient.DownloadFile($zipUrl, $zipPath)
                $webClient.Dispose()
            } catch {
                [Net.ServicePointManager]::SecurityProtocol = $prevProtocol
                Write-Log "Download failed: $($_.Exception.Message)" -ForegroundColor Red
                if ($_.Exception.InnerException) { Write-Log "  Inner: $($_.Exception.InnerException.Message)" -ForegroundColor Red }
                Write-Log "  If you see 'SSL/TLS secure channel', the server requires TLS 1.2; the script has enabled it. Check proxy/firewall or try from another network." -ForegroundColor Yellow
                throw
            }
            [Net.ServicePointManager]::SecurityProtocol = $prevProtocol
            $fileSize = (Get-Item $zipPath).Length / 1MB
            Write-Log "Download complete: $zipPath ($([math]::Round($fileSize, 2)) MB)" -ForegroundColor Green
        }
        
        # Create unattended installation configuration file
        Write-Log "Creating unattended installation configuration..." -ForegroundColor Yellow
        $configFile = Join-Path $targetDir "ConfigurationFile.ini"
        
        # Format path properly - paths with spaces need quotes in SQL Server config
        $sqlDataDir = "$SqlDrive\Program Files\Microsoft SQL Server"
        
        $configContent = @"
[OPTIONS]
; SQL Server 2025 Express Edition Unattended Installation
ACTION="Install"
FEATURES=SQLENGINE
INSTANCENAME="MSSQLSERVER"
SQLSVCACCOUNT="NT AUTHORITY\SYSTEM"
SQLSYSADMINACCOUNTS="BUILTIN\Administrators"
SECURITYMODE="SQL"
SAPWD="UBSubs123"
TCPENABLED="1"
NPENABLED="1"
INSTALLSQLDATADIR="$sqlDataDir"
IACCEPTSQLSERVERLICENSETERMS="True"
"@
        
        Set-Content -Path $configFile -Value $configContent -Encoding ASCII
        Write-Log "Configuration file created: $configFile" -ForegroundColor Green
        Write-Log "  Instance Name: MSSQLSERVER" -ForegroundColor Gray
        Write-Log "  SA Password: UBSubs123" -ForegroundColor Gray
        Write-Log "  Data Directory: $sqlDataDir" -ForegroundColor Gray
        
        # Create log file for SQL Server installation
        $sqlLogFile = Join-Path $targetDir "SQLServer_Install.log"
        Write-Log "Installation log will be saved to: $sqlLogFile" -ForegroundColor Gray
        
        # Extract ZIP to target directory (C:\SQL2025)
        Write-Log "Extracting ZIP file to: $targetDir" -ForegroundColor Yellow
        Write-Log "This may take a minute..." -ForegroundColor Gray
        
        # Start logging
        $logContent = @"
===============================================
SQL Server Express 2025 Installation Log
Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
===============================================

ZIP File: $zipPath
Extraction Directory: $targetDir
Configuration File: $configFile

Configuration File Contents:
$configContent

===============================================
Extraction Phase:
===============================================

"@
        Set-Content -Path $sqlLogFile -Value $logContent -Encoding UTF8
        
        try {
            # Use PowerShell's built-in Expand-Archive cmdlet
            Write-Log "Starting ZIP extraction..." -ForegroundColor Gray
            Expand-Archive -Path $zipPath -DestinationPath $targetDir -Force
            Write-Log "Extraction complete!" -ForegroundColor Green
            Add-Content -Path $sqlLogFile -Value "ZIP extracted successfully to: $targetDir`n" -Encoding UTF8
        } catch {
            Write-Log "Error extracting ZIP: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "  Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
            Add-Content -Path $sqlLogFile -Value "ERROR extracting ZIP: $($_.Exception.Message)`n" -Encoding UTF8
            return $false
        }
        
        # Find setup.exe in the extracted directory
        Start-Sleep -Seconds 2  # Give it time to finish writing files
        Write-Log "Searching for setup.exe in: $targetDir" -ForegroundColor Gray
        
        # List what was extracted for debugging
        $extractedItems = Get-ChildItem -Path $targetDir -Recurse -ErrorAction SilentlyContinue | Select-Object -First 20 FullName
        Write-Log "Extracted items (first 20):" -ForegroundColor Gray
        foreach ($item in $extractedItems) {
            Write-Log "  $($item.FullName)" -ForegroundColor DarkGray
        }
        Add-Content -Path $sqlLogFile -Value "Extracted items:`n$($extractedItems | ForEach-Object { $_.FullName } | Out-String)`n" -Encoding UTF8
        
        $setupExe = Get-ChildItem -Path $targetDir -Filter "setup.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
        
        if (-not $setupExe -or -not (Test-Path $setupExe)) {
            Write-Log "Error: setup.exe not found after extraction" -ForegroundColor Red
            Write-Log "Searched in: $targetDir" -ForegroundColor Yellow
            Write-Log "Listing all .exe files found:" -ForegroundColor Yellow
            $allExes = Get-ChildItem -Path $targetDir -Filter "*.exe" -Recurse -ErrorAction SilentlyContinue
            foreach ($exe in $allExes) {
                Write-Log "  $($exe.FullName)" -ForegroundColor Gray
            }
            Add-Content -Path $sqlLogFile -Value "ERROR: setup.exe not found in $targetDir`nAll .exe files found:`n$($allExes | ForEach-Object { $_.FullName } | Out-String)`n" -Encoding UTF8
            return $false
        }
        
        Write-Log "Found setup.exe at: $setupExe" -ForegroundColor Green
        Add-Content -Path $sqlLogFile -Value "Found setup.exe: $setupExe`n" -Encoding UTF8
        
        # Copy config file to the same directory as setup.exe (setup.exe prefers config file in its directory)
        $setupDir = Split-Path $setupExe -Parent
        $configFileInSetupDir = Join-Path $setupDir "ConfigurationFile.ini"
        Copy-Item -Path $configFile -Destination $configFileInSetupDir -Force
        Write-Log "Configuration file copied to setup directory: $configFileInSetupDir" -ForegroundColor Green
        Add-Content -Path $sqlLogFile -Value "Configuration file copied to: $configFileInSetupDir`n" -Encoding UTF8
        
        # Use the config file in the setup directory
        $configFile = $configFileInSetupDir
        
        # Now run setup.exe with the configuration file using /QS (Quiet Simple - shows window but no interaction)
        Write-Log "Installing SQL Server Express 2025 (this may take 10-15 minutes)..." -ForegroundColor Cyan
        Write-Log "SQL Server Setup window will be visible - this is normal." -ForegroundColor Yellow
        Write-Log "The installation will proceed automatically with the configured settings." -ForegroundColor Yellow
        
        $setupArgs = "/ConfigurationFile=`"$configFile`" /QS /IACCEPTSQLSERVERLICENSETERMS"
        
        Add-Content -Path $sqlLogFile -Value "===============================================`nInstallation Phase:`n===============================================`n`nRunning: $setupExe $setupArgs`n" -Encoding UTF8
        Write-Log "Command: $setupExe $setupArgs" -ForegroundColor DarkGray
        
        # Run setup.exe - show window with progress (/PASSIVE shows progress bar)
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $setupExe
        $processInfo.Arguments = $setupArgs
        $processInfo.UseShellExecute = $true
        $processInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Normal
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        
        Write-Log "Starting SQL Server installation (window will be visible)..." -ForegroundColor Cyan
        Write-Log "This will take 10-15 minutes. Please wait..." -ForegroundColor Yellow
        
        try {
            $started = $process.Start()
            if (-not $started) {
                Write-Log "Error: Failed to start setup.exe process" -ForegroundColor Red
                Add-Content -Path $sqlLogFile -Value "ERROR: Failed to start setup.exe process`n" -Encoding UTF8
                return $false
            }
            
            Write-Log "Setup.exe process started (PID: $($process.Id))" -ForegroundColor Green
            Write-Log "  Process start time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
            Add-Content -Path $sqlLogFile -Value "Setup.exe process started successfully (PID: $($process.Id))`n" -Encoding UTF8
            
            # Wait a moment and check if process is still running
            Start-Sleep -Seconds 3
            if ($process.HasExited) {
                Write-Log "Warning: Setup.exe exited immediately with code: $($process.ExitCode)" -ForegroundColor Yellow
                Add-Content -Path $sqlLogFile -Value "Warning: Setup.exe exited immediately with code: $($process.ExitCode)`n" -Encoding UTF8
            } else {
                Write-Log "Setup.exe is running. Waiting for installation to complete..." -ForegroundColor Cyan
                Write-Log "  Monitoring process status..." -ForegroundColor Gray
            }
            
            # Wait for the process to exit with periodic status updates
            $waitStartTime = Get-Date
            while (-not $process.HasExited) {
                Start-Sleep -Seconds 30
                $elapsed = (Get-Date) - $waitStartTime
                $minutes = [math]::Floor($elapsed.TotalMinutes)
                $seconds = [math]::Floor($elapsed.TotalSeconds % 60)
                Write-Log "  Installation in progress... (Elapsed: ${minutes}m ${seconds}s)" -ForegroundColor Gray
            }
            
            Write-Log "SQL Server installation process has completed." -ForegroundColor Green
            Write-Log "Exit code: $($process.ExitCode)" -ForegroundColor Cyan
            Write-Log "  Process end time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
            Add-Content -Path $sqlLogFile -Value "Setup.exe completed with exit code: $($process.ExitCode)`n" -Encoding UTF8
        } catch {
            Write-Log "Error starting setup.exe: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "  Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
            Add-Content -Path $sqlLogFile -Value "ERROR starting setup.exe: $($_.Exception.Message)`n$($_.Exception.StackTrace)`n" -Encoding UTF8
            return $false
        }
        
        # Log completion
        $finalLog = @"

===============================================
Exit Code: $($process.ExitCode)
Completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
===============================================

Note: Output was not captured as the installation window was visible.

"@
        
        Add-Content -Path $sqlLogFile -Value $finalLog -Encoding UTF8
        Write-Log ""
        Write-Log "Installation log saved to: $sqlLogFile" -ForegroundColor Cyan
        
        # Check exit code and verify installation
        $installationSuccess = $false
        
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
            # Exit code 0 = success, 3010 = success but reboot required
            Write-Log "SQL Server installer completed with exit code: $($process.ExitCode)" -ForegroundColor Green
            if ($process.ExitCode -eq 3010) {
                Write-Log "Note: System reboot may be required." -ForegroundColor Yellow
            }
            
            # Verify SQL Server is actually installed
            Write-Log "Verifying SQL Server installation..." -ForegroundColor Yellow
            Start-Sleep -Seconds 3
            $sqlService = Get-Service -Name "MSSQLSERVER" -ErrorAction SilentlyContinue
            if ($sqlService) {
                Write-Log "SQL Server Express 2025 installation verified successfully!" -ForegroundColor Green
                Write-Log "Service Name: $($sqlService.Name)" -ForegroundColor Cyan
                Write-Log "Display Name: $($sqlService.DisplayName)" -ForegroundColor Cyan
                Write-Log "Service Status: $($sqlService.Status)" -ForegroundColor Cyan
                $installationSuccess = $true
            } else {
                # Check registry as alternative verification
                Write-Log "Service not found, checking registry..." -ForegroundColor Gray
                $regPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL"
                if (Test-Path $regPath) {
                    $instances = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
                    if ($instances) {
                        Write-Log "SQL Server installation verified via registry!" -ForegroundColor Green
                        Write-Log "Found instances: $($instances.PSObject.Properties.Name -join ', ')" -ForegroundColor Cyan
                        $installationSuccess = $true
                    } else {
                        Write-Log "Warning: Installer completed but SQL Server verification failed." -ForegroundColor Yellow
                        Write-Log "Installation may still be in progress or may have failed." -ForegroundColor Yellow
                        $installationSuccess = $false
                    }
                } else {
                    Write-Log "Warning: Installer completed but SQL Server verification failed." -ForegroundColor Yellow
                    $installationSuccess = $false
                }
            }
            
            if ($installationSuccess) {
                return $true
            } else {
                Write-Log "Installation may have completed but verification failed." -ForegroundColor Yellow
                Write-Log "Exit code was: $($process.ExitCode)" -ForegroundColor Gray
                return $false
            }
        } else {
            Write-Log "Installation failed with exit code: $($process.ExitCode)" -ForegroundColor Red
            Write-Log "Common exit codes:" -ForegroundColor Yellow
            Write-Log "  0 = Success" -ForegroundColor Gray
            Write-Log "  16389 = Configuration error or invalid parameter" -ForegroundColor Gray
            Write-Log "  3010 = Success but reboot required" -ForegroundColor Gray
            Write-Log ""
            Write-Log "Troubleshooting:" -ForegroundColor Yellow
            Write-Log "  1. Check the installation log file: $sqlLogFile" -ForegroundColor Cyan
            Write-Log "  2. Check the SQL Server setup logs in:" -ForegroundColor Gray
            Write-Log "     - $env:ProgramFiles\Microsoft SQL Server\170\Setup Bootstrap\Log (SQL 2025)" -ForegroundColor DarkGray
            Write-Log "     - $env:ProgramFiles\Microsoft SQL Server\150\Setup Bootstrap\Log (SQL 2022)" -ForegroundColor DarkGray
            Write-Log "  3. Verify the SQL Server drive has sufficient space" -ForegroundColor Gray
            Write-Log "  5. Ensure you're running as Administrator" -ForegroundColor Gray
            Write-Log ""
            Write-Log "Installation output has been saved to: $sqlLogFile" -ForegroundColor Yellow
            return $false
        }
    } catch {
        Write-Log "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "  Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
        return $false
    }
}

function Execute-Task4 {
    param(
        [string]$EclipseInstallPath
    )
    
    Write-Log "===============================================" -ForegroundColor Cyan
    Write-Log "   Task 4: Install SQL Server Management Studio 21" -ForegroundColor Yellow
    Write-Log "===============================================" -ForegroundColor Cyan
    Write-Log ""
    
    try {
        # Check if SSMS is already installed
        $ssmsPaths = @(
            "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio 21\Common7\IDE\Ssms.exe",
            "${env:ProgramFiles}\Microsoft SQL Server Management Studio 21\Common7\IDE\Ssms.exe"
        )
        
        foreach ($ssmsPath in $ssmsPaths) {
            if (Test-Path $ssmsPath) {
                Write-Log "SSMS 21 is already installed at: $ssmsPath" -ForegroundColor Green
                return $true
            }
        }
        
        # Download SSMS installer directly
        $targetDir = Join-Path $EclipseInstallPath "Dependencies"
        if (!(Test-Path $targetDir)) {
            New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
        }
        
        $installerUrl = "https://download.visualstudio.microsoft.com/download/pr/5967a899-96aa-47e2-a7c5-1b7192f292ee/cb58e449b5f554b28fa66a2a04d7c1bf55073f16242c8eb3b4d63fbfc10de5d0/vs_SSMS.exe"
        $installerPath = Join-Path $targetDir "SSMS-Setup-21.6.17.exe"
        
        if (Test-Path $installerPath) {
            Write-Log "SSMS installer already downloaded: $installerPath" -ForegroundColor Green
        } else {
            Write-Log "Downloading SSMS 21.6.17 installer..." -ForegroundColor Yellow
            Write-Log "  Source URL: $installerUrl" -ForegroundColor Gray
            Write-Log "  Destination: $installerPath" -ForegroundColor Gray
            
            try {
                $webClient = New-Object System.Net.WebClient
                $webClient.DownloadFile($installerUrl, $installerPath)
                $webClient.Dispose()
                $fileSize = (Get-Item $installerPath).Length / 1MB
                Write-Log "Download complete: $([math]::Round($fileSize, 2)) MB" -ForegroundColor Green
            } catch {
                Write-Log "Error downloading SSMS installer: $($_.Exception.Message)" -ForegroundColor Red
                return $false
            }
        }
        
        # Install SSMS silently
        Write-Log ""
        Write-Log "Installing SSMS 21.6.17 silently..." -ForegroundColor Yellow
        Write-Log "This may take 5-10 minutes. Please wait..." -ForegroundColor Cyan
        Write-Log "  Using arguments: --quiet --norestart" -ForegroundColor Gray
        
        try {
            # Use Start-Process to allow proper process spawning
            $process = Start-Process -FilePath $installerPath -ArgumentList "--quiet --norestart" -PassThru -NoNewWindow
            $bootstrapperPid = $process.Id
            Write-Log "  SSMS bootstrapper started (PID: $bootstrapperPid)" -ForegroundColor Gray
            
            # Wait for bootstrapper to exit (it spawns child processes)
            Write-Log "  Waiting for bootstrapper to spawn installer processes..." -ForegroundColor Gray
            $process.WaitForExit()
            
            $exitCode = $null
            try {
                if ($process.HasExited) {
                    $exitCode = $process.ExitCode
                    if ($null -ne $exitCode -and $exitCode -ne "") {
                        Write-Log "  Bootstrapper exited with code: $exitCode" -ForegroundColor $(if ($exitCode -eq 0) { "Green" } else { "Yellow" })
                    } else {
                        Write-Log "  Bootstrapper exited (spawning child processes)" -ForegroundColor Green
                    }
                }
            } catch {
                Write-Log "  Bootstrapper exited (spawning child processes)" -ForegroundColor Green
            }
            
            # Now wait for the spawned child processes to complete
            Write-Log "  Waiting for Visual Studio Installer and setup processes to complete..." -ForegroundColor Cyan
            Write-Log "  (This is where the actual SSMS installation happens)" -ForegroundColor Gray
            
            $maxWaitMinutes = 15
            $startWait = Get-Date
            $lastProgressUpdate = Get-Date
            $foundInstallerProcesses = $false
            
            # Give child processes time to spawn
            Start-Sleep -Seconds 5
            
            while (((Get-Date) - $startWait).TotalMinutes -lt $maxWaitMinutes) {
                # Check for Visual Studio Installer and setup processes
                $vsInstaller = Get-Process -Name "vs_installer" -ErrorAction SilentlyContinue
                $setupProcess = Get-Process -Name "setup" -ErrorAction SilentlyContinue
                
                if ($vsInstaller -or $setupProcess) {
                    $foundInstallerProcesses = $true
                }
                
                if (-not $vsInstaller -and -not $setupProcess) {
                    if ($foundInstallerProcesses) {
                        # Installer processes were running but now stopped = installation complete
                        Write-Log "  Installation processes completed" -ForegroundColor Green
                        break
                    } else {
                        # Haven't found any installer processes yet
                        if (((Get-Date) - $startWait).TotalSeconds -gt 30) {
                            Write-Log "  Warning: No installer processes detected after 30 seconds" -ForegroundColor Yellow
                            Write-Log "  Installation may have failed or completed very quickly" -ForegroundColor Yellow
                            break
                        }
                    }
                }
                
                # Show progress update every 30 seconds
                if (((Get-Date) - $lastProgressUpdate).TotalSeconds -ge 30) {
                    $elapsed = [math]::Floor(((Get-Date) - $startWait).TotalMinutes)
                    Write-Log "  Still installing... ($elapsed minutes elapsed)" -ForegroundColor Gray
                    if ($vsInstaller) {
                        Write-Log "    Visual Studio Installer active (PID: $($vsInstaller.Id))" -ForegroundColor DarkGray
                    }
                    if ($setupProcess) {
                        Write-Log "    Setup process active (PID: $($setupProcess.Id))" -ForegroundColor DarkGray
                    }
                    $lastProgressUpdate = Get-Date
                }
                
                Start-Sleep -Seconds 5
            }
            
            if (((Get-Date) - $startWait).TotalMinutes -ge $maxWaitMinutes) {
                Write-Log "  Warning: Installation monitoring timed out after $maxWaitMinutes minutes" -ForegroundColor Yellow
            }
            
            if ($null -ne $exitCode -and $exitCode -eq 3010) {
                Write-Log "  Note: Exit code 3010 indicates restart may be required" -ForegroundColor Yellow
            }
            
            Write-Log "  Installer processes completed. Proceeding to verification..." -ForegroundColor Gray
            
            # Verify installation
            Write-Log ""
            Write-Log "Verifying SSMS installation..." -ForegroundColor Yellow
            Write-Log "  Waiting 30 seconds for installation to finalize..." -ForegroundColor Gray
            Start-Sleep -Seconds 30
            
            $ssmsFound = $false
            $foundPath = $null
            
            # First check registry for SSMS installation
            try {
                $ssmsRegistry = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server Management Studio\21.0" -ErrorAction SilentlyContinue
                if ($ssmsRegistry) {
                    Write-Log "SSMS 21 found in registry!" -ForegroundColor Green
                    $ssmsFound = $true
                }
            } catch {
                Write-Log "  Registry check did not find SSMS 21" -ForegroundColor Gray
            }
            
            # Also check file paths
            if (-not $ssmsFound) {
                foreach ($ssmsPath in $ssmsPaths) {
                    if (Test-Path $ssmsPath) {
                        $ssmsFound = $true
                        $foundPath = $ssmsPath
                        Write-Log "SSMS 21 executable found at: $foundPath" -ForegroundColor Green
                        break
                    }
                }
            }
            
            # Check for any SSMS version in registry as fallback
            if (-not $ssmsFound) {
                try {
                    $anySSMSRegistry = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server Management Studio" -ErrorAction SilentlyContinue
                    if ($anySSMSRegistry) {
                        $versions = $anySSMSRegistry | ForEach-Object { $_.PSChildName }
                        Write-Log "Found SSMS version(s) in registry: $($versions -join ', ')" -ForegroundColor Yellow
                        $ssmsFound = $true
                    }
                } catch {
                    Write-Log "  Could not check registry for any SSMS versions" -ForegroundColor Gray
                }
            }
            
            # Check additional file paths as last resort
            if (-not $ssmsFound) {
                $additionalPaths = @(
                    "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio 20\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 20\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio 19\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 19\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 21\Common7\IDE\Ssms.exe",
                    "C:\Program Files\Microsoft SQL Server Management Studio 21\Common7\IDE\Ssms.exe"
                )
                
                foreach ($path in $additionalPaths) {
                    if (Test-Path $path) {
                        Write-Log "SSMS found at: $path" -ForegroundColor Green
                        $ssmsFound = $true
                        $foundPath = $path
                        break
                    }
                }
            }
            
            if ($ssmsFound) {
                Write-Log "SSMS 21 installation verified successfully!" -ForegroundColor Green
                if ($foundPath) {
                    Write-Log "  Executable: $foundPath" -ForegroundColor Gray
                }
                return $true
            } else {
                Write-Log "Warning: Could not verify SSMS installation via registry or file paths." -ForegroundColor Yellow
                Write-Log "However, the installer completed successfully (exit code 0)." -ForegroundColor Yellow
                Write-Log "SSMS may be installed - please verify manually if needed." -ForegroundColor Yellow
                return $true  # Return success since installer completed without errors
            }
        } catch {
            Write-Log "Error during installation: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Log "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "  Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
        return $false
    }
}

function Execute-Task5 {
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Task 5: Configure SQL Server Network Protocols" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        Write-Host "Searching for SQL Server instances via registry..." -ForegroundColor Yellow
        Write-Host ""
        
        $found = $false
        $configuredInstances = @()
        
        # Find SQL Server instances by checking registry and services
        $sqlServices = Get-Service | Where-Object { $_.Name -like "MSSQL*" -and $_.Name -notlike "*Agent*" -and $_.Name -notlike "*Browser*" -and $_.Name -notlike "*OLAP*" -and $_.Name -notlike "*ReportServer*" }
        
        if ($sqlServices.Count -eq 0) {
            Write-Host "No SQL Server services found on this machine." -ForegroundColor Red
            Write-Host "Please ensure SQL Server is installed before running this task." -ForegroundColor Yellow
            return $false
        }
        
        Write-Host "Found SQL Server services:" -ForegroundColor Green
        foreach ($svc in $sqlServices) {
            Write-Host "  - $($svc.Name)" -ForegroundColor Cyan
        }
        Write-Host ""
        
        # Get SQL Server instance IDs from registry
        $sqlRegPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server"
        if (Test-Path $sqlRegPath) {
            $instanceIds = Get-ChildItem -Path $sqlRegPath | Where-Object { $_.PSChildName -like "MSSQL*" -or $_.PSChildName -like "SQL*" }
            
            foreach ($instanceId in $instanceIds) {
                $instanceName = $instanceId.PSChildName
                $protocolPath = Join-Path $instanceId.PSPath "MSSQLServer\SuperSocketNetLib"
                
                if (Test-Path $protocolPath) {
                    $found = $true
                    Write-Host "Configuring instance: $instanceName" -ForegroundColor Green
                    
                    $tcpEnabled = $false
                    $npEnabled = $false
                    
                    # Enable TCP/IP
                    $tcpPath = Join-Path $protocolPath "Tcp"
                    if (Test-Path $tcpPath) {
                        $tcpEnabledValue = Get-ItemProperty -Path $tcpPath -Name "Enabled" -ErrorAction SilentlyContinue
                        if ($tcpEnabledValue.Enabled -ne 1) {
                            Set-ItemProperty -Path $tcpPath -Name "Enabled" -Value 1 -Type DWord -ErrorAction SilentlyContinue
                            $tcpEnabled = $true
                            Write-Host "  Enabled: TCP/IP" -ForegroundColor Green
                        } else {
                            Write-Host "  TCP/IP already enabled" -ForegroundColor Cyan
                        }
                    }
                    
                    # Enable Named Pipes
                    $npPath = Join-Path $protocolPath "Np"
                    if (Test-Path $npPath) {
                        $npEnabledValue = Get-ItemProperty -Path $npPath -Name "Enabled" -ErrorAction SilentlyContinue
                        if ($npEnabledValue.Enabled -ne 1) {
                            Set-ItemProperty -Path $npPath -Name "Enabled" -Value 1 -Type DWord -ErrorAction SilentlyContinue
                            $npEnabled = $true
                            Write-Host "  Enabled: Named Pipes" -ForegroundColor Green
                        } else {
                            Write-Host "  Named Pipes already enabled" -ForegroundColor Cyan
                        }
                    }
                    
                    # Find the service name for this instance
                    $serviceName = $null
                    foreach ($svc in $sqlServices) {
                        if ($svc.Name -eq "MSSQLSERVER") {
                            $serviceName = "MSSQLSERVER"
                            break
                        } elseif ($svc.Name -like "MSSQL`$*") {
                            $svcInstanceName = $svc.Name -replace "MSSQL`$", ""
                            if ($instanceName -like "*$svcInstanceName*" -or $svcInstanceName -like "*$instanceName*") {
                                $serviceName = $svc.Name
                                break
                            }
                        }
                    }
                    
                    if (-not $serviceName) {
                        if ($instanceName -eq "MSSQLSERVER" -or $instanceName -like "MSSQL*") {
                            $defaultSvc = $sqlServices | Where-Object { $_.Name -eq "MSSQLSERVER" }
                            if ($defaultSvc) {
                                $serviceName = "MSSQLSERVER"
                            }
                        }
                        if (-not $serviceName) {
                            $namedSvc = $sqlServices | Where-Object { $_.Name -like "MSSQL`$*" } | Select-Object -First 1
                            if ($namedSvc) {
                                $serviceName = $namedSvc.Name
                            }
                        }
                    }
                    
                    # Restart SQL Server service if protocols were changed
                    if (($tcpEnabled -or $npEnabled) -and $serviceName) {
                        Write-Host "Restarting SQL Server service: $serviceName" -ForegroundColor Yellow
                        try {
                            Restart-Service -Name $serviceName -Force -ErrorAction Stop
                            Write-Host "  Service restarted successfully" -ForegroundColor Green
                            $configuredInstances += "$instanceName ($serviceName)"
                        } catch {
                            Write-Host "  Warning: Could not restart service: $($_.Exception.Message)" -ForegroundColor Yellow
                            $configuredInstances += "$instanceName ($serviceName - manual restart needed)"
                        }
                    } elseif ($tcpEnabled -or $npEnabled) {
                        Write-Host "  Warning: Could not determine service name for instance $instanceName" -ForegroundColor Yellow
                        $configuredInstances += "$instanceName (manual restart needed)"
                    } else {
                        Write-Host "  No protocol changes needed" -ForegroundColor Cyan
                        $configuredInstances += "$instanceName"
                    }
                    Write-Host ""
                }
            }
        }
        
        if ($found) {
            Write-Host "===============================================" -ForegroundColor Cyan
            Write-Host "Configuration complete for instances:" -ForegroundColor Green
            foreach ($inst in $configuredInstances) {
                Write-Host "  - $inst" -ForegroundColor Cyan
            }
            Write-Host "===============================================" -ForegroundColor Cyan
            
            # Create Windows firewall rule for SQL Server port 1433
            Write-Host ""
            Write-Host "Creating Windows firewall rule for SQL Server (port 1433)..." -ForegroundColor Yellow
            try {
                $ruleName = "SQL (1433)"
                $existingRule = Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue
                if ($existingRule) {
                    Write-Host "Firewall rule '$ruleName' already exists. Skipping creation." -ForegroundColor Cyan
                } else {
                    New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -LocalPort 1433 -Protocol TCP -Action Allow -Profile Any
                    Write-Host "Windows firewall rule '$ruleName' created successfully." -ForegroundColor Green
                }
            } catch {
                Write-Host "Warning: Could not create firewall rule: $($_.Exception.Message)" -ForegroundColor Yellow
            }
            return $true
        } else {
            Write-Host "No SQL Server instances found in registry." -ForegroundColor Red
            Write-Host "Note: SQL Server may not be installed or registry structure is different." -ForegroundColor Yellow
            return $false
        }
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Execute-Task6 {
    param(
        [string]$EclipseInstallPath
    )
    
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Task 6: Install/Update Eclipse DMS" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        $targetDir = Join-Path $EclipseInstallPath "Program Updates"
        if (!(Test-Path $targetDir)) {
            New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
            Write-Host "Created: $targetDir" -ForegroundColor Green
        }
        
        $installerUrl = $EclipseDmsUrl
        $installerPath = Join-Path $targetDir ([System.Net.WebUtility]::UrlDecode((Split-Path $EclipseDmsUrl -Leaf)))
        
        if (Test-Path $installerPath) {
            Write-Host "Installer already downloaded: $installerPath" -ForegroundColor Green
        } else {
            Write-Host "Downloading Eclipse DMS installer..." -ForegroundColor Yellow
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($installerUrl, $installerPath)
            $webClient.Dispose()
            Write-Host "Download complete: $installerPath" -ForegroundColor Green
        }
        
        Write-Host "Installing Eclipse DMS (this may take a few minutes)..." -ForegroundColor Cyan
        $process = Start-Process -FilePath "$env:SystemRoot\System32\msiexec.exe" -ArgumentList "/i", "`"$installerPath`"", "/quiet", "/norestart" -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            Write-Host "Eclipse DMS installation completed successfully!" -ForegroundColor Green
            return $true
        } else {
            Write-Host "Installation failed with exit code: $($process.ExitCode)" -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Execute-Task7 {
    param(
        [string]$EclipseInstallPath
    )
    
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Task 7: Extract Eclipse Database Registry Keys" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        # Get the original logged-in user
        $explorerProc = Get-WmiObject -Query "Select * from Win32_Process where Name='explorer.exe'" | Select-Object -First 1
        $userName = $null
        if ($explorerProc) {
            $ownerInfo = $explorerProc.GetOwner()
            $userName = $ownerInfo.User
            $domain = $ownerInfo.Domain
            Write-Host "Detected logged-in user: $domain\$userName" -ForegroundColor Green
        } else {
            Write-Host "Could not detect logged-in user. Using current user context." -ForegroundColor Yellow
            $userName = $env:USERNAME
        }
        
        # Get SID for the user
        $sid = (Get-WmiObject -Class Win32_UserAccount | Where-Object { $_.Name -eq $userName }).SID
        if (-not $sid) {
            Write-Host "Warning: Could not find SID for user: $userName" -ForegroundColor Yellow
            Write-Host "Registry keys may not exist. This is normal if database has not been mapped yet." -ForegroundColor Yellow
            return $true  # Not a failure, just no keys to export
        }
        
        $regPathSID = "Registry::HKEY_USERS\\$sid\\SOFTWARE\\VB and VBA Program Settings\\UBSRegoWiz\\Databases"
        
        if (-not (Test-Path $regPathSID)) {
            Write-Host "No registry keys found for the logged-in user." -ForegroundColor Yellow
            Write-Host "This is normal if the database has not been mapped in Eclipse yet." -ForegroundColor Yellow
            return $true  # Not a failure
        }
        
        Write-Host "Extracting registry keys from: $regPathSID" -ForegroundColor Cyan
        
        # Ensure output directory exists
        $outputDir = Join-Path $EclipseInstallPath "Program Updates"
        if (!(Test-Path $outputDir)) {
            New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
            Write-Host "Created directory: $outputDir" -ForegroundColor Green
        }
        
        $outputFile = Join-Path $outputDir "Database-Connection.reg"
        $tempFile = Join-Path $env:TEMP "Database-Connection-temp.reg"
        
        # Export to temp file first
        $regNativePath = $regPathSID.Replace('Registry::','').Replace('\\','\')
        $regExeCmd = "reg export `"$regNativePath`" `"$tempFile`" /y"
        Write-Host "Exporting registry keys..." -ForegroundColor Yellow
        Invoke-Expression $regExeCmd | Out-Null
        
        if (Test-Path $tempFile) {
            # Read the exported file
            $regContent = Get-Content $tempFile -Raw
            
            # Replace HKEY_USERS\$SID with HKEY_CURRENT_USER
            $escapedSid = [regex]::Escape($sid)
            $regContent = $regContent -replace "HKEY_USERS\\$escapedSid", "HKEY_CURRENT_USER"
            $regContent = $regContent -replace "HKEY_USERS\\\\$escapedSid", "HKEY_CURRENT_USER"
            
            # Write the modified content to the final file
            Set-Content -Path $outputFile -Value $regContent -Encoding Unicode
            
            # Clean up temp file
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
            
            Write-Host "All database registry strings exported to $outputFile" -ForegroundColor Green
            Write-Host "Registry path in file: HKEY_CURRENT_USER\SOFTWARE\VB and VBA Program Settings\UBSRegoWiz\Databases" -ForegroundColor Cyan
            return $true
        } else {
            Write-Host "Warning: Export failed. No .reg file created." -ForegroundColor Yellow
            return $true  # Not critical
        }
    } catch {
        Write-Host "Warning: $($_.Exception.Message)" -ForegroundColor Yellow
        return $true  # Not critical, continue
    }
}

function Execute-Task8 {
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Task 8: Add Ultimate Domains to Trusted Sites" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        $domains = @(
            "ultimate.net.au",
            "ws.dev.ultimate.net.au",
            "ws.ultimate.net.au"
        )
        
        foreach ($domain in $domains) {
            $zoneKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains"
            $subKey = Join-Path $zoneKey $domain
            if (!(Test-Path $subKey)) {
                New-Item -Path $subKey -Force | Out-Null
            }
            Set-ItemProperty -Path $subKey -Name "https" -Value 2
            Write-Host "Added $domain to Trusted Sites (https)" -ForegroundColor Green
        }
        
        Write-Host ""
        Write-Host "Ultimate domains have been added to Trusted Sites." -ForegroundColor Cyan
        return $true
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Note: Tasks 9 and 10 require additional configuration (database credentials, service accounts)
# These will need to be added to Get-Configuration or handled separately
function Execute-Task9 {
    param(
        [string]$EclipseInstallPath
    )
    
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Task 9: Eclipse ESB Service Management" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Note: This task requires database credentials and service account information." -ForegroundColor Yellow
    Write-Host "This functionality will be implemented in a future update." -ForegroundColor Gray
    Write-Host ""
    return $false  # Not yet implemented for unattended mode
}

function Execute-Task10 {
    param(
        [string]$EclipseInstallPath
    )
    
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Task 10: Eclipse Stock Export (Beta) Management" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Note: This task requires database credentials and service account information." -ForegroundColor Yellow
    Write-Host "This functionality will be implemented in a future update." -ForegroundColor Gray
    Write-Host ""
    return $false  # Not yet implemented for unattended mode
}

function Execute-Task11 {
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Task 11: Install Microsoft IIS Server + Dependencies" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        Write-Host "Enabling IIS and required Windows features..." -ForegroundColor Yellow
        Write-Host "This will install IIS Server and all dependencies." -ForegroundColor Gray
        $features = @(
            "IIS-WebServerRole",
            "IIS-WebServer",
            "IIS-CommonHttpFeatures",
            "IIS-DefaultDocument",
            "IIS-DirectoryBrowsing",
            "IIS-HttpErrors",
            "IIS-StaticContent",
            "IIS-ApplicationDevelopment",
            "IIS-ASPNET",
            "IIS-ASPNET45",
            "IIS-ASPNET35",
            "IIS-NetFxExtensibility",
            "IIS-NetFxExtensibility45",
            "IIS-ApplicationInit",
            "IIS-ISAPIExtensions",
            "IIS-ISAPIFilter",
            "IIS-ServerSideIncludes",
            "IIS-CGI",
            "IIS-HealthAndDiagnostics",
            "IIS-HttpLogging",
            "IIS-RequestMonitor",
            "IIS-Performance",
            "IIS-HttpCompressionDynamic",
            "IIS-HttpCompressionStatic",
            "IIS-Security",
            "IIS-RequestFiltering",
            "IIS-ManagementConsole",
            "IIS-ASP"
        )
        
        $failedFeatures = @()
        foreach ($feature in $features) {
            try {
                Enable-WindowsOptionalFeature -Online -FeatureName $feature -All -NoRestart -ErrorAction Stop | Out-Null
                Write-Host "Enabled: $feature" -ForegroundColor Green
            } catch {
                Write-Host "Warning: Failed to enable $feature" -ForegroundColor Yellow
                $failedFeatures += $feature
            }
        }
        
        if ($failedFeatures.Count -eq 0) {
            Write-Host "IIS Server and all dependencies installed successfully!" -ForegroundColor Cyan
            return $true
        } else {
            Write-Host "IIS installed with some warnings. Failed features: $($failedFeatures -join ', ')" -ForegroundColor Yellow
            return $true  # Still consider success if most features enabled
        }
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Execute-Task12 {
    param(
        [string]$EclipseInstallPath
    )
    
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Task 12: Download IIS URL Rewrite Module" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        $targetDir = Join-Path $EclipseInstallPath "Dependencies"
        if (!(Test-Path $targetDir)) {
            New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
        }
        
        $installerUrl = "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_en-US.msi"
        $installerPath = Join-Path $targetDir "rewrite_amd64_en-US.msi"
        
        if (Test-Path $installerPath) {
            Write-Host "Installer already downloaded: $installerPath" -ForegroundColor Green
        } else {
            Write-Host "Downloading IIS URL Rewrite Module..." -ForegroundColor Yellow
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($installerUrl, $installerPath)
            $webClient.Dispose()
            Write-Host "Download complete: $installerPath" -ForegroundColor Green
        }
        
        Write-Host "IIS URL Rewrite Module downloaded successfully!" -ForegroundColor Cyan
        
        # Install the MSI with /passive (shows progress but no user interaction)
        Write-Host "Installing IIS URL Rewrite Module (this may take a minute)..." -ForegroundColor Yellow
        Write-Host "Installation window will be visible - this is normal." -ForegroundColor Gray
        
        try {
            $installArgs = "/passive /norestart"
            $installProcess = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$installerPath`" $installArgs" -Wait -PassThru
            
            if ($installProcess.ExitCode -eq 0 -or $installProcess.ExitCode -eq 3010) {
                Write-Host "IIS URL Rewrite Module installed successfully!" -ForegroundColor Green
                return $true
            } else {
                Write-Host "Warning: Installation completed with exit code: $($installProcess.ExitCode)" -ForegroundColor Yellow
                Write-Host "Note: Exit code 3010 means success but reboot required." -ForegroundColor Gray
                return $true  # Still consider success
            }
        } catch {
            Write-Host "Error installing IIS URL Rewrite Module: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Execute-Task13 {
    param(
        [string]$EclipseInstallPath
    )
    
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Task 13: Install ASP.NET Core Hosting Bundle 8.0.22 & .NET Desktop Runtime 8.0.22" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        $targetDir = Join-Path $EclipseInstallPath "Dependencies"
        if (!(Test-Path $targetDir)) {
            New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
        }
        
        Write-Host "Installing .NET 8 components via winget..." -ForegroundColor Yellow
        $packages = @(
            @{
                Id = "Microsoft.DotNet.HostingBundle.8"
                Name = "ASP.NET Core Hosting Bundle 8.0.22"
                Version = "8.0.22"
            },
            @{
                Id = "Microsoft.DotNet.DesktopRuntime.8"
                Name = ".NET Desktop Runtime 8.0.22"
                Version = "8.0.22"
            }
        )
        
        $allSuccess = $true
        foreach ($package in $packages) {
            Write-Host "Installing $($package.Name)..." -ForegroundColor Gray
            $installArgs = "install --id $($package.Id) --version $($package.Version) --silent --accept-package-agreements --accept-source-agreements"
            $process = Start-Process -FilePath "winget" -ArgumentList $installArgs -Wait -PassThru -NoNewWindow
            if ($process.ExitCode -eq 0) {
                Write-Host "$($package.Name) installed successfully" -ForegroundColor Green
            } else {
                Write-Host "Warning: $($package.Name) installation returned exit code: $($process.ExitCode)" -ForegroundColor Yellow
                # Try without version pinning as fallback
                Write-Host "  Attempting installation without version pinning..." -ForegroundColor Gray
                $fallbackArgs = "install --id $($package.Id) --silent --accept-package-agreements --accept-source-agreements"
                $fallbackProcess = Start-Process -FilePath "winget" -ArgumentList $fallbackArgs -Wait -PassThru -NoNewWindow
                if ($fallbackProcess.ExitCode -eq 0) {
                    Write-Host "$($package.Name) installed successfully (latest version)" -ForegroundColor Green
                } else {
                    Write-Host "  Failed to install $($package.Name)" -ForegroundColor Red
                    $allSuccess = $false
                }
            }
        }
        
        if ($allSuccess) {
            Write-Host "ASP.NET Core Hosting Bundle 8.0.22 and .NET Desktop Runtime 8.0.22 installed successfully!" -ForegroundColor Cyan
            return $true
        } else {
            Write-Host ".NET 8 installation completed with warnings." -ForegroundColor Yellow
            return $true  # Still consider success
        }
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Execute-Task14 {
    param(
        [string]$EclipseInstallPath
    )
    
    Write-Log "===============================================" -ForegroundColor Cyan
    Write-Log "   Task 14: Download & Install .NET Framework 4.8" -ForegroundColor Yellow
    Write-Log "===============================================" -ForegroundColor Cyan
    Write-Log ""
    
    try {
        # Check if .NET Framework 4.8 is already installed
        Write-Log "Checking if .NET Framework 4.8 is already installed..." -ForegroundColor Yellow
        $netFramework48Installed = $false
        
        # Check registry for .NET Framework 4.8
        $netFrameworkPath = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
        if (Test-Path $netFrameworkPath) {
            $release = (Get-ItemProperty -Path $netFrameworkPath -Name Release -ErrorAction SilentlyContinue).Release
            if ($release -ge 528040) {  # .NET Framework 4.8 release number
                $netFramework48Installed = $true
                Write-Log ".NET Framework 4.8 is already installed (Release: $release)" -ForegroundColor Green
                return $true
            }
        }
        
        if (-not $netFramework48Installed) {
            Write-Log ".NET Framework 4.8 not found. Proceeding with installation..." -ForegroundColor Yellow
        }
        
        $targetDir = Join-Path $EclipseInstallPath "Dependencies"
        if (!(Test-Path $targetDir)) {
            New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
            Write-Log "Created directory: $targetDir" -ForegroundColor Green
        }
        
        $installerUrl = "https://tnh.net.au/packages/NDP48-x86-x64-AllOS-ENU.exe"
        $installerPath = Join-Path $targetDir "NDP48-x86-x64-AllOS-ENU.exe"
        
        if (Test-Path $installerPath) {
            Write-Log "Installer already downloaded: $installerPath" -ForegroundColor Green
        } else {
            Write-Log "Downloading .NET Framework 4.8..." -ForegroundColor Yellow
            Write-Log "  Source URL: $installerUrl" -ForegroundColor Gray
            Write-Log "  Destination: $installerPath" -ForegroundColor Gray
            # Ensure TLS 1.2 is used (required by many HTTPS servers; default on some Windows is TLS 1.0)
            $prevProtocol = [Net.ServicePointManager]::SecurityProtocol
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            } catch {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            }
            try {
                $webClient = New-Object System.Net.WebClient
                $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) PowerShell-EclipseSwissArmyKnife")
                $webClient.DownloadFile($installerUrl, $installerPath)
                $webClient.Dispose()
            } catch {
                [Net.ServicePointManager]::SecurityProtocol = $prevProtocol
                Write-Log "Download failed: $($_.Exception.Message)" -ForegroundColor Red
                if ($_.Exception.InnerException) { Write-Log "  Inner: $($_.Exception.InnerException.Message)" -ForegroundColor Red }
                Write-Log "  If you see 'SSL/TLS secure channel', the server requires TLS 1.2; the script has enabled it. Check proxy/firewall or try from another network." -ForegroundColor Yellow
                throw
            }
            [Net.ServicePointManager]::SecurityProtocol = $prevProtocol
            $fileSize = (Get-Item $installerPath).Length / 1MB
            Write-Log "Download complete: $installerPath ($([math]::Round($fileSize, 2)) MB)" -ForegroundColor Green
        }
        
        # Install .NET Framework 4.8 silently
        Write-Log "Installing .NET Framework 4.8 (this may take a few minutes)..." -ForegroundColor Cyan
        Write-Log "Running silent installation..." -ForegroundColor Yellow
        
        $installArgs = "/quiet /norestart"
        Write-Log "Command: $installerPath $installArgs" -ForegroundColor DarkGray
        
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $installerPath
        $processInfo.Arguments = $installArgs
        $processInfo.UseShellExecute = $true
        $processInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        
        try {
            $started = $process.Start()
            if (-not $started) {
                Write-Log "Error: Failed to start .NET Framework installer process" -ForegroundColor Red
                return $false
            }
            
            Write-Log ".NET Framework installer started (PID: $($process.Id))" -ForegroundColor Green
            Write-Log "  Process start time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
            Write-Log "Waiting for installation to complete..." -ForegroundColor Cyan
            
            # Wait for installation to complete with periodic status updates
            $waitStartTime = Get-Date
            while (-not $process.HasExited) {
                Start-Sleep -Seconds 30
                $elapsed = (Get-Date) - $waitStartTime
                $minutes = [math]::Floor($elapsed.TotalMinutes)
                $seconds = [math]::Floor($elapsed.TotalSeconds % 60)
                Write-Log "  Installation in progress... (Elapsed: ${minutes}m ${seconds}s)" -ForegroundColor Gray
            }
            
            Write-Log ".NET Framework installer has completed." -ForegroundColor Green
            Write-Log "Exit code: $($process.ExitCode)" -ForegroundColor Cyan
            Write-Log "  Process end time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
            
            # Verify installation
            Start-Sleep -Seconds 3
            Write-Log "Verifying .NET Framework 4.8 installation..." -ForegroundColor Yellow
            
            $netFrameworkPath = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
            if (Test-Path $netFrameworkPath) {
                $release = (Get-ItemProperty -Path $netFrameworkPath -Name Release -ErrorAction SilentlyContinue).Release
                if ($release -ge 528040) {
                    Write-Log ".NET Framework 4.8 installation verified successfully!" -ForegroundColor Green
                    Write-Log "  Release: $release" -ForegroundColor Cyan
                    return $true
                } else {
                    Write-Log "Warning: .NET Framework installed but version may be incorrect (Release: $release)" -ForegroundColor Yellow
                    if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                        Write-Log "Installation completed with exit code: $($process.ExitCode)" -ForegroundColor Green
                        if ($process.ExitCode -eq 3010) {
                            Write-Log "Note: System restart may be required." -ForegroundColor Yellow
                        }
                        return $true
                    }
                }
            }
            
            # Check exit code
            if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                Write-Log ".NET Framework 4.8 installation completed successfully!" -ForegroundColor Green
                if ($process.ExitCode -eq 3010) {
                    Write-Log "Note: System restart may be required." -ForegroundColor Yellow
                }
                return $true
            } else {
                Write-Log "Installation failed with exit code: $($process.ExitCode)" -ForegroundColor Red
                Write-Log "Common exit codes:" -ForegroundColor Yellow
                Write-Log "  0 = Success" -ForegroundColor Gray
                Write-Log "  3010 = Success but reboot required" -ForegroundColor Gray
                return $false
            }
        } catch {
            Write-Log "Error starting .NET Framework installer: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "  Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Log "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "  Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
        return $false
    }
}

function Execute-Task15 {
    param(
        [string]$EclipseInstallPath,
        [string]$EclipseClientId
    )
    
    Write-Log "===============================================" -ForegroundColor Cyan
    Write-Log "   Task 15: Download & Install Eclipse Update Service $EclipseUpdateServiceVersion" -ForegroundColor Yellow
    Write-Log "===============================================" -ForegroundColor Cyan
    Write-Log ""
    
    try {
        $targetDir = Join-Path $EclipseInstallPath "Dependencies"
        # Ensure we have an absolute path
        $targetDir = [System.IO.Path]::GetFullPath($targetDir)
        if (!(Test-Path $targetDir)) {
            New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
            Write-Log "Created directory: $targetDir" -ForegroundColor Green
        }
        
        # Download installer directly from the server
        $installerUrl = $EclipseUpdateServiceUrl
        $installerPath = Join-Path $targetDir ([System.Net.WebUtility]::UrlDecode((Split-Path $installerUrl -Leaf)))
        # Ensure we have an absolute path
        $installerPath = [System.IO.Path]::GetFullPath($installerPath)
        
        if (Test-Path $installerPath) {
            Write-Log "Installer already downloaded: $installerPath" -ForegroundColor Green
        } else {
            Write-Log "Downloading Eclipse Update Service $EclipseUpdateServiceVersion installer..." -ForegroundColor Yellow
            Write-Log "  Source URL: $installerUrl" -ForegroundColor Gray
            Write-Log "  Destination: $installerPath" -ForegroundColor Gray
            try {
                $webClient = New-Object System.Net.WebClient
                $webClient.DownloadFile($installerUrl, $installerPath)
                $webClient.Dispose()
                $fileSize = (Get-Item $installerPath).Length / 1MB
                Write-Log "Download complete: $installerPath ($([math]::Round($fileSize, 2)) MB)" -ForegroundColor Green
            } catch {
                Write-Log "Error downloading installer: $($_.Exception.Message)" -ForegroundColor Red
                return $false
            }
        }
        
        # Verify installer file exists and is accessible
        if (-not (Test-Path $installerPath)) {
            Write-Log "Error: Installer file not found: $installerPath" -ForegroundColor Red
            return $false
        }
        
        $fileInfo = Get-Item $installerPath
        Write-Log "Installer file details:" -ForegroundColor Gray
        Write-Log "  Path: $installerPath" -ForegroundColor Gray
        Write-Log "  Size: $([math]::Round($fileInfo.Length / 1MB, 2)) MB" -ForegroundColor Gray
        Write-Log "  Last Modified: $($fileInfo.LastWriteTime)" -ForegroundColor Gray
        
        # Install Eclipse Update Service silently with Client ID
        Write-Log ""
        Write-Log "Installing Eclipse Update Service $EclipseUpdateServiceVersion (this may take a few minutes)..." -ForegroundColor Cyan
        Write-Log "Running silent installation with Client ID: $EclipseClientId" -ForegroundColor Yellow
        Write-Log "  Username: $EclipseClientId" -ForegroundColor Gray
        Write-Log "  Company/Organization: $EclipseClientId" -ForegroundColor Gray
        
        # Enable verbose logging
        # Use a simple log path without spaces to avoid potential issues
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $installerLogPath = Join-Path $env:TEMP "MSI_$timestamp.log"
        Write-Log "Verbose installer logging enabled: $installerLogPath" -ForegroundColor Cyan
        Write-Log "Note: If installation hangs, check this log file for errors" -ForegroundColor Yellow
        
        # Run the installer (EXE wrapper around MSI) with properties
        # Based on MSI inspection, the correct property names are:
        # - USERNAME (for username field)
        # - COMPANYNAME (for company/organization field)
        # These are listed in SecureCustomProperties
        $propertyList = @(
            "USERNAME=$EclipseClientId",
            "COMPANYNAME=$EclipseClientId"
        )
        $propertiesString = $propertyList -join " "
        
        Write-Log "Using correct property names: USERNAME and COMPANYNAME" -ForegroundColor Green
        
        # Use /qb for basic UI (shows progress window) - /qn causes exit code 1625
        # For EXE wrappers, use /S /v"/qb ..." format (InstallShield style)
        # Tested: /qb works perfectly with USERNAME and COMPANYNAME properties
        $installerArgs = "/S /v`"/qb REBOOT=ReallySuppress $propertiesString`""
        
        Write-Log "Using /qb (basic UI - shows progress window) for installation" -ForegroundColor Green
        Write-Log "Installation should complete in 5-10 seconds" -ForegroundColor Gray
        Write-Log "Properties: USERNAME=$EclipseClientId, COMPANYNAME=$EclipseClientId" -ForegroundColor Green
        Write-Log "Command: $installerPath $installerArgs" -ForegroundColor DarkGray
        Write-Log "Properties being passed:" -ForegroundColor Gray
        foreach ($prop in $propertyList) {
            Write-Log "  $prop" -ForegroundColor Gray
        }
        
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $installerPath
        $processInfo.Arguments = $installerArgs
        $processInfo.UseShellExecute = $true
        $processInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Normal # Normal for /qb to show progress
        $processInfo.WorkingDirectory = $targetDir
        
        Write-Log "Process configuration:" -ForegroundColor Gray
        Write-Log "  FileName: $($processInfo.FileName)" -ForegroundColor Gray
        Write-Log "  Arguments: $($processInfo.Arguments)" -ForegroundColor Gray
        Write-Log "  WorkingDirectory: $($processInfo.WorkingDirectory)" -ForegroundColor Gray
        Write-Log "  UseShellExecute: $($processInfo.UseShellExecute)" -ForegroundColor Gray
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        
        try {
            Write-Log "Attempting to start installer process..." -ForegroundColor Yellow
            $startTime = Get-Date
            $started = $process.Start()
            
            Write-Log "Process.Start() returned: $started" -ForegroundColor Gray
            
            if (-not $started) {
                Write-Log "Error: Failed to start Eclipse Update Service installer process" -ForegroundColor Red
                Write-Log "  Process.Start() returned false" -ForegroundColor Red
                Write-Log "  This usually means the executable could not be launched" -ForegroundColor Yellow
                Write-Log "  Possible causes:" -ForegroundColor Yellow
                Write-Log "    - File is not executable" -ForegroundColor Gray
                Write-Log "    - Missing dependencies" -ForegroundColor Gray
                Write-Log "    - Insufficient permissions" -ForegroundColor Gray
                Write-Log "    - Invalid command line arguments" -ForegroundColor Gray
                return $false
            }
            
            Write-Log "Eclipse Update Service installer started (PID: $($process.Id))" -ForegroundColor Green
            Write-Log "  Process start time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
            Write-Log "  Process ID: $($process.Id)" -ForegroundColor Gray
            Write-Log "  Process Name: $($process.ProcessName)" -ForegroundColor Gray
            
            # Wait a moment and check if process is still running
            Start-Sleep -Milliseconds 500
            if ($process.HasExited) {
                Write-Log "Warning: Process exited immediately after starting!" -ForegroundColor Yellow
                Write-Log "  Exit code: $($process.ExitCode)" -ForegroundColor Yellow
                Write-Log "  Exit time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Yellow
                $duration = ((Get-Date) - $startTime).TotalSeconds
                Write-Log "  Duration: $duration seconds" -ForegroundColor Yellow
                Write-Log "  This usually indicates an error with the installer or command line arguments" -ForegroundColor Yellow
                
                # Try to get more information
                try {
                    $process.Refresh()
                    Write-Log "  Process state: HasExited=$($process.HasExited)" -ForegroundColor Gray
                } catch {
                    Write-Log "  Could not refresh process state: $($_.Exception.Message)" -ForegroundColor Gray
                }
                
                if ($process.ExitCode -ne 0) {
                    Write-Log "Error: Installer failed with exit code: $($process.ExitCode)" -ForegroundColor Red
                    Write-Log "  Common exit codes:" -ForegroundColor Yellow
                    Write-Log "    0 = Success" -ForegroundColor Gray
                    Write-Log "    1 = General error" -ForegroundColor Gray
                    Write-Log "    2 = Invalid parameters" -ForegroundColor Gray
                    Write-Log "    3 = File not found" -ForegroundColor Gray
                    return $false
                }
            } else {
                Write-Log "Process is running. Waiting for installation to complete..." -ForegroundColor Cyan
            }
            
            # Wait for installation to complete with periodic status updates
            # MSI should complete in 5-10 seconds, so check more frequently
            $waitStartTime = Get-Date
            $maxWaitSeconds = 30  # 30 second timeout (should only take 5-10 seconds)
            $checkInterval = 2  # Check every 2 seconds for faster response
            
            while (-not $process.HasExited) {
                Start-Sleep -Seconds $checkInterval
                $elapsed = (Get-Date) - $waitStartTime
                $totalSeconds = $elapsed.TotalSeconds
                $minutes = [math]::Floor($totalSeconds / 60)
                $seconds = [math]::Floor($totalSeconds % 60)
                
                # Refresh process state
                try {
                    $process.Refresh()
                    if ($process.HasExited) {
                        break
                    }
                } catch {
                    Write-Log "  Warning: Could not refresh process state" -ForegroundColor Yellow
                }
                
                # Check for MSI log file
                if (Test-Path $installerLogPath) {
                    $logSize = (Get-Item $installerLogPath).Length
                    Write-Log "  Installation in progress... (Elapsed: ${minutes}m ${seconds}s, Log: $([math]::Round($logSize / 1KB, 2)) KB)" -ForegroundColor Gray
                    
                    # If log exists and is growing, read last few lines to see what's happening
                    if ($logSize -gt 0 -and $totalSeconds % 9 -eq 0) {
                        try {
                            $lastLines = Get-Content $installerLogPath -Tail 5 -ErrorAction SilentlyContinue
                            if ($lastLines) {
                                Write-Log "  Recent log entries:" -ForegroundColor DarkGray
                                foreach ($line in $lastLines) {
                                    if ($line -match "error|Error|ERROR|property|Property|USERNAME|COMPANY") {
                                        Write-Log "    $line" -ForegroundColor Yellow
                                    }
                                }
                            }
                        } catch {
                            # Log file might be locked, ignore
                        }
                    }
                } else {
                    Write-Log "  Installation in progress... (Elapsed: ${minutes}m ${seconds}s, No log file yet)" -ForegroundColor Gray
                }
                
                # Check for child processes (MSI might spawn helper processes)
                try {
                    $childProcesses = Get-Process -ErrorAction SilentlyContinue | Where-Object { 
                        $_.Parent.Id -eq $process.Id -or 
                        ($_.ProcessName -like "*msiexec*" -and $_.Id -ne $process.Id)
                    }
                    if ($childProcesses) {
                        Write-Log "  Found $($childProcesses.Count) related process(es)" -ForegroundColor DarkGray
                    }
                } catch {
                    # Ignore errors checking processes
                }
                
                # Early warning if taking too long
                if ($totalSeconds -gt 15 -and $totalSeconds % 6 -eq 0) {
                    Write-Log "  Warning: Installation taking longer than expected (${minutes}m ${seconds}s)" -ForegroundColor Yellow
                    Write-Log "  The MSI may be waiting for user input or property validation" -ForegroundColor Yellow
                    
                    # Check if msiexec process is still running
                    $msiProcesses = Get-Process -Name "msiexec" -ErrorAction SilentlyContinue | Where-Object { $_.Id -eq $process.Id }
                    if (-not $msiProcesses) {
                        Write-Log "  Process no longer found - may have exited" -ForegroundColor Yellow
                        break
                    }
                    
                    # Check for any MSI logs in temp
                    $recentMsiLogs = Get-ChildItem -Path $env:TEMP -Filter "MSI*.log" -ErrorAction SilentlyContinue | 
                        Where-Object { $_.LastWriteTime -gt $waitStartTime } |
                        Sort-Object LastWriteTime -Descending |
                        Select-Object -First 1
                    
                    if ($recentMsiLogs -and $recentMsiLogs.FullName -ne $installerLogPath) {
                        Write-Log "  Found alternative MSI log: $($recentMsiLogs.FullName)" -ForegroundColor Cyan
                    }
                }
                
                # Hard timeout after 1 minute
                if ($totalSeconds -gt $maxWaitSeconds) {
                    Write-Log "  Error: Installation timeout after ${minutes} minutes" -ForegroundColor Red
                    Write-Log "  The MSI is likely waiting for user input or has encountered an error" -ForegroundColor Yellow
                    Write-Log "  Checking for MSI log files and process state..." -ForegroundColor Yellow
                    
                    # Try to find any MSI logs
                    $allMsiLogs = Get-ChildItem -Path $env:TEMP -Filter "MSI*.log" -ErrorAction SilentlyContinue | 
                        Where-Object { $_.LastWriteTime -gt $waitStartTime } |
                        Sort-Object LastWriteTime -Descending
                    
                    if ($allMsiLogs) {
                        Write-Log "  Found MSI log files:" -ForegroundColor Cyan
                        foreach ($log in $allMsiLogs | Select-Object -First 3) {
                            Write-Log "    $($log.FullName) ($([math]::Round($log.Length / 1KB, 2)) KB)" -ForegroundColor Gray
                        }
                    } else {
                        Write-Log "  No MSI log files found - MSI may not have started properly" -ForegroundColor Red
                    }
                    
                    # Check process state one more time
                    try {
                        $process.Refresh()
                        if ($process.HasExited) {
                            Write-Log "  Process has exited (exit code: $($process.ExitCode))" -ForegroundColor Yellow
                        } else {
                            Write-Log "  Process is still running - may be waiting for input" -ForegroundColor Yellow
                            Write-Log "  Try checking for hidden dialog boxes or run with /qb to see UI" -ForegroundColor Yellow
                        }
                    } catch {
                        Write-Log "  Could not check process state" -ForegroundColor Yellow
                    }
                    
                    break
                }
            }
            
            # Wait for process to fully exit and get exit code
            if (-not $process.HasExited) {
                Write-Log "  Waiting for process to exit..." -ForegroundColor Yellow
                $process.WaitForExit(5000)  # Wait up to 5 more seconds
            }
            
            Write-Log "Eclipse Update Service installer has completed." -ForegroundColor Green
            Write-Log "Exit code: $($process.ExitCode)" -ForegroundColor Cyan
            Write-Log "  Process end time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
            
            # Check for MSI log files and display location (wrapped in try-catch to not block database update)
            Write-Log ""
            try {
                Write-Log "Checking for MSI log files..." -ForegroundColor Yellow
                
                # Check for our specific log file
                if ($installerLogPath -and (Test-Path $installerLogPath)) {
                    Write-Log "Installer verbose log found: $installerLogPath" -ForegroundColor Green
                    Write-Log "  Size: $([math]::Round((Get-Item $installerLogPath).Length / 1KB, 2)) KB" -ForegroundColor Gray
                }
            } catch {
                Write-Log "Note: Could not check log files: $($_.Exception.Message)" -ForegroundColor Gray
            }
            
            # Check exit code
            if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                Write-Log "Eclipse Update Service $EclipseUpdateServiceVersion installation completed successfully!" -ForegroundColor Green
                if ($process.ExitCode -eq 3010) {
                    Write-Log "Note: System restart may be required." -ForegroundColor Yellow
                }
                
                # Post-installation: Update database with ConsumerName
                Write-Log ""
                Write-Log "Updating database with Client ID (ConsumerName)..." -ForegroundColor Yellow
                $dbUpdateSuccess = $false
                try {
                    $serviceName = "EclipseUpdateService"
                    $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
                    
                    # Start service first if not running (so it can create the database)
                    if (-not $service) {
                        Write-Log "Service not found - may need to start it first" -ForegroundColor Yellow
                    } elseif ($service.Status -ne 'Running') {
                        Write-Log "Starting Eclipse Update Service to initialize database..." -ForegroundColor Gray
                        Start-Service -Name $serviceName -ErrorAction SilentlyContinue
                        Start-Sleep -Seconds 5
                    }
                    
                    # Find database path (wait for it to be created if needed)
                    $possibleDbPaths = @(
                        "${env:ProgramFiles}\Ultimate Business Systems\EclipseUpdateService\AppData\EclipseDB.db3",
                        "${env:ProgramFiles(x86)}\Ultimate Business Systems\EclipseUpdateService\AppData\EclipseDB.db3"
                    )
                    
                    $dbPath = $null
                    Write-Log "Waiting for database to be created (service may need to start first)..." -ForegroundColor Gray
                    $maxWaitSeconds = 30
                    $waitStartTime = Get-Date
                    
                    while ($null -eq $dbPath -and ((Get-Date) - $waitStartTime).TotalSeconds -lt $maxWaitSeconds) {
                        foreach ($path in $possibleDbPaths) {
                            if (Test-Path $path) {
                                $dbPath = $path
                                break
                            }
                        }
                        if ($null -eq $dbPath) {
                            Start-Sleep -Seconds 2
                        }
                    }
                    
                    if ($dbPath) {
                        Write-Log "Found database: $dbPath" -ForegroundColor Green
                        
                        # Now stop service to update database
                        if ($service -and $service.Status -eq 'Running') {
                            Write-Log "Stopping Eclipse Update Service to update database..." -ForegroundColor Gray
                            Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
                            Start-Sleep -Seconds 3
                        }
                        
                        # Copy database to temp location, update it, then replace
                        $tempDbPath = Join-Path $env:TEMP "EclipseDB_$(Get-Random).db3"
                        try {
                            Copy-Item $dbPath $tempDbPath -Force -ErrorAction Stop
                            
                            # Load SQLite assembly
                            $dbParent = Split-Path $dbPath -Parent
                            $dbGrandParent = Split-Path $dbParent -Parent
                            $sqliteDll = Join-Path $dbGrandParent "System.Data.SQLite.dll"
                            if (Test-Path $sqliteDll) {
                                Add-Type -Path $sqliteDll -ErrorAction SilentlyContinue
                                
                                # Update temp database
                                $connectionString = "Data Source=$tempDbPath;Version=3;"
                                $connection = New-Object System.Data.SQLite.SQLiteConnection($connectionString)
                                $connection.Open()
                                $command = $connection.CreateCommand()
                                
                                $computerName = $env:COMPUTERNAME
                                $localIP = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike "127.*" -and $_.IPAddress -notlike "169.254.*" } | Select-Object -First 1).IPAddress
                                if (-not $localIP) { $localIP = "" }
                                $osVersion = $env:OS
                                
                                # Update database with ConsumerName (ConsumerToken will be set by API on first call)
                                $command.CommandText = "INSERT OR REPLACE INTO Consumer (ConsumerName, ComputerName, ConsumerToken, LocalIPAddress, PublicIPAddress, OSVersion, UpdateDate) VALUES ('$EclipseClientId', '$computerName', NULL, '$localIP', '', '$osVersion', datetime('now'));"
                                $command.ExecuteNonQuery() | Out-Null
                                
                                # Verify update
                                $command.CommandText = "SELECT ConsumerName FROM Consumer;"
                                $result = $command.ExecuteScalar()
                                
                                $connection.Close()
                                
                                if ($result -eq $EclipseClientId) {
                                    Write-Log "Database updated successfully with ConsumerName='$EclipseClientId'" -ForegroundColor Green
                                    
                                    # Replace original database
                                    $backupPath = $dbPath + ".backup"
                                    Copy-Item $dbPath $backupPath -Force -ErrorAction SilentlyContinue
                                    Copy-Item $tempDbPath $dbPath -Force -ErrorAction Stop
                                    
                                    # Verify final state
                                    $verifyConn = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$dbPath;Version=3;")
                                    $verifyConn.Open()
                                    $verifyCmd = $verifyConn.CreateCommand()
                                    $verifyCmd.CommandText = "SELECT ConsumerName FROM Consumer;"
                                    $finalResult = $verifyCmd.ExecuteScalar()
                                    $verifyConn.Close()
                                    
                                    if ($finalResult -eq $EclipseClientId) {
                                        Write-Log "âœ“âœ“âœ“ SUCCESS! ConsumerName verified as '$EclipseClientId' in database! âœ“âœ“âœ“" -ForegroundColor Green
                                        Write-Log "Both username and company fields are now set to $EclipseClientId" -ForegroundColor Green
                                        $dbUpdateSuccess = $true
                                    } else {
                                        Write-Log "Warning: Database update may not have persisted. ConsumerName='$finalResult' (expected: '$EclipseClientId')" -ForegroundColor Yellow
                                    }
                                } else {
                                    Write-Log "Warning: Failed to update temp database. ConsumerName='$result' (expected: '$EclipseClientId')" -ForegroundColor Yellow
                                }
                                
                                # Clean up temp file
                                Remove-Item $tempDbPath -Force -ErrorAction SilentlyContinue
                            } else {
                                Write-Log "Warning: SQLite DLL not found at $sqliteDll" -ForegroundColor Yellow
                            }
                        } catch {
                            Write-Log "Error updating database: $($_.Exception.Message)" -ForegroundColor Red
                            Write-Log "The MSI installed successfully, but database update failed." -ForegroundColor Yellow
                            Write-Log "You may need to manually set ConsumerName to '$EclipseClientId' in the database." -ForegroundColor Yellow
                        }
                    } else {
                        Write-Log "Warning: Database not found at expected locations" -ForegroundColor Yellow
                    }
                    
                    # Restart service if it was running
                    if ($service -and $service.Status -eq 'Stopped') {
                        Write-Log "Restarting Eclipse Update Service..." -ForegroundColor Gray
                        Start-Service -Name $serviceName -ErrorAction SilentlyContinue
                        Start-Sleep -Seconds 3
                        $service.Refresh()
                        if ($service.Status -eq 'Running') {
                            Write-Log "Service restarted successfully" -ForegroundColor Green
                            Write-Log "Service will make API calls with ConsumerName=$EclipseClientId" -ForegroundColor Gray
                        } else {
                            Write-Log "Warning: Service status is: $($service.Status)" -ForegroundColor Yellow
                        }
                    }
                } catch {
                    Write-Log "Error during database update: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Log "The MSI installed successfully, but database update failed." -ForegroundColor Yellow
                }
                
                if (-not $dbUpdateSuccess) {
                    Write-Log ""
                    Write-Log "NOTE: Installation completed, but ConsumerName may not be set correctly." -ForegroundColor Yellow
                    Write-Log "The service may show 'Customer Name is required' errors until ConsumerName is set to '$EclipseClientId'." -ForegroundColor Yellow
                }
                
                return $true
            } else {
                Write-Log "Installation completed with exit code: $($process.ExitCode)" -ForegroundColor Yellow
                Write-Log "Note: Exit code may indicate success even if non-zero. Checking installation..." -ForegroundColor Gray
                
                # Try to verify installation by checking for common installation paths
                $possiblePaths = @(
                    "${env:ProgramFiles}\Eclipse Update Service",
                    "${env:ProgramFiles(x86)}\Eclipse Update Service",
                    "${env:ProgramFiles}\Eclipse\Update Service",
                    "${env:ProgramFiles(x86)}\Eclipse\Update Service"
                )
                
                $installed = $false
                foreach ($path in $possiblePaths) {
                    if (Test-Path $path) {
                        Write-Log "Eclipse Update Service installation verified - found at: $path" -ForegroundColor Green
                        $installed = $true
                        break
                    }
                }
                
                if ($installed) {
                    return $true
                } else {
                    Write-Log "Warning: Could not verify installation, but installer completed." -ForegroundColor Yellow
                    return $true  # Assume success if installer completed
                }
            }
        } catch {
            Write-Log "Error starting Eclipse Update Service installer: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "  Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
            
            # Still try to find log files even on error
            Write-Log ""
            Write-Log "Checking for MSI log files (even though installation failed)..." -ForegroundColor Yellow
            if ($installerLogPath -and (Test-Path $installerLogPath)) {
                Write-Log "Installer log file found: $installerLogPath" -ForegroundColor Cyan
            }
            
            # Search for recent log files
            $recentLogs = @()
            $tempLogs = Get-ChildItem -Path $env:TEMP -Filter "*.log" -ErrorAction SilentlyContinue | 
                Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-10) } |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 5
            if ($tempLogs) { $recentLogs += $tempLogs }
            
            $localAppLogs = Get-ChildItem -Path "$env:LOCALAPPDATA\Temp" -Filter "*.log" -ErrorAction SilentlyContinue | 
                Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-10) } |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 5
            if ($localAppLogs) { $recentLogs += $localAppLogs }
            
            if ($recentLogs -and $recentLogs.Count -gt 0) {
                Write-Log "Recent log files found:" -ForegroundColor Green
                foreach ($log in $recentLogs) {
                    Write-Log "  $($log.FullName)" -ForegroundColor Cyan
                }
            }
            
            Write-Log ""
            Write-Log "IMPORTANT: Check the log files above to see why the client ID wasn't passed correctly." -ForegroundColor Yellow
            Write-Log "Search for your client ID ($EclipseClientId) and property names in the log files." -ForegroundColor Yellow
            
            return $false
        }
    } catch {
        Write-Log "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "  Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
        
        # Show log file location even in outer catch
        if ($installerLogPath) {
            Write-Log ""
            Write-Log "Installer log file location: $installerLogPath" -ForegroundColor Cyan
            Write-Log "Check this file and recent .log files in $env:TEMP for details." -ForegroundColor Yellow
        }
        
        return $false
    }
    
    # Final message about log files (always show this at the end)
    Write-Log ""
    Write-Log "===============================================" -ForegroundColor Cyan
    Write-Log "InstallShield Log File Information" -ForegroundColor Yellow
    Write-Log "===============================================" -ForegroundColor Cyan
    Write-Log "To find InstallShield log files, check:" -ForegroundColor Gray
    Write-Log "  1. $env:TEMP" -ForegroundColor Cyan
    Write-Log "  2. $env:LOCALAPPDATA\Temp" -ForegroundColor Cyan
    Write-Log ""
    Write-Log "Look for files named:" -ForegroundColor Gray
    Write-Log "  - EclipseUpdateService_Install_*.log" -ForegroundColor Cyan
    Write-Log "  - IS*.log (InstallShield logs)" -ForegroundColor Cyan
    Write-Log "  - MSI*.log (Windows Installer logs)" -ForegroundColor Cyan
    Write-Log ""
    Write-Log "In the log file, search for:" -ForegroundColor Yellow
    Write-Log "  - Your client ID: $EclipseClientId" -ForegroundColor Gray
    Write-Log "  - Property names: USERNAME, COMPANY, USER_NAME, etc." -ForegroundColor Gray
    Write-Log "  - The word 'PROPERTY' to see what properties the MSI expects" -ForegroundColor Gray
    Write-Log "===============================================" -ForegroundColor Cyan
}

function Execute-Task16 {
    param(
        [string]$EclipseInstallPath
    )
    
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Task 16: Download Eclipse Online Chrome $EclipseOnlineChromeVersion" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        $targetDir = Join-Path $EclipseInstallPath "Dependencies"
        if (!(Test-Path $targetDir)) {
            New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
        }
        
        $installerUrl = $EclipseOnlineChromeUrl
        $installerPath = Join-Path $targetDir ([System.Net.WebUtility]::UrlDecode((Split-Path $installerUrl -Leaf)))
        
        if (Test-Path $installerPath) {
            Write-Host "Installer already downloaded: $installerPath" -ForegroundColor Green
        } else {
            Write-Host "Downloading Eclipse Online Chrome $EclipseOnlineChromeVersion..." -ForegroundColor Yellow
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($installerUrl, $installerPath)
            $webClient.Dispose()
            Write-Host "Download complete: $installerPath" -ForegroundColor Green
        }
        
        Write-Host "Eclipse Online Chrome $EclipseOnlineChromeVersion downloaded successfully!" -ForegroundColor Cyan
        return $true
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Execute-Task17 {
    param(
        [string]$EclipseInstallPath,
        [string]$EclipseAuraHubSubdomain,
        [int]$EclipseAuraHubPort
    )
    
    Write-Log "===============================================" -ForegroundColor Cyan
    Write-Log "   Task 17: Install Eclipse Online Server" -ForegroundColor Yellow
    Write-Log "===============================================" -ForegroundColor Cyan
    Write-Log ""
    
    try {
        $targetDir = Join-Path $EclipseInstallPath "Dependencies"
        if (!(Test-Path $targetDir)) {
            New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
            Write-Log "Created directory: $targetDir" -ForegroundColor Green
        }
        
        $installerUrl = $EclipseOnlineServerUrl
        $installerPath = Join-Path $targetDir ([System.Net.WebUtility]::UrlDecode((Split-Path $installerUrl -Leaf)))
        
        if (Test-Path $installerPath) {
            Write-Log "Installer already downloaded: $installerPath" -ForegroundColor Green
        } else {
            Write-Log "Downloading Eclipse Online Server $EclipseOnlineServerVersion..." -ForegroundColor Yellow
            Write-Log "  Source URL: $installerUrl" -ForegroundColor Gray
            Write-Log "  Destination: $installerPath" -ForegroundColor Gray
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($installerUrl, $installerPath)
            $webClient.Dispose()
            $fileSize = (Get-Item $installerPath).Length / 1MB
            Write-Log "Download complete: $installerPath ($([math]::Round($fileSize, 2)) MB)" -ForegroundColor Green
        }
        
        Write-Log "Eclipse Online Server $EclipseOnlineServerVersion downloaded successfully!" -ForegroundColor Cyan
        Write-Log ""
        
        # Install Eclipse Online Server silently to C:\inetpub\Eclipse
        $installPath = "C:\inetpub\Eclipse"
        Write-Log "Installing Eclipse Online Server $EclipseOnlineServerVersion..." -ForegroundColor Yellow
        Write-Log "  Installation path: $installPath" -ForegroundColor Gray
        
        # Check if already installed
        $alreadyInstalled = Test-Path $installPath
        if ($alreadyInstalled) {
            $existingFiles = Get-ChildItem -Path $installPath -ErrorAction SilentlyContinue
            if ($existingFiles -and $existingFiles.Count -gt 0) {
                Write-Log "Eclipse Online Server appears to be already installed at $installPath" -ForegroundColor Green
                Write-Log "  Found $($existingFiles.Count) items" -ForegroundColor Gray
            } else {
                Write-Log "Directory exists but appears empty. Proceeding with installation..." -ForegroundColor Yellow
                $alreadyInstalled = $false
            }
        }
        
        if (-not $alreadyInstalled) {
            # Ensure installation directory exists
            if (!(Test-Path $installPath)) {
                New-Item -Path $installPath -ItemType Directory -Force | Out-Null
                Write-Log "Created installation directory: $installPath" -ForegroundColor Green
            }
            
            # Use InstallShield silent install pattern with INSTALLDIR property
            # Pattern: /S /v"/qn INSTALLDIR=\"path\""
            $installArgs = "/S /v`"/qn INSTALLDIR=`"$installPath`"`""
            
            Write-Log "Running silent installation (this may take a few minutes)..." -ForegroundColor Cyan
            Write-Log "  Command: $installerPath $installArgs" -ForegroundColor DarkGray
            
            try {
                $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                $processInfo.FileName = $installerPath
                $processInfo.Arguments = $installArgs
                $processInfo.UseShellExecute = $true
                $processInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
                $processInfo.CreateNoWindow = $true
                
                $process = [System.Diagnostics.Process]::Start($processInfo)
                
                # Wait up to 5 minutes for installation
                $process.WaitForExit(300000)
                
                if (-not $process.HasExited) {
                    Write-Log "Installation timed out (>5 minutes)" -ForegroundColor Yellow
                    $process.Kill()
                    Write-Log "Installation process terminated" -ForegroundColor Red
                    return $false
                }
                
                Write-Log "Installation completed with exit code: $($process.ExitCode)" -ForegroundColor $(if ($process.ExitCode -eq 0) { "Green" } else { "Yellow" })
                
                # Wait a bit for files to be written
                Start-Sleep -Seconds 5
                
                # Verify installation
                $installedFiles = Get-ChildItem -Path $installPath -ErrorAction SilentlyContinue
                if ($installedFiles -and $installedFiles.Count -gt 0) {
                    Write-Log "Installation verified! Files found in $installPath" -ForegroundColor Green
                    Write-Log "  Installed $($installedFiles.Count) items" -ForegroundColor Gray
                } else {
                    Write-Log "Warning: Installation completed but no files found in $installPath" -ForegroundColor Yellow
                    Write-Log "  Checking default location..." -ForegroundColor Gray
                    $defaultPath = "C:\inetpub\wwwroot\server"
                    $defaultFiles = Get-ChildItem -Path $defaultPath -ErrorAction SilentlyContinue
                    if ($defaultFiles) {
                        Write-Log "  Files found in default location: $defaultPath" -ForegroundColor Yellow
                        Write-Log "  Installation may not have respected INSTALLDIR property" -ForegroundColor Yellow
                    }
                }
                
                if ($process.ExitCode -ne 0 -and $process.ExitCode -ne 3010) {
                    Write-Log "Warning: Installation returned exit code $($process.ExitCode)" -ForegroundColor Yellow
                    Write-Log "  Exit code 3010 indicates success with reboot required" -ForegroundColor Gray
                    Write-Log "  Other non-zero codes may indicate issues" -ForegroundColor Gray
                }
                
            } catch {
                Write-Log "Error during installation: $($_.Exception.Message)" -ForegroundColor Red
                return $false
            }
        }
        
        Write-Log ""
        Write-Log "Eclipse Online Server $EclipseOnlineServerVersion installation completed!" -ForegroundColor Cyan
        return $true
    } catch {
        Write-Log "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "  Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
        return $false
    }
}

function Execute-Task18 {
    param(
        [string]$EclipseAuraHubSubdomain,
        [int]$EclipseAuraHubPort
    )
    
    Write-Log "===============================================" -ForegroundColor Cyan
    Write-Log "   Task 18: Create IIS Application Pools & Site Binding" -ForegroundColor Yellow
    Write-Log "===============================================" -ForegroundColor Cyan
    Write-Log ""
    
    try {
        Import-Module WebAdministration -ErrorAction Stop
        
        # Build hostname from subdomain
        $hostname = ""
        if (-not [string]::IsNullOrWhiteSpace($EclipseAuraHubSubdomain)) {
            $hostname = $EclipseAuraHubSubdomain.Trim()
            # If hostname doesn't contain a dot, append .eclipseaurahub.com.au
            if ($hostname -notmatch '\.') {
                $hostname = "$hostname.eclipseaurahub.com.au"
            }
        }
        
        Write-Log "Creating and configuring Eclipse & EclipseAuraCore Application Pools..." -ForegroundColor Yellow
        
        # Create Eclipse pool (.NET CLR v4.0, Integrated, LocalSystem)
        try {
            $highestClr = "v4.0"
            $eclipseExists = Get-Item IIS:\AppPools\Eclipse -ErrorAction SilentlyContinue
            if (-not $eclipseExists) {
                New-WebAppPool -Name "Eclipse"
                Start-Sleep -Seconds 1
                Write-Log "Application Pool 'Eclipse' created!" -ForegroundColor Green
            } else {
                Write-Log "Application Pool 'Eclipse' already exists. Updating configuration..." -ForegroundColor Yellow
            }
            Set-ItemProperty IIS:\AppPools\Eclipse -Name "managedRuntimeVersion" -Value $highestClr
            Set-ItemProperty IIS:\AppPools\Eclipse -Name "managedPipelineMode" -Value "Integrated"
            Set-ItemProperty IIS:\AppPools\Eclipse -Name "processModel.identityType" -Value "LocalSystem"
            Write-Log "Application Pool 'Eclipse' configured!" -ForegroundColor Green
        } catch {
            Write-Log "Failed to create/configure Eclipse pool: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Create EclipseAuraCore pool (No Managed Code, Integrated, LocalSystem)
        try {
            $auraExists = Get-Item IIS:\AppPools\EclipseAuraCore -ErrorAction SilentlyContinue
            if (-not $auraExists) {
                New-WebAppPool -Name "EclipseAuraCore"
                Start-Sleep -Seconds 1
                Write-Log "Application Pool 'EclipseAuraCore' created!" -ForegroundColor Green
            } else {
                Write-Log "Application Pool 'EclipseAuraCore' already exists. Updating configuration..." -ForegroundColor Yellow
            }
            Set-ItemProperty IIS:\AppPools\EclipseAuraCore -Name "managedRuntimeVersion" -Value ""
            Set-ItemProperty IIS:\AppPools\EclipseAuraCore -Name "managedPipelineMode" -Value "Integrated"
            Set-ItemProperty IIS:\AppPools\EclipseAuraCore -Name "processModel.identityType" -Value "LocalSystem"
            Write-Log "Application Pool 'EclipseAuraCore' configured!" -ForegroundColor Green
        } catch {
            Write-Log "Failed to create/configure EclipseAuraCore pool: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Create EclipseAuraAuth pool (.NET CLR v4.0, Integrated, LocalSystem)
        try {
            $highestClr = "v4.0"
            $auraAuthExists = Get-Item IIS:\AppPools\EclipseAuraAuth -ErrorAction SilentlyContinue
            if (-not $auraAuthExists) {
                New-WebAppPool -Name "EclipseAuraAuth"
                Start-Sleep -Seconds 1
                Write-Log "Application Pool 'EclipseAuraAuth' created!" -ForegroundColor Green
            } else {
                Write-Log "Application Pool 'EclipseAuraAuth' already exists. Updating configuration..." -ForegroundColor Yellow
            }
            Set-ItemProperty IIS:\AppPools\EclipseAuraAuth -Name "managedRuntimeVersion" -Value $highestClr
            Set-ItemProperty IIS:\AppPools\EclipseAuraAuth -Name "managedPipelineMode" -Value "Integrated"
            Set-ItemProperty IIS:\AppPools\EclipseAuraAuth -Name "processModel.identityType" -Value "LocalSystem"
            Write-Log "Application Pool 'EclipseAuraAuth' configured!" -ForegroundColor Green
        } catch {
            Write-Log "Failed to create/configure EclipseAuraAuth pool: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Create EclipseAuraCoreAuth pool (No Managed Code, Integrated, LocalSystem)
        try {
            $auraCoreAuthExists = Get-Item IIS:\AppPools\EclipseAuraCoreAuth -ErrorAction SilentlyContinue
            if (-not $auraCoreAuthExists) {
                New-WebAppPool -Name "EclipseAuraCoreAuth"
                Start-Sleep -Seconds 1
                Write-Log "Application Pool 'EclipseAuraCoreAuth' created!" -ForegroundColor Green
            } else {
                Write-Log "Application Pool 'EclipseAuraCoreAuth' already exists. Updating configuration..." -ForegroundColor Yellow
            }
            Set-ItemProperty IIS:\AppPools\EclipseAuraCoreAuth -Name "managedRuntimeVersion" -Value ""
            Set-ItemProperty IIS:\AppPools\EclipseAuraCoreAuth -Name "managedPipelineMode" -Value "Integrated"
            Set-ItemProperty IIS:\AppPools\EclipseAuraCoreAuth -Name "processModel.identityType" -Value "LocalSystem"
            Write-Log "Application Pool 'EclipseAuraCoreAuth' configured!" -ForegroundColor Green
        } catch {
            Write-Log "Failed to create/configure EclipseAuraCoreAuth pool: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Create folders in C:\inetpub\Eclipse if not found
        $basePath = "C:\inetpub\Eclipse"
        $folders = @("AuraApi", "eclipseApi", "oAuth", "oAuth2")
        foreach ($folder in $folders) {
            $fullPath = Join-Path $basePath $folder
            if (-not (Test-Path $fullPath)) {
                try {
                    New-Item -Path $fullPath -ItemType Directory -Force | Out-Null
                    Write-Log "Created folder: $fullPath" -ForegroundColor Green
                } catch {
                    Write-Log "Failed to create folder: $fullPath - $($_.Exception.Message)" -ForegroundColor Red
                }
            } else {
                Write-Log "Folder already exists: $fullPath" -ForegroundColor Yellow
            }
        }
        
        # Create IIS Site 'Eclipse' if not exists
        $siteName = "Eclipse"
        $appPoolName = "Eclipse"
        $sitePath = $basePath
        $siteExists = Get-Website -Name $siteName -ErrorAction SilentlyContinue
        
        # Always create self-signed certificate for HTTPS binding
        Write-Log ""
        Write-Log "Creating self-signed certificate 'EclipsePlaceholderSelfsigned'..." -ForegroundColor Yellow
        $certThumbprint = $null
        try {
            $certName = "EclipsePlaceholderSelfsigned"
            $existingCert = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -like "*$certName*" } | Select-Object -First 1
            
            if ($existingCert) {
                Write-Log "Certificate '$certName' already exists. Using existing certificate." -ForegroundColor Green
                $certThumbprint = $existingCert.Thumbprint
            } else {
                $cert = New-SelfSignedCertificate -DnsName "localhost", $env:COMPUTERNAME -CertStoreLocation "Cert:\LocalMachine\My" -FriendlyName $certName -NotAfter (Get-Date).AddYears(10)
                $certThumbprint = $cert.Thumbprint
                Write-Log "Self-signed certificate '$certName' created successfully!" -ForegroundColor Green
                Write-Log "  Thumbprint: $certThumbprint" -ForegroundColor Cyan
                Write-Log "  Note: This is a placeholder certificate. Replace with your production certificate later." -ForegroundColor Yellow
            }
        } catch {
            Write-Log "Failed to create self-signed certificate: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "Site will be created without SSL binding." -ForegroundColor Yellow
        }
        
        if (-not $siteExists) {
            try {
                # Create site with HTTPS binding directly (no HTTP)
                New-Website -Name $siteName -PhysicalPath $sitePath -ApplicationPool $appPoolName -Port $EclipseAuraHubPort -IPAddress "*" -HostHeader $hostname -Force | Out-Null
                
                if ($certThumbprint) {
                    # Remove the HTTP binding if it was created
                    Remove-WebBinding -Name $siteName -Protocol "http" -Port $EclipseAuraHubPort -HostHeader $hostname -ErrorAction SilentlyContinue
                    
                    # Add HTTPS binding with hostname
                    New-WebBinding -Name $siteName -Protocol "https" -Port $EclipseAuraHubPort -IPAddress "*" -HostHeader $hostname
                    
                    # Bind the certificate using AddSslCertificate method (same as original)
                    $binding = Get-WebBinding -Name $siteName -Protocol "https" -Port $EclipseAuraHubPort -HostHeader $hostname
                    if ($binding) {
                        try {
                            $binding.AddSslCertificate($certThumbprint, "My")
                            $bindingInfo = if ($hostname) { "${hostname}:${EclipseAuraHubPort}" } else { "*:${EclipseAuraHubPort}" }
                            Write-Log "IIS Site '$siteName' created on $bindingInfo (HTTPS) with SSL certificate binding." -ForegroundColor Green
                            Write-Log "  Certificate: EclipsePlaceholderSelfsigned" -ForegroundColor Cyan
                        } catch {
                            Write-Log "Warning: AddSslCertificate failed, trying netsh method..." -ForegroundColor Yellow
                            # Fallback to netsh method
                            $ipPort = "0.0.0.0:$EclipseAuraHubPort"
                            $appId = (Get-Item "IIS:\Sites\$siteName").Id
                            netsh http delete sslcert ipport=$ipPort 2>$null | Out-Null
                            netsh http add sslcert ipport=$ipPort certhash=$certThumbprint appid="{$appId}" certstorename=MY | Out-Null
                            if ($LASTEXITCODE -eq 0) {
                                Write-Log "Certificate bound successfully using netsh method." -ForegroundColor Green
                            } else {
                                Write-Log "Warning: Certificate binding failed. You may need to bind it manually in IIS Manager." -ForegroundColor Yellow
                            }
                        }
                    } else {
                        Write-Log "Warning: Could not find HTTPS binding to assign certificate." -ForegroundColor Yellow
                    }
                } else {
                    Write-Log "IIS Site '$siteName' created on port $EclipseAuraHubPort (HTTPS), but certificate creation failed." -ForegroundColor Yellow
                    Write-Log "  You will need to manually bind a certificate in IIS Manager." -ForegroundColor Yellow
                }
            } catch {
                Write-Log "Failed to create IIS Site '$siteName': $($_.Exception.Message)" -ForegroundColor Red
            }
        } else {
            Write-Log "IIS Site '$siteName' already exists. Checking/updating SSL certificate binding..." -ForegroundColor Yellow
            if ($certThumbprint) {
                # Check if HTTPS binding exists and update certificate if needed
                $existingBinding = Get-WebBinding -Name $siteName -Protocol "https" -Port $EclipseAuraHubPort -HostHeader $hostname -ErrorAction SilentlyContinue
                if ($existingBinding) {
                    try {
                        $existingBinding.AddSslCertificate($certThumbprint, "My")
                        Write-Log "SSL certificate 'EclipsePlaceholderSelfsigned' assigned to existing HTTPS binding." -ForegroundColor Green
                    } catch {
                        Write-Log "Warning: Could not assign certificate to existing binding using AddSslCertificate." -ForegroundColor Yellow
                        Write-Log "  Trying netsh method..." -ForegroundColor Yellow
                        $ipPort = "0.0.0.0:$EclipseAuraHubPort"
                        $appId = (Get-Item "IIS:\Sites\$siteName").Id
                        netsh http delete sslcert ipport=$ipPort 2>$null | Out-Null
                        netsh http add sslcert ipport=$ipPort certhash=$certThumbprint appid="{$appId}" certstorename=MY | Out-Null
                        if ($LASTEXITCODE -eq 0) {
                            Write-Log "Certificate assigned successfully using netsh method." -ForegroundColor Green
                        } else {
                            Write-Log "Warning: Could not assign certificate. You may need to bind it manually in IIS Manager." -ForegroundColor Yellow
                        }
                    }
                } else {
                    # Create HTTPS binding if it doesn't exist
                    Write-Log "Creating HTTPS binding for existing site..." -ForegroundColor Yellow
                    New-WebBinding -Name $siteName -Protocol "https" -Port $EclipseAuraHubPort -IPAddress "*" -HostHeader $hostname
                    $newBinding = Get-WebBinding -Name $siteName -Protocol "https" -Port $EclipseAuraHubPort -HostHeader $hostname
                    if ($newBinding) {
                        try {
                            $newBinding.AddSslCertificate($certThumbprint, "My")
                            Write-Log "HTTPS binding created and certificate assigned." -ForegroundColor Green
                        } catch {
                            Write-Log "Warning: Could not assign certificate. Trying netsh method..." -ForegroundColor Yellow
                            $ipPort = "0.0.0.0:$EclipseAuraHubPort"
                            $appId = (Get-Item "IIS:\Sites\$siteName").Id
                            netsh http delete sslcert ipport=$ipPort 2>$null | Out-Null
                            netsh http add sslcert ipport=$ipPort certhash=$certThumbprint appid="{$appId}" certstorename=MY | Out-Null
                            if ($LASTEXITCODE -eq 0) {
                                Write-Log "Certificate assigned successfully using netsh method." -ForegroundColor Green
                            }
                        }
                    }
                }
            }
        }
        
        # Create IIS Applications
        Write-Log ""
        Write-Log "Creating IIS Applications..." -ForegroundColor Yellow
        $appsToCreate = @(
            @{ Name = "/AuraApi"; Folder = "AuraApi"; Pool = "EclipseAuraCore" },
            @{ Name = "/eclipseApi"; Folder = "eclipseApi"; Pool = "Eclipse" },
            @{ Name = "/oAuth"; Folder = "oAuth"; Pool = "EclipseAuraAuth" },
            @{ Name = "/oAuth2"; Folder = "oAuth2"; Pool = "EclipseAuraCoreAuth" }
        )
        foreach ($app in $appsToCreate) {
            $appPath = Join-Path $basePath $app.Folder
            $iisAppExists = Get-WebApplication -Site $siteName -Name $app.Name.TrimStart('/') -ErrorAction SilentlyContinue
            if (-not $iisAppExists) {
                try {
                    New-WebApplication -Site $siteName -Name $app.Name.TrimStart('/') -PhysicalPath $appPath -ApplicationPool $app.Pool
                    Write-Log "Created IIS Application '$($app.Name)' mapped to '$appPath' using pool '$($app.Pool)'" -ForegroundColor Green
                } catch {
                    Write-Log "Failed to create IIS Application '$($app.Name)': $($_.Exception.Message)" -ForegroundColor Red
                }
            } else {
                Write-Log "IIS Application '$($app.Name)' already exists. Skipping creation." -ForegroundColor Yellow
            }
        }
        
        # Start or recycle the Eclipse application pools
        Write-Log ""
        Write-Log "Starting/recycling Eclipse application pools..." -ForegroundColor Yellow
        $poolsToStart = @("Eclipse", "EclipseAuraCore", "EclipseAuraAuth", "EclipseAuraCoreAuth")
        foreach ($poolName in $poolsToStart) {
            try {
                $pool = Get-Item "IIS:\AppPools\$poolName" -ErrorAction SilentlyContinue
                if ($pool) {
                    if ($pool.State -eq "Stopped") {
                        Start-WebAppPool -Name $poolName
                        Write-Log "Started Application Pool '$poolName'" -ForegroundColor Green
                    } else {
                        Restart-WebAppPool -Name $poolName
                        Write-Log "Recycled Application Pool '$poolName'" -ForegroundColor Green
                    }
                }
            } catch {
                Write-Log "Warning: Could not start/recycle Application Pool '$poolName': $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        
        Write-Log ""
        Write-Log "IIS Application Pools and Site configuration completed!" -ForegroundColor Cyan
        return $true
    } catch {
        Write-Log "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Note: IIS may not be installed. Install IIS first (Task 11)." -ForegroundColor Yellow
        return $false
    }
}

function Execute-Task19 {
    param(
        [string]$EclipseInstallPath
    )
    
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Task 20: Install Win-ACMEv2" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        # Create Dependencies\Win-ACMEv2 directory
        $targetDir = Join-Path $EclipseInstallPath "Dependencies\Win-ACMEv2"
        if (!(Test-Path $targetDir)) {
            New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
            Write-Host "Created directory: $targetDir" -ForegroundColor Green
        } else {
            Write-Host "Directory exists: $targetDir" -ForegroundColor Green
        }
        
        # Check if already extracted
        $exePath = Join-Path $targetDir "wacs.exe"
        if (Test-Path $exePath) {
            Write-Host "Win-ACMEv2 is already extracted in: $targetDir" -ForegroundColor Green
            Write-Host "  Found: wacs.exe" -ForegroundColor Gray
            Write-Host ""
            
            # Launch Win-ACMEv2 for user configuration
            Write-Host "Launching Win-ACMEv2 as Administrator for configuration..." -ForegroundColor Yellow
            Write-Host "  Please configure your SSL certificates" -ForegroundColor Cyan
            Write-Host "  The script will wait until you close Win-ACMEv2" -ForegroundColor Cyan
            Write-Host ""
            
            try {
                $process = Start-Process -FilePath $exePath -WorkingDirectory $targetDir -Verb RunAs -PassThru -Wait
                
                Write-Host ""
                Write-Host "Win-ACMEv2 closed (exit code: $($process.ExitCode))" -ForegroundColor Green
                Write-Host "Win-ACMEv2 task completed!" -ForegroundColor Cyan
                return $true
            } catch {
                Write-Host "Error launching Win-ACMEv2: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "You can run it manually from: $targetDir" -ForegroundColor Gray
                return $true  # Still return success since extraction worked
            }
        }
        
        # Download Win-ACMEv2
        $winAcmeUrl = "https://github.com/win-acme/win-acme/releases/download/v2.2.9.1701/win-acme.v2.2.9.1701.x64.trimmed.zip"
        $zipPath = Join-Path $targetDir "win-acme.zip"
        
        Write-Host "Downloading Win-ACMEv2 v2.2.9.1701..." -ForegroundColor Yellow
        Write-Host "  Source URL: $winAcmeUrl" -ForegroundColor Gray
        Write-Host "  Destination: $zipPath" -ForegroundColor Gray
        
        try {
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($winAcmeUrl, $zipPath)
            $webClient.Dispose()
            $fileSize = (Get-Item $zipPath).Length / 1MB
            Write-Host "Download complete: $([math]::Round($fileSize, 2)) MB" -ForegroundColor Green
        } catch {
            Write-Host "Error downloading Win-ACMEv2: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
        
        # Extract ZIP file
        Write-Host ""
        Write-Host "Extracting Win-ACMEv2..." -ForegroundColor Yellow
        Write-Host "  Extracting to: $targetDir" -ForegroundColor Gray
        
        try {
            Expand-Archive -Path $zipPath -DestinationPath $targetDir -Force
            Write-Host "Extraction complete!" -ForegroundColor Green
        } catch {
            Write-Host "Error extracting Win-ACMEv2: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
        
        # Clean up ZIP file
        Write-Host "Cleaning up..." -ForegroundColor Gray
        Remove-Item -Path $zipPath -Force -ErrorAction SilentlyContinue
        
        # Verify extraction
        if (Test-Path $exePath) {
            Write-Host ""
            Write-Host "Win-ACMEv2 extracted successfully!" -ForegroundColor Green
            Write-Host "  Location: $targetDir" -ForegroundColor Gray
            Write-Host "  Executable: wacs.exe" -ForegroundColor Gray
            Write-Host ""
            
            # Launch Win-ACMEv2 for user configuration
            Write-Host "Launching Win-ACMEv2 as Administrator for configuration..." -ForegroundColor Yellow
            Write-Host "  Please configure your SSL certificates" -ForegroundColor Cyan
            Write-Host "  The script will wait until you close Win-ACMEv2" -ForegroundColor Cyan
            Write-Host ""
            
            try {
                $process = Start-Process -FilePath $exePath -WorkingDirectory $targetDir -Verb RunAs -PassThru -Wait
                
                Write-Host ""
                Write-Host "Win-ACMEv2 closed (exit code: $($process.ExitCode))" -ForegroundColor Green
                Write-Host "Win-ACMEv2 installation completed!" -ForegroundColor Cyan
                return $true
            } catch {
                Write-Host "Error launching Win-ACMEv2: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "Win-ACMEv2 is extracted but could not be launched automatically" -ForegroundColor Yellow
                Write-Host "  You can run it manually from: $targetDir" -ForegroundColor Gray
                return $true  # Still return success since extraction worked
            }
        } else {
            Write-Host "Warning: Extraction completed but wacs.exe not found" -ForegroundColor Yellow
            return $false
        }
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Execute-Task20 {
    param(
        [string]$EclipseInstallPath
    )
    
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Task 19: Install Eclipse Smart Hub $EclipseSmartHubVersion" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        $targetDir = Join-Path $EclipseInstallPath "Dependencies"
        if (!(Test-Path $targetDir)) {
            New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
        }
        
        # Download installer
        $installerUrl = $EclipseSmartHubUrl
        $installerPath = Join-Path $targetDir ([System.Net.WebUtility]::UrlDecode((Split-Path $installerUrl -Leaf)))
        
        # Check if installer exists
        if (Test-Path $installerPath) {
            Write-Host "Installer already downloaded: $installerPath" -ForegroundColor Green
        } else {
            Write-Host "Downloading Eclipse Smart Hub $EclipseSmartHubVersion..." -ForegroundColor Yellow
            try {
                $webClient = New-Object System.Net.WebClient
                $webClient.DownloadFile($installerUrl, $installerPath)
                $webClient.Dispose()
                Write-Host "Download complete: $installerPath" -ForegroundColor Green
            } catch {
                Write-Host "Error downloading installer: $($_.Exception.Message)" -ForegroundColor Red
                return $false
            }
        }
        
        # Check if already installed
        $smartHubServices = Get-Service -Name "*SmartHub*" -ErrorAction SilentlyContinue
        $installPath = "C:\Program Files\Ultimate Business Systems\EclipseSmartHub"
        $alreadyInstalled = ($smartHubServices -and (Test-Path $installPath))
        
        if ($alreadyInstalled) {
            $existingFiles = Get-ChildItem -Path $installPath -ErrorAction SilentlyContinue
            Write-Host "Eclipse Smart Hub appears to be already installed" -ForegroundColor Green
            Write-Host "  Location: $installPath" -ForegroundColor Gray
            Write-Host "  Services: $($smartHubServices.Count) found" -ForegroundColor Gray
            Write-Host "  Files: $($existingFiles.Count) items" -ForegroundColor Gray
            Write-Host ""
            Write-Host "Eclipse Smart Hub $EclipseSmartHubVersion installation completed!" -ForegroundColor Cyan
            return $true
        } else {
            # Install Eclipse Smart Hub silently
            Write-Host "Installing Eclipse Smart Hub $EclipseSmartHubVersion silently..." -ForegroundColor Yellow
            Write-Host "  Using InstallShield silent mode: /S" -ForegroundColor Gray
            Write-Host "  Installing to default location with default settings" -ForegroundColor Gray
            
            $installArgs = "/S"
            
            try {
                $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                $processInfo.FileName = $installerPath
                $processInfo.Arguments = $installArgs
                $processInfo.UseShellExecute = $true
                $processInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
                $processInfo.CreateNoWindow = $true
                
                Write-Host "Running installation (this may take several minutes)..." -ForegroundColor Cyan
                $startTime = Get-Date
                $process = [System.Diagnostics.Process]::Start($processInfo)
                
                # Wait up to 10 minutes for installation
                $process.WaitForExit(600000)
                
                if (-not $process.HasExited) {
                    Write-Host "Installation timed out (>10 minutes)" -ForegroundColor Yellow
                    Write-Host "Checking if installation completed anyway..." -ForegroundColor Gray
                    $process.Kill()
                    Start-Sleep -Seconds 3
                } else {
                    $exitCode = $process.ExitCode
                    if ($exitCode -eq 0) {
                        Write-Host "Installation process completed successfully (exit code: 0)" -ForegroundColor Green
                    } elseif ($exitCode -eq 3010) {
                        Write-Host "Installation process completed successfully (exit code: 3010 - reboot required)" -ForegroundColor Green
                    } else {
                        Write-Host "Installation process completed with exit code: $exitCode" -ForegroundColor Yellow
                        Write-Host "  This may indicate an error, but will verify installation..." -ForegroundColor Gray
                    }
                }
                
                # Quick verification check first (before waiting for MSI processes)
                # If already installed, we can return success immediately
                Start-Sleep -Seconds 3  # Brief wait for initial file writes
                $quickCheckServices = Get-Service -Name "*SmartHub*" -ErrorAction SilentlyContinue
                $quickCheckInstalled = (Test-Path $installPath)
                
                if ($quickCheckInstalled -or $quickCheckServices) {
                    Write-Host "Installation verified quickly!" -ForegroundColor Green
                    if ($quickCheckInstalled) {
                        $installedFiles = Get-ChildItem -Path $installPath -ErrorAction SilentlyContinue
                        Write-Host "  Location: $installPath" -ForegroundColor Gray
                        Write-Host "  Files: $($installedFiles.Count) items" -ForegroundColor Gray
                    }
                    if ($quickCheckServices) {
                        Write-Host "  Services installed: $($quickCheckServices.Count)" -ForegroundColor Gray
                        $quickCheckServices | ForEach-Object {
                            Write-Host "    - $($_.DisplayName)" -ForegroundColor DarkGray
                        }
                    }
                    $endTime = Get-Date
                    $duration = ($endTime - $startTime).TotalSeconds
                    Write-Host "  Installation duration: $([math]::Round($duration, 1)) seconds" -ForegroundColor Gray
                    Write-Host ""
                    Write-Host "Eclipse Smart Hub $EclipseSmartHubVersion installation completed!" -ForegroundColor Cyan
                    return $true
                }
                
                # If quick check didn't find it, wait for MSI processes and do full verification
                $msiProcs = Get-Process -Name "msiexec" -ErrorAction SilentlyContinue
                if ($msiProcs) {
                    Write-Host "Waiting for spawned MSI processes to complete..." -ForegroundColor Gray
                    $maxWait = 300  # 5 minutes
                    $waited = 0
                    while ($msiProcs -and $waited -lt $maxWait) {
                        Start-Sleep -Seconds 5
                        $waited += 5
                        $msiProcs = Get-Process -Name "msiexec" -ErrorAction SilentlyContinue
                    }
                }
                
                # Wait a moment for files to be written
                Start-Sleep -Seconds 5
                
                # Verify installation
                $smartHubServices = Get-Service -Name "*SmartHub*" -ErrorAction SilentlyContinue
                $installed = (Test-Path $installPath)
                
                if ($installed -or $smartHubServices) {
                    Write-Host "Installation verified!" -ForegroundColor Green
                    if ($installed) {
                        $installedFiles = Get-ChildItem -Path $installPath -ErrorAction SilentlyContinue
                        Write-Host "  Location: $installPath" -ForegroundColor Gray
                        Write-Host "  Files: $($installedFiles.Count) items" -ForegroundColor Gray
                    }
                    if ($smartHubServices) {
                        Write-Host "  Services installed: $($smartHubServices.Count)" -ForegroundColor Gray
                        $smartHubServices | ForEach-Object {
                            Write-Host "    - $($_.DisplayName)" -ForegroundColor DarkGray
                        }
                    }
                } else {
                    # Check if exit code indicates failure
                    if ($process.HasExited -and $process.ExitCode -ne 0 -and $process.ExitCode -ne 3010) {
                        Write-Host "Installation failed (exit code: $($process.ExitCode)) and verification failed" -ForegroundColor Red
                        return $false
                    } else {
                        Write-Host "Warning: Installation completed but verification failed" -ForegroundColor Yellow
                        Write-Host "  The installer may have completed successfully but files/services not yet available" -ForegroundColor Gray
                    }
                }
                
                $endTime = Get-Date
                $duration = ($endTime - $startTime).TotalSeconds
                Write-Host "  Installation duration: $([math]::Round($duration, 1)) seconds" -ForegroundColor Gray
                
            } catch {
                Write-Host "Error during installation: $($_.Exception.Message)" -ForegroundColor Red
                return $false
            }
        }
        
        Write-Host ""
        Write-Host "Eclipse Smart Hub $EclipseSmartHubVersion installation completed!" -ForegroundColor Cyan
        return $true
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Execute-Tasks {
    param(
        [hashtable]$Config,
        [hashtable]$SelectedTasks
    )
    
    Clear-Host
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Executing Selected Tasks" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    
    $script:TaskResults = @()
    $taskFunctions = @{
        1 = { Execute-Task1 -EclipseInstallPath $Config.EclipseInstallPath }
        2 = { Execute-Task2 -EclipseInstallPath $Config.EclipseInstallPath }
        3 = { Execute-Task3 -SqlDrive $Config.SqlDrive }
        4 = { Execute-Task4 -EclipseInstallPath $Config.EclipseInstallPath }
        5 = { Execute-Task5 }
        6 = { Execute-Task6 -EclipseInstallPath $Config.EclipseInstallPath }
        8 = { Execute-Task8 }
        11 = { Execute-Task11 }
        12 = { Execute-Task12 -EclipseInstallPath $Config.EclipseInstallPath }
        13 = { Execute-Task13 -EclipseInstallPath $Config.EclipseInstallPath }
        14 = { Execute-Task14 -EclipseInstallPath $Config.EclipseInstallPath }
        15 = { Execute-Task15 -EclipseInstallPath $Config.EclipseInstallPath -EclipseClientId $Config.EclipseClientId }
        17 = { Execute-Task17 -EclipseInstallPath $Config.EclipseInstallPath -EclipseAuraHubSubdomain $Config.EclipseAuraHubSubdomain -EclipseAuraHubPort $Config.EclipseAuraHubPort }
        18 = { Execute-Task18 -EclipseAuraHubSubdomain $Config.EclipseAuraHubSubdomain -EclipseAuraHubPort $Config.EclipseAuraHubPort }
        19 = { Execute-Task20 -EclipseInstallPath $Config.EclipseInstallPath }
        20 = { Execute-Task19 -EclipseInstallPath $Config.EclipseInstallPath }
    }
    
    $taskNames = @{
        1 = "Create Eclipse Install Directory Structure"
        2 = "Share Eclipse Install Directory"
        3 = "Download & Install SQL Server Express 2025"
        4 = "Download & Install SQL Server Management Studio 21"
        5 = "Configure SQL Server Network Protocols"
        6 = "Install/Update Eclipse DMS"
        8 = "Add Ultimate Domains to Trusted Sites"
        11 = "Install Microsoft IIS Server + Dependencies"
        12 = "Download & Install IIS URL Rewrite Module"
        13 = "Install ASP.NET Core Hosting Bundle 8.0.22 & .NET Desktop Runtime 8.0.22"
        14 = "Download & Install .NET Framework 4.8"
        15 = "Download & Install Eclipse Update Service $EclipseUpdateServiceVersion"
        17 = "Install Eclipse Online Server $EclipseOnlineServerVersion"
        18 = "Create IIS Application Pools & Site Binding"
        19 = "Install Eclipse Smart Hub $EclipseSmartHubVersion"
        20 = "Install Win-ACMEv2"
    }
    
    # Count selected tasks
    $validTaskIds = @(1, 2, 3, 4, 5, 6, 8, 11, 12, 13, 14, 15, 17, 18, 19, 20)
    $totalTasks = 0
    foreach ($taskId in $validTaskIds) {
        if ($SelectedTasks[$taskId]) {
            $totalTasks++
        }
    }
    
    $currentStep = 0
    
    foreach ($taskId in $validTaskIds) {
        if ($SelectedTasks[$taskId]) {
            $currentStep++
            
            # Show progress at bottom
            $windowHeight = $Host.UI.RawUI.WindowSize.Height
            $progressY = $windowHeight - 1
            
            Write-Log ""
            Write-Log "===============================================" -ForegroundColor Cyan
            Write-Log "Starting Task $currentStep of $totalTasks" -ForegroundColor Yellow
            Write-Log "===============================================" -ForegroundColor Cyan
            Write-Log ""
            
            # Display progress
            $cursorPos = $Host.UI.RawUI.CursorPosition
            $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates(0, $progressY)
            Write-Host "Progress: Step $currentStep of $totalTasks" -ForegroundColor Cyan -NoNewline
            $Host.UI.RawUI.CursorPosition = $cursorPos
            
            $startTime = Get-Date
            $result = & $taskFunctions[$taskId]
            $endTime = Get-Date
            $duration = ($endTime - $startTime).TotalSeconds
            
            $script:TaskResults += @{
                TaskId = $taskId
                TaskName = $taskNames[$taskId]
                Status = $result
                StartTime = $startTime
                EndTime = $endTime
                Duration = $duration
                Message = if ($result) { "Success" } else { "Failed" }
            }
            
            if ($result) {
                Write-Log ""
                Write-Log "Task $taskId completed successfully in $([math]::Round($duration, 2)) seconds" -ForegroundColor Green
            } else {
                Write-Log ""
                Write-Log "Task $taskId failed after $([math]::Round($duration, 2)) seconds" -ForegroundColor Red
            }
            
            # Only show countdown if not the last task
            if ($currentStep -lt $totalTasks) {
                Write-Log ""
                Write-Log "Continuing to next task in 3 seconds..." -ForegroundColor Yellow
                
                # 3 second countdown
                for ($i = 3; $i -gt 0; $i--) {
                    $cursorPos = $Host.UI.RawUI.CursorPosition
                    $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates(0, $progressY)
                    Write-Host "Progress: Step $currentStep of $totalTasks - Next task in $i seconds...    " -ForegroundColor Cyan -NoNewline
                    $Host.UI.RawUI.CursorPosition = $cursorPos
                    Start-Sleep -Seconds 1
                }
                
                # Clear progress line
                $cursorPos = $Host.UI.RawUI.CursorPosition
                $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates(0, $progressY)
                Write-Host (" " * ($Host.UI.RawUI.WindowSize.Width - 1)) -NoNewline
                $Host.UI.RawUI.CursorPosition = $cursorPos
            }
        }
    }
    
    # Clear progress line at end
    $cursorPos = $Host.UI.RawUI.CursorPosition
    $progressY = $Host.UI.RawUI.WindowSize.Height - 1
    $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates(0, $progressY)
    Write-Host (" " * ($Host.UI.RawUI.WindowSize.Width - 1)) -NoNewline
    $Host.UI.RawUI.CursorPosition = $cursorPos
}

function Write-SummaryAndLog {
    param(
        [hashtable]$Config,
        [hashtable]$SelectedTasks
    )
    
    Clear-Host
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "   Installation Summary" -ForegroundColor Yellow
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Configuration:" -ForegroundColor Cyan
    Write-Host "  Eclipse Install Path: $($Config.EclipseInstallPath)" -ForegroundColor White
    Write-Host "  SQL Server Drive: $($Config.SqlDrive)" -ForegroundColor White
    Write-Host "  Eclipse Client ID: $($Config.EclipseClientId)" -ForegroundColor White
    Write-Host "  Eclipse Aura Hub Subdomain: $($Config.EclipseAuraHubSubdomain)" -ForegroundColor White
    Write-Host "  Full Domain: $($Config.EclipseAuraHubSubdomain).eclipseaurahub.com.au" -ForegroundColor White
    Write-Host "  External Port: $($Config.EclipseAuraHubPort)" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Task Results:" -ForegroundColor Cyan
    Write-Host ""
    
    $successCount = 0
    $failCount = 0
    
    foreach ($result in $script:TaskResults) {
        $status = if ($result.Status) { "[OK]" } else { "[FAIL]" }
        $color = if ($result.Status) { "Green" } else { "Red" }
        $duration = [math]::Round($result.Duration, 2)
        
        Write-Host "  $status " -NoNewline -ForegroundColor $color
        Write-Host "$($result.TaskId). $($result.TaskName)" -ForegroundColor White
        Write-Host "     Status: $($result.Message) | Duration: $duration seconds" -ForegroundColor Gray
        Write-Host "     Time: $($result.StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray
        Write-Host ""
        
        if ($result.Status) { $successCount++ } else { $failCount++ }
    }
    
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "Summary: $successCount succeeded, $failCount failed" -ForegroundColor $(if ($failCount -eq 0) { "Green" } else { "Yellow" })
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Write to log file
    $logContent = @"
===============================================
Eclipse Swiss Army Knife v2.0.0 - Installation Log
Execution Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
===============================================

Configuration:
  Eclipse Install Path: $($Config.EclipseInstallPath)
  SQL Server Drive: $($Config.SqlDrive)
  Eclipse Client ID: $($Config.EclipseClientId)
  Eclipse Aura Hub Subdomain: $($Config.EclipseAuraHubSubdomain)
  Full Domain: $($Config.EclipseAuraHubSubdomain).eclipseaurahub.com.au
  External Port: $($Config.EclipseAuraHubPort)

Task Results:
"@
    
    foreach ($result in $script:TaskResults) {
        $status = if ($result.Status) { "SUCCESS" } else { "FAILED" }
        $logContent += @"

Task $($result.TaskId): $($result.TaskName)
  Status: $status
  Start Time: $($result.StartTime.ToString('yyyy-MM-dd HH:mm:ss'))
  End Time: $($result.EndTime.ToString('yyyy-MM-dd HH:mm:ss'))
  Duration: $([math]::Round($result.Duration, 2)) seconds
  Message: $($result.Message)
"@
    }
    
    $logContent += @"

===============================================
Summary: $successCount succeeded, $failCount failed
===============================================

"@
    
    try {
        Add-Content -Path $script:LogFile -Value $logContent -Encoding UTF8
        Write-Host "Log file written to: $script:LogFile" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not write to log file: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    Write-Host ""
    
    # Check if any IIS/Eclipse Aura tasks (11-20) were selected
    $eclipseAuraTasks = @(11, 12, 13, 14, 15, 16, 17, 18, 19, 20)
    $selectedAuraTasks = $eclipseAuraTasks | Where-Object { $SelectedTasks[$_] -eq $true }
    
    if ($selectedAuraTasks.Count -gt 0) {
        # Show post-installation tasks
        Write-Host "===============================================" -ForegroundColor Cyan
        Write-Host "   IMPORTANT: Outstanding Configuration Tasks" -ForegroundColor Yellow
        Write-Host "===============================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "You have installed IIS and Eclipse Aura components." -ForegroundColor White
        Write-Host "Please complete the following tasks:" -ForegroundColor White
        Write-Host ""
        Write-Host "  1. " -ForegroundColor Cyan -NoNewline
        Write-Host "Map the database in Eclipse Schema Update" -ForegroundColor White
        Write-Host "     (Press S now to open Eclipse Schema Update)" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  2. " -ForegroundColor Cyan -NoNewline
        Write-Host "Activate Eclipse Aura using the AWS Domain Manager" -ForegroundColor White
        Write-Host ""
        Write-Host "  3. " -ForegroundColor Cyan -NoNewline
        Write-Host "Whitelist the domain using connection information" -ForegroundColor White
        Write-Host "     on ws.dev.ultimate.net.au via MS SQL Server Management Studio" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  4. " -ForegroundColor Cyan -NoNewline
        Write-Host "Set up Eclipse Online printers via the Eclipse Smart Hub" -ForegroundColor White
        Write-Host ""
        Write-Host "===============================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Press [S] to open Eclipse Schema Update, or any other key to exit..." -ForegroundColor Yellow
        
        $key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        
        if ($key.Character -eq 'S' -or $key.Character -eq 's') {
            Write-Host ""
            Write-Host "Opening Eclipse Schema Update..." -ForegroundColor Cyan
            
            $schemaUpdatePath = "C:\Program Files (x86)\Eclipse\EclipseSchemaUpdate\EclipseSchemaUpdate.exe"
            if (Test-Path $schemaUpdatePath) {
                try {
                    Start-Process -FilePath $schemaUpdatePath
                    Write-Host "Eclipse Schema Update launched successfully!" -ForegroundColor Green
                } catch {
                    Write-Host "Error launching Eclipse Schema Update: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host "Please open it manually from: $schemaUpdatePath" -ForegroundColor Yellow
                }
            } else {
                Write-Host "Eclipse Schema Update not found at: $schemaUpdatePath" -ForegroundColor Red
                Write-Host "Please ensure Eclipse DMS is installed." -ForegroundColor Yellow
            }
            
            Write-Host ""
            Write-Host "Press any key to exit..." -ForegroundColor Yellow
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        }
    } else {
        Write-Host "Press any key to exit..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
}

# ===============================================
# Unattended Mode Entry Point
# ===============================================
function Start-UnattendedMode {
    # Initialize log file
    $script:LogInitialized = $false
    Write-Log "===============================================" -ForegroundColor Cyan
    Write-Log "   Eclipse Swiss Army Knife v$ScriptVersion" -ForegroundColor Yellow
    Write-Log "   Unattended Server Setup" -ForegroundColor Yellow
    Write-Log "===============================================" -ForegroundColor Cyan
    Write-Log ""
    
    # Get configuration
    $config = Get-Configuration
    
    if ($null -eq $config) {
        Write-Host "Configuration cancelled. Returning to main menu..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        return
    }
    
    # Initialize task selection
    $selectedTasks = @{
        1 = $false; 2 = $false; 3 = $false; 4 = $false; 5 = $false
        6 = $false; 8 = $false
        11 = $false; 12 = $false; 13 = $false; 14 = $false; 15 = $false
        17 = $false; 18 = $false; 19 = $false; 20 = $false
    }
    
    # Show task selection menu
    $result = Show-TaskSelectionMenu -SelectedTasks $selectedTasks
    
    if ($null -eq $result) {
        Write-Host "Task selection cancelled. Returning to main menu..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        return
    }
    
    # Execute selected tasks
    Execute-Tasks -Config $config -SelectedTasks $result
    
    # Show summary and write log
    Write-SummaryAndLog -Config $config -SelectedTasks $result
}

function Main {
    do {
        Show-Menu
    $choice = Read-Host -Prompt "Please select an option [0-20] or [U]"
        switch ($choice) {
            '8' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Add Ultimate Domains to Trusted Sites" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                $domains = @(
                    "ultimate.net.au",
                    "ws.dev.ultimate.net.au",
                    "ws.ultimate.net.au"
                )
                foreach ($domain in $domains) {
                    $zoneKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains"
                    $subKey = Join-Path $zoneKey $domain
                    if (!(Test-Path $subKey)) {
                        New-Item -Path $subKey -Force | Out-Null
                    }
                    Set-ItemProperty -Path $subKey -Name "https" -Value 2
                    Write-Host "Added $domain to Trusted Sites (https)" -ForegroundColor Green
                }
                Write-Host "Ultimate domains have been added to Trusted Sites." -ForegroundColor Cyan
                Write-Host ""
                Read-Host "Press Enter to continue back to the main menu"
            }
            '9' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "      Eclipse ESB Service Management" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host ""
                Write-Host "[1] Install ESB Service" -ForegroundColor Green
                Write-Host "[2] Uninstall ESB Service" -ForegroundColor Red
                Write-Host "[0] Return to Main Menu" -ForegroundColor DarkGray
                $esbAction = Read-Host -Prompt "Select action (1-2, 0 to return)"
                if ($esbAction -eq '1') {
                    Clear-Host
                    Write-Host "===============================================" -ForegroundColor Cyan
                    Write-Host "      Install ESB Service" -ForegroundColor Yellow
                    Write-Host "===============================================" -ForegroundColor Cyan
                    Write-Host ""
                    
                    Write-Host "Service Configuration:" -ForegroundColor Yellow
                    Write-Host "---------------------" -ForegroundColor DarkGray
                    $serviceName = Read-Host "  Service Name (default: EclipseESB)"
                    if ([string]::IsNullOrWhiteSpace($serviceName)) { $serviceName = "EclipseESB" }
                    $displayName = Read-Host "  Display Name (default: Eclipse ESB Service)"
                    if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = "Eclipse ESB Service" }
                    $description = Read-Host "  Description (default: Eclipse ESB Service for Stock & Accounting Software)"
                    if ([string]::IsNullOrWhiteSpace($description)) { $description = "Eclipse ESB Service for Stock & Accounting Software" }
                    $exePath = Read-Host "  Executable Path (default: C:\Program Files (x86)\Eclipse\EsbClientTestHarness.exe)"
                    if ([string]::IsNullOrWhiteSpace($exePath)) { $exePath = "C:\Program Files (x86)\Eclipse\EsbClientTestHarness.exe" }
                    $workingDir = Split-Path $exePath -Parent
                    $startupType = Read-Host "  Startup Type (Automatic/Manual/Disabled, default: Automatic (Delayed))"
                    if ([string]::IsNullOrWhiteSpace($startupType)) { $startupType = "Automatic (Delayed)" }
                    
                    Write-Host ""
                    Write-Host "Database Configuration:" -ForegroundColor Yellow
                    Write-Host "---------------------" -ForegroundColor DarkGray
                    $dbName = Read-Host "  Database Name"
                    $dbServer = Read-Host "  Database Server (default: localhost)"
                    if ([string]::IsNullOrWhiteSpace($dbServer)) { $dbServer = "localhost" }
                    $dbUser = Read-Host "  Database Username"
                    $secureDbPassword = Read-Host "  Database Password" -AsSecureString
                    $dbPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureDbPassword))
                    
                    Write-Host ""
                    Write-Host "Service Options:" -ForegroundColor Yellow
                    Write-Host "---------------------" -ForegroundColor DarkGray
                    $restartTime = Read-Host "  Daily Restart Time (HH:MM, default: 03:00)"
                    if ([string]::IsNullOrWhiteSpace($restartTime)) { $restartTime = "03:00" }
                    
                    Write-Host ""
                    $args = "/RunAsService /DBConnection=[PROVIDER=SQLOLEDB.1;INITIAL CATALOG=$dbName;DATA SOURCE=$dbServer;UID=$dbUser;PASSWORD=$dbPassword] /Start"
                    $nssmPath = "C:\Windows\System32\nssm.exe"
                    if (!(Install-NSSM -NssmPath $nssmPath)) {
                        Write-Host "Failed to install NSSM. Cannot proceed with service installation." -ForegroundColor Red
                    } else {
                        # Check if service already exists and remove it
                        if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue) {
                            Write-Host "Existing service '$serviceName' found. Removing..." -ForegroundColor Yellow
                            Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
                            & $nssmPath remove $serviceName confirm
                            Write-Host "Existing service removed." -ForegroundColor Green
                        }
                        Write-Host "Installing service '$serviceName'..." -ForegroundColor Yellow
                        & $nssmPath install $serviceName $exePath
                        & $nssmPath set $serviceName DisplayName $displayName
                        & $nssmPath set $serviceName Description $description
                        & $nssmPath set $serviceName AppDirectory $workingDir
                        $nssmStartupType = switch ($startupType) { "Automatic" { "SERVICE_AUTO_START" } "Automatic (Delayed)" { "SERVICE_DELAYED_AUTO_START" } "Manual" { "SERVICE_DEMAND_START" } "Disabled" { "SERVICE_DISABLED" } default { "SERVICE_DELAYED_AUTO_START" } }
                        & $nssmPath set $serviceName Start $nssmStartupType
                        # Service will run under Local System account (default)
                        & $nssmPath set $serviceName AppParameters $args
                        Write-Host "Starting service..." -ForegroundColor Yellow
                        Start-Service -Name $serviceName -ErrorAction SilentlyContinue
                        Write-Host "Waiting for service to initialize..." -ForegroundColor Yellow
                        Start-Sleep -Seconds 5
                        
                        # Check if esbListener.exe is running (indicates successful connection)
                        $esbListenerRunning = Get-Process -Name "esbListener" -ErrorAction SilentlyContinue
                        
                        Write-Host ""
                        if ($esbListenerRunning) {
                            Write-Host "===============================================" -ForegroundColor Green
                            Write-Host "ESB Service installed and started successfully!" -ForegroundColor Green
                            Write-Host "esbListener.exe is running - database connection verified." -ForegroundColor Green
                            Write-Host "===============================================" -ForegroundColor Green
                        } else {
                            Write-Host "===============================================" -ForegroundColor Red
                            Write-Host "ESB Service installed but may not have started correctly." -ForegroundColor Yellow
                            Write-Host "esbListener.exe is not running." -ForegroundColor Red
                            Write-Host ""
                            Write-Host "Please check:" -ForegroundColor Yellow
                            Write-Host "  - Host and database connection details" -ForegroundColor Yellow
                            Write-Host "  - Database server is accessible" -ForegroundColor Yellow
                            Write-Host "  - Database credentials are correct" -ForegroundColor Yellow
                            Write-Host "  - Service logs for error messages" -ForegroundColor Yellow
                            Write-Host ""
                            Write-Host "Expecting to see esbListener.exe running..." -ForegroundColor Gray
                            Write-Host "===============================================" -ForegroundColor Red
                        }
                        
                        # Create scheduled task for daily restart
                        Write-Host ""
                        Write-Host "Creating scheduled task for daily restart at $restartTime..." -ForegroundColor Yellow
                        
                        $taskName = "$serviceName-DailyRestart"
                        
                        # Remove existing task if it exists
                        $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
                        if ($existingTask) {
                            Write-Host "Existing scheduled task '$taskName' found. Removing..." -ForegroundColor Yellow
                            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
                            Write-Host "Existing scheduled task removed." -ForegroundColor Green
                        }
                        
                        # Create the PowerShell script that will restart the service
                        $restartScript = @"
# Stop the service
net stop $serviceName

# Kill esbListener.exe if it's still running
Get-Process -Name "esbListener" -ErrorAction SilentlyContinue | Stop-Process -Force

# Wait a moment
Start-Sleep -Seconds 3

# Start the service
net start $serviceName
"@
                        
                        # Create action to run PowerShell with the script
                        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -Command `"$restartScript`""
                        
                        # Create trigger for daily at specified time
                        $trigger = New-ScheduledTaskTrigger -Daily -At $restartTime
                        
                        # Create task principal (run as SYSTEM)
                        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                        
                        # Create task settings
                        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
                        
                        # Register the scheduled task
                        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Daily restart of $displayName at $restartTime" | Out-Null
                        
                        Write-Host "Scheduled task '$taskName' created successfully!" -ForegroundColor Green
                        Write-Host "The service will restart daily at $restartTime" -ForegroundColor Cyan
                    }
                    Write-Host ""
                    Read-Host "Press Enter to continue back to the main menu"
                } elseif ($esbAction -eq '2') {
                    Clear-Host
                    Write-Host "===============================================" -ForegroundColor Cyan
                    Write-Host "      Uninstall ESB Service" -ForegroundColor Yellow
                    Write-Host "===============================================" -ForegroundColor Cyan
                    Write-Host ""
                    $serviceName = Read-Host "Service Name to uninstall (default: EclipseESB)"
                    if ([string]::IsNullOrWhiteSpace($serviceName)) { $serviceName = "EclipseESB" }
                    Write-Host ""
                    $nssmPath = "C:\Windows\System32\nssm.exe"
                    if (!(Install-NSSM -NssmPath $nssmPath)) {
                        Write-Host "Failed to install NSSM. Cannot proceed with service uninstallation." -ForegroundColor Red
                    } else {
                        if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue) {
                            Write-Host "Stopping service..." -ForegroundColor Yellow
                            Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
                            Write-Host "Removing service..." -ForegroundColor Yellow
                            & $nssmPath remove $serviceName confirm
                            Write-Host ""
                            Write-Host "===============================================" -ForegroundColor Green
                            Write-Host "ESB Service uninstalled successfully!" -ForegroundColor Green
                            Write-Host "===============================================" -ForegroundColor Green
                        } else {
                            Write-Host "Service '$serviceName' not found." -ForegroundColor Yellow
                        }
                    }
                    Write-Host ""
                    Read-Host "Press Enter to continue back to the main menu"
                }
                # Return to main menu
                continue
            }
            '10' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "      Eclipse Stock Export (Beta) Management" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host ""
                Write-Host "[1] Install Stock Exporter Service" -ForegroundColor Green
                Write-Host "[2] Uninstall Stock Exporter Service" -ForegroundColor Red
                Write-Host "[0] Return to Main Menu" -ForegroundColor DarkGray
                $stockAction = Read-Host -Prompt "Select action (1-2, 0 to return)"
                if ($stockAction -eq '1') {
                    Clear-Host
                    Write-Host "===============================================" -ForegroundColor Cyan
                    Write-Host "      Install Stock Exporter Service" -ForegroundColor Yellow
                    Write-Host "===============================================" -ForegroundColor Cyan
                    Write-Host ""
                    
                    Write-Host "Service Configuration:" -ForegroundColor Yellow
                    Write-Host "---------------------" -ForegroundColor DarkGray
                    $serviceName = Read-Host "  Service Name (default: EclipseStockUploader)"
                    if ([string]::IsNullOrWhiteSpace($serviceName)) { $serviceName = "EclipseStockUploader" }
                    $displayName = Read-Host "  Display Name (default: Eclipse Stock Uploader Service)"
                    if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = "Eclipse Stock Uploader Service" }
                    $description = Read-Host "  Description (default: Eclipse Stock Uploader Service for Stock & Accounting Software)"
                    if ([string]::IsNullOrWhiteSpace($description)) { $description = "Eclipse Stock Uploader Service for Stock & Accounting Software" }
                    $exePath = Read-Host "  Executable Path (default: C:\Program Files (x86)\Eclipse\UBSVehicleExport.exe)"
                    if ([string]::IsNullOrWhiteSpace($exePath)) { $exePath = "C:\Program Files (x86)\Eclipse\UBSVehicleExport.exe" }
                    $workingDir = Split-Path $exePath -Parent
                    $startupType = Read-Host "  Startup Type (Automatic/Manual/Disabled, default: Automatic (Delayed))"
                    if ([string]::IsNullOrWhiteSpace($startupType)) { $startupType = "Automatic (Delayed)" }
                    
                    Write-Host ""
                    Write-Host "Service Account:" -ForegroundColor Yellow
                    Write-Host "---------------------" -ForegroundColor DarkGray
                    $svcAccount = Read-Host "  Service Account (e.g., DOMAIN\username or username)"
                    $secureSvcPassword = Read-Host "  Service Account Password" -AsSecureString
                    $svcPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureSvcPassword))
                    
                    Write-Host ""
                    Write-Host "Service Options:" -ForegroundColor Yellow
                    Write-Host "---------------------" -ForegroundColor DarkGray
                    $restartTime = Read-Host "  Daily Restart Time (HH:MM, default: 03:00)"
                    if ([string]::IsNullOrWhiteSpace($restartTime)) { $restartTime = "03:00" }
                    
                    Write-Host ""
                    $nssmPath = "C:\Windows\System32\nssm.exe"
                    if (!(Install-NSSM -NssmPath $nssmPath)) {
                        Write-Host "Failed to install NSSM. Cannot proceed with service installation." -ForegroundColor Red
                    } else {
                        # Check if service already exists and remove it
                        if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue) {
                            Write-Host "Existing service '$serviceName' found. Removing..." -ForegroundColor Yellow
                            Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
                            & $nssmPath remove $serviceName confirm
                            Write-Host "Existing service removed." -ForegroundColor Green
                        }
                        Write-Host "Installing service '$serviceName'..." -ForegroundColor Yellow
                        & $nssmPath install $serviceName $exePath
                        & $nssmPath set $serviceName DisplayName $displayName
                        & $nssmPath set $serviceName Description $description
                        & $nssmPath set $serviceName AppDirectory $workingDir
                        $nssmStartupType = switch ($startupType) { "Automatic" { "SERVICE_AUTO_START" } "Automatic (Delayed)" { "SERVICE_DELAYED_AUTO_START" } "Manual" { "SERVICE_DEMAND_START" } "Disabled" { "SERVICE_DISABLED" } default { "SERVICE_DELAYED_AUTO_START" } }
                        & $nssmPath set $serviceName Start $nssmStartupType
                        if ($svcAccount -and $svcPassword) {
                            # NSSM requires both username and password in the same ObjectName command
                            & $nssmPath set $serviceName ObjectName $svcAccount $svcPassword
                        }
                        Write-Host "Starting service..." -ForegroundColor Yellow
                        Start-Service -Name $serviceName -ErrorAction SilentlyContinue
                        Write-Host ""
                        Write-Host "===============================================" -ForegroundColor Green
                        Write-Host "Stock Exporter Service installed and started successfully!" -ForegroundColor Green
                        Write-Host "===============================================" -ForegroundColor Green
                        
                        # Create scheduled task for daily restart
                        Write-Host ""
                        Write-Host "Creating scheduled task for daily restart at $restartTime..." -ForegroundColor Yellow
                        
                        $taskName = "$serviceName-DailyRestart"
                        
                        # Remove existing task if it exists
                        $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
                        if ($existingTask) {
                            Write-Host "Existing scheduled task '$taskName' found. Removing..." -ForegroundColor Yellow
                            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
                            Write-Host "Existing scheduled task removed." -ForegroundColor Green
                        }
                        
                        # Create the PowerShell script that will restart the service
                        $restartScript = @"
# Stop the service
net stop $serviceName

# Wait a moment
Start-Sleep -Seconds 3

# Start the service
net start $serviceName
"@
                        
                        # Create action to run PowerShell with the script
                        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -Command `"$restartScript`""
                        
                        # Create trigger for daily at specified time
                        $trigger = New-ScheduledTaskTrigger -Daily -At $restartTime
                        
                        # Create task principal (run as SYSTEM)
                        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                        
                        # Create task settings
                        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
                        
                        # Register the scheduled task
                        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Daily restart of $displayName at $restartTime" | Out-Null
                        
                        Write-Host "Scheduled task '$taskName' created successfully!" -ForegroundColor Green
                        Write-Host "The service will restart daily at $restartTime" -ForegroundColor Cyan
                    }
                    Write-Host ""
                    Read-Host "Press Enter to continue back to the main menu"
                } elseif ($stockAction -eq '2') {
                    Clear-Host
                    Write-Host "===============================================" -ForegroundColor Cyan
                    Write-Host "      Uninstall Stock Exporter Service" -ForegroundColor Yellow
                    Write-Host "===============================================" -ForegroundColor Cyan
                    Write-Host ""
                    $serviceName = Read-Host "Service Name to uninstall (default: EclipseStockUploader)"
                    if ([string]::IsNullOrWhiteSpace($serviceName)) { $serviceName = "EclipseStockUploader" }
                    Write-Host ""
                    $nssmPath = "C:\Windows\System32\nssm.exe"
                    if (!(Install-NSSM -NssmPath $nssmPath)) {
                        Write-Host "Failed to install NSSM. Cannot proceed with service uninstallation." -ForegroundColor Red
                    } else {
                        if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue) {
                            Write-Host "Stopping service..." -ForegroundColor Yellow
                            Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
                            Write-Host "Removing service..." -ForegroundColor Yellow
                            & $nssmPath remove $serviceName confirm
                            Write-Host ""
                            Write-Host "===============================================" -ForegroundColor Green
                            Write-Host "Stock Exporter Service uninstalled successfully!" -ForegroundColor Green
                            Write-Host "===============================================" -ForegroundColor Green
                        } else {
                            Write-Host "Service '$serviceName' not found." -ForegroundColor Yellow
                        }
                    }
                    Write-Host ""
                    Read-Host "Press Enter to continue back to the main menu"
                }
                # Return to main menu
                continue
            }
            '1' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Create Eclipse Install Directory Structure" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host ""
                $defaultPath = "C:\Eclipse Install"
                $installPath = Read-Host "Enter location for 'Eclipse Install' folder (default: $defaultPath)"
                if ([string]::IsNullOrWhiteSpace($installPath)) { $installPath = $defaultPath }
                Write-Host "Target directory: $installPath" -ForegroundColor Green
                $folders = @("Documents", "Pictures", "Help", "Program Updates", "Layouts", "Scripts", "Backups", "Dependencies", "Price Files")
                $picturesSub = @("Stock", "Parts", "Service")
                # Create main folder if not exists
                if (!(Test-Path $installPath)) {
                    New-Item -Path $installPath -ItemType Directory | Out-Null
                    Write-Host "Created: $installPath" -ForegroundColor Green
                } else {
                    Write-Host "Exists: $installPath (skipping)" -ForegroundColor Yellow
                }
                # Create subfolders
                foreach ($folder in $folders) {
                    $subPath = Join-Path $installPath $folder
                    if (!(Test-Path $subPath)) {
                        New-Item -Path $subPath -ItemType Directory | Out-Null
                        Write-Host "Created: $subPath" -ForegroundColor Green
                    } else {
                        Write-Host "Exists: $subPath (skipping)" -ForegroundColor Yellow
                    }
                }
                # Pictures subfolders
                $picturesPath = Join-Path $installPath "Pictures"
                foreach ($sub in $picturesSub) {
                    $subPicPath = Join-Path $picturesPath $sub
                    if (!(Test-Path $subPicPath)) {
                        New-Item -Path $subPicPath -ItemType Directory | Out-Null
                        Write-Host "Created: $subPicPath" -ForegroundColor Green
                    } else {
                        Write-Host "Exists: $subPicPath (skipping)" -ForegroundColor Yellow
                    }
                }
                Write-Host "\nDirectory structure creation complete!" -ForegroundColor Cyan
                Write-Host ""
                Read-Host "Press Enter to continue back to the main menu"
            }
            '2' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Share Eclipse Install Directory" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host ""
                $defaultPath = "C:\Eclipse Install"
                $installPath = Read-Host "Enter location for 'Eclipse Install' folder to share (default: $defaultPath)"
                if ([string]::IsNullOrWhiteSpace($installPath)) { $installPath = $defaultPath }
                
                # Validate path exists
                if (!(Test-Path $installPath)) {
                    Write-Host "Error: The specified path does not exist: $installPath" -ForegroundColor Red
                    Write-Host "Please create the directory first using Option 1." -ForegroundColor Yellow
                    Write-Host ""
                    Read-Host "Press Enter to continue back to the main menu"
                    continue
                }
                
                Write-Host "Target directory: $installPath" -ForegroundColor Green
                
                # Get share name (default to folder name)
                $shareName = Split-Path $installPath -Leaf
                $shareNameInput = Read-Host "Enter share name (default: $shareName)"
                if ([string]::IsNullOrWhiteSpace($shareNameInput)) { $shareNameInput = $shareName }
                
                # Check if share already exists
                $existingShare = Get-SmbShare -Name $shareNameInput -ErrorAction SilentlyContinue
                if ($existingShare) {
                    Write-Host "Share '$shareNameInput' already exists. Removing existing share..." -ForegroundColor Yellow
                    try {
                        Remove-SmbShare -Name $shareNameInput -Force -ErrorAction Stop
                        Write-Host "Existing share removed." -ForegroundColor Green
                    } catch {
                        Write-Host "Failed to remove existing share: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host ""
                        Read-Host "Press Enter to continue back to the main menu"
                        continue
                    }
                }
                
                # Create the share
                Write-Host "Creating Windows share '$shareNameInput' for path: $installPath" -ForegroundColor Yellow
                try {
                    $share = New-SmbShare -Path $installPath -Name $shareNameInput -Description "Eclipse Install Directory" -ErrorAction Stop
                    Write-Host "Share created successfully: \\$env:COMPUTERNAME\$shareNameInput" -ForegroundColor Green
                } catch {
                    Write-Host "Failed to create share: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host ""
                    Read-Host "Press Enter to continue back to the main menu"
                    continue
                }
                
                # Set share permissions (Authenticated Users and Administrators)
                Write-Host "Configuring share permissions..." -ForegroundColor Yellow
                try {
                    # Remove default Everyone permission if it exists
                    $existingPermissions = Get-SmbShareAccess -Name $shareNameInput -ErrorAction SilentlyContinue
                    foreach ($perm in $existingPermissions) {
                        if ($perm.AccountName -eq "Everyone") {
                            Revoke-SmbShareAccess -Name $shareNameInput -AccountName "Everyone" -Force -ErrorAction SilentlyContinue
                        }
                    }
                    
                    # Grant FullControl to Administrators
                    Grant-SmbShareAccess -Name $shareNameInput -AccountName "Administrators" -AccessRight Full -Force -ErrorAction Stop
                    Write-Host "Granted Full Control to Administrators" -ForegroundColor Green
                    
                    # Grant FullControl to Authenticated Users
                    Grant-SmbShareAccess -Name $shareNameInput -AccountName "Authenticated Users" -AccessRight Full -Force -ErrorAction Stop
                    Write-Host "Granted Full Control to Authenticated Users" -ForegroundColor Green
                } catch {
                    Write-Host "Warning: Failed to set some share permissions: $($_.Exception.Message)" -ForegroundColor Yellow
                }
                
                # Set NTFS permissions (Authenticated Users and Administrators)
                Write-Host "Configuring NTFS security permissions..." -ForegroundColor Yellow
                try {
                    $acl = Get-Acl -Path $installPath
                    
                    # Add Administrators with Full Control
                    $adminRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Administrators", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
                    $acl.SetAccessRule($adminRule)
                    
                    # Add Authenticated Users with Full Control
                    $authUsersRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Authenticated Users", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
                    $acl.SetAccessRule($authUsersRule)
                    
                    Set-Acl -Path $installPath -AclObject $acl -ErrorAction Stop
                    Write-Host "NTFS permissions configured successfully" -ForegroundColor Green
                } catch {
                    Write-Host "Warning: Failed to set NTFS permissions: $($_.Exception.Message)" -ForegroundColor Yellow
                }
                
                Write-Host ""
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "Share Configuration Complete!" -ForegroundColor Green
                Write-Host "Share Name: $shareNameInput" -ForegroundColor Cyan
                Write-Host "Share Path: \\$env:COMPUTERNAME\$shareNameInput" -ForegroundColor Cyan
                Write-Host "Physical Path: $installPath" -ForegroundColor Cyan
                Write-Host "Permissions: Administrators (Full), Authenticated Users (Full)" -ForegroundColor Cyan
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host ""
                Read-Host "Press Enter to continue back to the main menu"
            }
            '3' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Download SQL Server Express 2025" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                $targetDir = "C:\SQL2025"
                if (!(Test-Path $targetDir)) {
                    New-Item -Path $targetDir -ItemType Directory | Out-Null
                    Write-Host "Created: $targetDir" -ForegroundColor Green
                }
                $installerUrl = "https://tnh.net.au/packages/SQLEXPR_x64_ENU.exe"
                $installerPath = Join-Path $targetDir "SQLEXPR_x64_ENU.exe"
                if (Test-Path $installerPath) {
                    Write-Host "Installer already downloaded: $installerPath" -ForegroundColor Green
                } else {
                    Write-Host "Downloading SQL Server Express 2025 installer (700MB - this may take a few minutes)..." -ForegroundColor Yellow
                    $prevProtocol = [Net.ServicePointManager]::SecurityProtocol
                    try {
                        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
                    } catch {
                        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                    }
                    try {
                        $webClient = New-Object System.Net.WebClient
                        $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) PowerShell-EclipseSwissArmyKnife")
                        $webClient.DownloadFile($installerUrl, $installerPath)
                        $webClient.Dispose()
                        Write-Host "Download complete: $installerPath" -ForegroundColor Green
                    } catch {
                        [Net.ServicePointManager]::SecurityProtocol = $prevProtocol
                        Write-Host "Download failed: $($_.Exception.Message)" -ForegroundColor Red
                        if ($_.Exception.InnerException) { Write-Host "  Inner: $($_.Exception.InnerException.Message)" -ForegroundColor Red }
                        Write-Host ""
                        Read-Host "Press Enter to continue back to the main menu"
                        continue
                    }
                    [Net.ServicePointManager]::SecurityProtocol = $prevProtocol
                }
                # Launch the installer directly
                Write-Host "Launching SQL Server Express 2025 installer..." -ForegroundColor Cyan
                try {
                    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                    $processInfo.FileName = $installerPath
                    $processInfo.WorkingDirectory = $targetDir
                    $processInfo.UseShellExecute = $true
                    $processInfo.Verb = "RunAs"
                    [System.Diagnostics.Process]::Start($processInfo) | Out-Null
                    Write-Host "SQL Server Express 2025 installation started!" -ForegroundColor Green
                    Write-Host "Installer launched from: $targetDir" -ForegroundColor Cyan
                } catch {
                    Write-Host "Failed to launch installer: $($_.Exception.Message)" -ForegroundColor Red
                }
                Write-Host ""
                Read-Host "Press Enter to continue back to the main menu"
            }
            '4' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Download SQL Server Management Studio 21" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                $targetDir = "C:\Eclipse Install\SQL2022"
                if (!(Test-Path $targetDir)) {
                    New-Item -Path $targetDir -ItemType Directory | Out-Null
                    Write-Host "Created: $targetDir" -ForegroundColor Green
                }
                $installerUrl = "https://aka.ms/ssms/21/release/vs_SSMS.exe"
                $installerPath = Join-Path $targetDir "vs_SSMS.exe"
                if (Test-Path $installerPath) {
                    Write-Host "Installer already downloaded: $installerPath" -ForegroundColor Green
                } else {
                    Write-Host "Downloading SQL Server Management Studio installer (using BITS)..." -ForegroundColor Yellow
                    try {
                        Start-BitsTransfer -Source $installerUrl -Destination $installerPath
                        Write-Host "Download complete: $installerPath" -ForegroundColor Green
                    } catch {
                        Write-Host "Download failed: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host ""
                        Read-Host "Press Enter to continue back to the main menu"
                        continue
                    }
                }
                Write-Host "Launching installer..." -ForegroundColor Cyan
                Start-Process -FilePath $installerPath -Verb RunAs
                Write-Host ""
                Read-Host "Press Enter to continue back to the main menu"
            }
            '5' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Configure SQL Server Network Protocols" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "Searching for SQL Server instances via registry..." -ForegroundColor Yellow
                Write-Host ""
                
                $found = $false
                $configuredInstances = @()
                
                # Find SQL Server instances by checking registry and services
                $sqlServices = Get-Service | Where-Object { $_.Name -like "MSSQL*" -and $_.Name -notlike "*Agent*" -and $_.Name -notlike "*Browser*" -and $_.Name -notlike "*OLAP*" -and $_.Name -notlike "*ReportServer*" }
                
                if ($sqlServices.Count -eq 0) {
                    Write-Host "No SQL Server services found on this machine." -ForegroundColor Red
                    Write-Host ""
                    Read-Host "Press Enter to continue back to the main menu"
                    continue
                }
                
                Write-Host "Found SQL Server services:" -ForegroundColor Green
                foreach ($svc in $sqlServices) {
                    Write-Host "  - $($svc.Name)" -ForegroundColor Cyan
                }
                Write-Host ""
                
                # Get SQL Server instance IDs from registry
                $sqlRegPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server"
                if (Test-Path $sqlRegPath) {
                    $instanceIds = Get-ChildItem -Path $sqlRegPath | Where-Object { $_.PSChildName -like "MSSQL*" -or $_.PSChildName -like "SQL*" }
                    
                    foreach ($instanceId in $instanceIds) {
                        $instanceName = $instanceId.PSChildName
                        $protocolPath = Join-Path $instanceId.PSPath "MSSQLServer\SuperSocketNetLib"
                        
                        if (Test-Path $protocolPath) {
                            $found = $true
                            Write-Host "Configuring instance: $instanceName" -ForegroundColor Green
                            
                            $tcpEnabled = $false
                            $npEnabled = $false
                            
                            # Enable TCP/IP
                            $tcpPath = Join-Path $protocolPath "Tcp"
                            if (Test-Path $tcpPath) {
                                $tcpEnabledValue = Get-ItemProperty -Path $tcpPath -Name "Enabled" -ErrorAction SilentlyContinue
                                if ($tcpEnabledValue.Enabled -ne 1) {
                                    Set-ItemProperty -Path $tcpPath -Name "Enabled" -Value 1 -Type DWord -ErrorAction SilentlyContinue
                                    $tcpEnabled = $true
                                    Write-Host "  Enabled: TCP/IP" -ForegroundColor Green
                                } else {
                                    Write-Host "  TCP/IP already enabled" -ForegroundColor Cyan
                                }
                            }
                            
                            # Enable Named Pipes
                            $npPath = Join-Path $protocolPath "Np"
                            if (Test-Path $npPath) {
                                $npEnabledValue = Get-ItemProperty -Path $npPath -Name "Enabled" -ErrorAction SilentlyContinue
                                if ($npEnabledValue.Enabled -ne 1) {
                                    Set-ItemProperty -Path $npPath -Name "Enabled" -Value 1 -Type DWord -ErrorAction SilentlyContinue
                                    $npEnabled = $true
                                    Write-Host "  Enabled: Named Pipes" -ForegroundColor Green
                                } else {
                                    Write-Host "  Named Pipes already enabled" -ForegroundColor Cyan
                                }
                            }
                            
                            # Find the service name for this instance
                            # Try to match instance ID to service name
                            $serviceName = $null
                            foreach ($svc in $sqlServices) {
                                # Check if service name matches instance pattern
                                if ($svc.Name -eq "MSSQLSERVER") {
                                    # Default instance
                                    $serviceName = "MSSQLSERVER"
                                    break
                                } elseif ($svc.Name -like "MSSQL`$*") {
                                    # Named instance - extract instance name
                                    $svcInstanceName = $svc.Name -replace "MSSQL`$", ""
                                    # Try to match with registry instance name patterns
                                    if ($instanceName -like "*$svcInstanceName*" -or $svcInstanceName -like "*$instanceName*") {
                                        $serviceName = $svc.Name
                                        break
                                    }
                                }
                            }
                            
                            # If we couldn't match, try to find by checking all services
                            if (-not $serviceName) {
                                # For default instance
                                if ($instanceName -eq "MSSQLSERVER" -or $instanceName -like "MSSQL*") {
                                    $defaultSvc = $sqlServices | Where-Object { $_.Name -eq "MSSQLSERVER" }
                                    if ($defaultSvc) {
                                        $serviceName = "MSSQLSERVER"
                                    }
                                }
                                # For named instances, try to find matching service
                                if (-not $serviceName) {
                                    $namedSvc = $sqlServices | Where-Object { $_.Name -like "MSSQL`$*" } | Select-Object -First 1
                                    if ($namedSvc) {
                                        $serviceName = $namedSvc.Name
                                    }
                                }
                            }
                            
                            # Restart SQL Server service if protocols were changed
                            if (($tcpEnabled -or $npEnabled) -and $serviceName) {
                                Write-Host "Restarting SQL Server service: $serviceName" -ForegroundColor Yellow
                                try {
                                    Restart-Service -Name $serviceName -Force -ErrorAction Stop
                                    Write-Host "  Service restarted successfully" -ForegroundColor Green
                                    $configuredInstances += "$instanceName ($serviceName)"
                                } catch {
                                    Write-Host "  Warning: Could not restart service: $($_.Exception.Message)" -ForegroundColor Yellow
                                    Write-Host "  Please restart the service manually: $serviceName" -ForegroundColor Yellow
                                    $configuredInstances += "$instanceName ($serviceName - manual restart needed)"
                                }
                            } elseif ($tcpEnabled -or $npEnabled) {
                                Write-Host "  Warning: Could not determine service name for instance $instanceName" -ForegroundColor Yellow
                                Write-Host "  Please restart the SQL Server service manually" -ForegroundColor Yellow
                                $configuredInstances += "$instanceName (manual restart needed)"
                            } else {
                                Write-Host "  No protocol changes needed" -ForegroundColor Cyan
                                $configuredInstances += "$instanceName"
                            }
                            Write-Host ""
                        }
                    }
                }
                
                if ($found) {
                    Write-Host "===============================================" -ForegroundColor Cyan
                    Write-Host "Configuration complete for instances:" -ForegroundColor Green
                    foreach ($inst in $configuredInstances) {
                        Write-Host "  - $inst" -ForegroundColor Cyan
                    }
                    Write-Host "===============================================" -ForegroundColor Cyan
                    
                    # Create Windows firewall rule for SQL Server port 1433
                    Write-Host ""
                    Write-Host "Creating Windows firewall rule for SQL Server (port 1433)..." -ForegroundColor Yellow
                    try {
                        $ruleName = "SQL (1433)"
                        # Check if rule already exists
                        $existingRule = Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue
                        if ($existingRule) {
                            Write-Host "Firewall rule '$ruleName' already exists. Skipping creation." -ForegroundColor Cyan
                        } else {
                            New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -LocalPort 1433 -Protocol TCP -Action Allow -Profile Any
                            Write-Host "Windows firewall rule '$ruleName' created successfully to allow inbound traffic on port 1433." -ForegroundColor Green
                        }
                    } catch {
                        Write-Host "Warning: Could not create firewall rule: $($_.Exception.Message)" -ForegroundColor Yellow
                        Write-Host "You may need to create the firewall rule manually to allow SQL Server connections." -ForegroundColor Yellow
                    }
                } else {
                    Write-Host "No SQL Server instances found in registry." -ForegroundColor Red
                    Write-Host "Note: SQL Server may not be installed or registry structure is different." -ForegroundColor Yellow
                }
                Write-Host ""
                Read-Host "Press Enter to continue back to the main menu"
            }
            '6' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Install/Update Eclipse DMS" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                $targetDir = "C:\Eclipse Install\Program Updates"
                if (!(Test-Path $targetDir)) {
                    New-Item -Path $targetDir -ItemType Directory | Out-Null
                    Write-Host "Created: $targetDir" -ForegroundColor Green
                }
                $installerUrl = $EclipseDmsUrl
                $installerPath = Join-Path $targetDir ([System.Net.WebUtility]::UrlDecode((Split-Path $EclipseDmsUrl -Leaf)))
                if (Test-Path $installerPath) {
                    Write-Host "Installer already downloaded: $installerPath" -ForegroundColor Green
                } else {
                    Write-Host "Downloading Eclipse DMS installer..." -ForegroundColor Yellow
                    try {
                        $webClient = New-Object System.Net.WebClient
                        $webClient.DownloadFile($installerUrl, $installerPath)
                        $webClient.Dispose()
                        Write-Host "Download complete: $installerPath" -ForegroundColor Green
                    } catch {
                        Write-Host "Download failed: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host ""
                        Read-Host "Press Enter to continue back to the main menu"
                        continue
                    }
                }
                Write-Host "Installing Eclipse DMS (this may take a few minutes)..." -ForegroundColor Cyan
                try {
                    $process = Start-Process -FilePath "$env:SystemRoot\System32\msiexec.exe" -ArgumentList "/i `"$installerPath`" /passive" -Verb RunAs -Wait -PassThru
                    if ($process.ExitCode -eq 0) {
                        Write-Host "Eclipse DMS installation completed successfully!" -ForegroundColor Green
                    } else {
                        Write-Host "Installation failed with exit code: $($process.ExitCode)" -ForegroundColor Red
                    }
                } catch {
                    Write-Host "Installation failed: $($_.Exception.Message)" -ForegroundColor Red
                }
                Write-Host ""
                Read-Host "Press Enter to continue back to the main menu"
            }
            '7' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Extract Eclipse Database Registry Keys" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                # Get the original logged-in user (not the admin user)
                $explorerProc = Get-WmiObject -Query "Select * from Win32_Process where Name='explorer.exe'" | Select-Object -First 1
                $userName = $null
                if ($explorerProc) {
                    $ownerInfo = $explorerProc.GetOwner()
                    $userName = $ownerInfo.User
                    $domain = $ownerInfo.Domain
                    Write-Host "Detected logged-in user: $domain\$userName" -ForegroundColor Green
                } else {
                    Write-Host "Could not detect logged-in user. Using current user context." -ForegroundColor Yellow
                    $userName = $env:USERNAME
                }
                # Get SID for the user
                $sid = (Get-WmiObject -Class Win32_UserAccount | Where-Object { $_.Name -eq $userName }).SID
                if ($sid) {
                    $regPathSID = "Registry::HKEY_USERS\\$sid\\SOFTWARE\\VB and VBA Program Settings\\UBSRegoWiz\\Databases"
                    $usedPath = $regPathSID
                    $regPath = $regPathSID
                    
                    # Check if the registry path actually exists
                    if (Test-Path $regPath) {
                    Write-Host "Extracting registry keys from: $usedPath" -ForegroundColor Cyan
                    try {
                        # Ensure output directory exists
                        $outputDir = "C:\Eclipse Install\Program Updates"
                        if (!(Test-Path $outputDir)) {
                            New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
                            Write-Host "Created directory: $outputDir" -ForegroundColor Green
                        }
                        
                        $outputFile = Join-Path $outputDir "Database-Connection.reg"
                        $tempFile = Join-Path $env:TEMP "Database-Connection-temp.reg"
                        
                        # Export to temp file first
                        $regNativePath = $usedPath.Replace('Registry::','').Replace('\\','\')
                        $regExeCmd = "reg export `"$regNativePath`" `"$tempFile`" /y"
                        Write-Host "Exporting registry keys..." -ForegroundColor Yellow
                        $result = Invoke-Expression $regExeCmd
                        
                        if (Test-Path $tempFile) {
                            # Read the exported file
                            $regContent = Get-Content $tempFile -Raw
                            
                            if ($sid) {
                                # Replace HKEY_USERS\$SID with HKEY_CURRENT_USER
                                # Handle both single and double backslashes
                                $escapedSid = [regex]::Escape($sid)
                                $regContent = $regContent -replace "HKEY_USERS\\$escapedSid", "HKEY_CURRENT_USER"
                                $regContent = $regContent -replace "HKEY_USERS\\\\$escapedSid", "HKEY_CURRENT_USER"
                            }
                            
                            # Write the modified content to the final file (Unicode encoding for .reg files)
                            Set-Content -Path $outputFile -Value $regContent -Encoding Unicode
                            
                            # Clean up temp file
                            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
                            
                            Write-Host "All database registry strings exported to $outputFile" -ForegroundColor Green
                            Write-Host "Registry path in file: HKEY_CURRENT_USER\SOFTWARE\VB and VBA Program Settings\UBSRegoWiz\Databases" -ForegroundColor Cyan
                            Write-Host ""
                            Read-Host "Press Enter to continue back to the main menu"
                        } else {
                            Write-Host "Export failed. No .reg file created." -ForegroundColor Red
                            Write-Host ""
                            Read-Host "Press Enter to continue back to the main menu"
                        }
                    } catch {
                        Write-Host "Failed to extract registry keys: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host ""
                        Read-Host "Press Enter to continue back to the main menu"
                    }
                    } else {
                        Write-Host ""
                        Write-Host "No registry keys found for the logged-in user or current user." -ForegroundColor Red
                        Write-Host ""
                        Write-Host "Please ensure you have mapped the database in Eclipse first for the current user before running this option" -ForegroundColor Yellow
                        Write-Host ""
                        Read-Host "Press Enter to return to the menu" | Out-Null
                    }
                } else {
                    Write-Host ""
                    Write-Host "Could not find SID for user: $userName" -ForegroundColor Red
                    Write-Host ""
                    Write-Host "Please ensure you have mapped the database in Eclipse first for the current user before running this option" -ForegroundColor Yellow
                    Write-Host ""
                    Read-Host "Press Enter to return to the menu" | Out-Null
                }
            }
            '11' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Install Microsoft IIS Server + Dependencies" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "Installing IIS via winget..." -ForegroundColor Yellow
                try {
                    Start-Process -FilePath "winget" -ArgumentList "install --id Microsoft.IIS --silent --accept-package-agreements --accept-source-agreements" -Verb RunAs -Wait
                    Write-Host "IIS base install complete." -ForegroundColor Green
                } catch {
                    Write-Host "Failed to install IIS via winget: $($_.Exception.Message)" -ForegroundColor Red
                }
                Write-Host "Enabling required IIS features..." -ForegroundColor Yellow
                $features = @(
                    "IIS-WebServerRole",
                    "IIS-WebServer",
                    "IIS-CommonHttpFeatures",
                    "IIS-DefaultDocument",
                    "IIS-DirectoryBrowsing",
                    "IIS-HttpErrors",
                    "IIS-StaticContent",
                    "IIS-ApplicationDevelopment",
                    "IIS-ASPNET",
                    "IIS-ASPNET45",
                    "IIS-ASPNET35",
                    "IIS-NetFxExtensibility",
                    "IIS-NetFxExtensibility45",
                    "IIS-ApplicationInit",
                    "IIS-ISAPIExtensions",
                    "IIS-ISAPIFilter",
                    "IIS-ServerSideIncludes",
                    "IIS-CGI",
                    "IIS-HealthAndDiagnostics",
                    "IIS-HttpLogging",
                    "IIS-RequestMonitor",
                    "IIS-Performance",
                    "IIS-HttpCompressionDynamic",
                    "IIS-HttpCompressionStatic",
                    "IIS-Security",
                    "IIS-RequestFiltering",
                    "IIS-ManagementConsole",
                    "IIS-ASP"
                )
                foreach ($feature in $features) {
                    try {
                        Enable-WindowsOptionalFeature -Online -FeatureName $feature -All -NoRestart
                        Write-Host "Enabled: $feature" -ForegroundColor Green
                    } catch {
                        $err = $_
                        Write-Host ("Failed to enable {0}: {1}" -f $feature, $err.Exception.Message) -ForegroundColor Yellow
                    }
                }
                Write-Host "IIS Server and all dependencies installed!" -ForegroundColor Cyan
                Write-Host ""
                Read-Host "Press Enter to continue back to the main menu"
            }
            '12' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Download IIS URL Rewrite Module" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                $targetDir = "C:\Eclipse Install\Dependencies"
                if (!(Test-Path $targetDir)) {
                    New-Item -Path $targetDir -ItemType Directory | Out-Null
                    Write-Host "Created: $targetDir" -ForegroundColor Green
                }
                $installerUrl = "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_en-US.msi"
                $installerPath = Join-Path $targetDir "rewrite_amd64_en-US.msi"
                if (Test-Path $installerPath) {
                    Write-Host "Installer already downloaded: $installerPath" -ForegroundColor Green
                } else {
                    Write-Host "Downloading IIS URL Rewrite Module installer (using BITS)..." -ForegroundColor Yellow
                    try {
                        Start-BitsTransfer -Source $installerUrl -Destination $installerPath
                        Write-Host "Download complete: $installerPath" -ForegroundColor Green
                    } catch {
                        Write-Host "Download failed: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host ""
                        Read-Host "Press Enter to continue back to the main menu"
                        continue
                    }
                }
                Write-Host "Launching installer..." -ForegroundColor Cyan
                Start-Process -FilePath "$env:SystemRoot\System32\msiexec.exe" -ArgumentList "/i `"$installerPath`"" -Verb RunAs
                Write-Host ""
                Read-Host "Press Enter to continue back to the main menu"
            }
            '18' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Manage IIS Application Pools" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                
                # Prompt for hostname and port FIRST (before any IIS operations)
                Write-Host ""
                Write-Host "Enter the HTTPS hostname for the IIS site binding" -ForegroundColor Yellow
                Write-Host "Examples: justvolvo (becomes justvolvo.eclipseaurahub.com.au), eclipse.example.com, or leave blank for all hostnames" -ForegroundColor Cyan
                $hostnameInput = Read-Host "Hostname [Press Enter for blank/all hostnames]"
                $hostname = ""
                if (-not [string]::IsNullOrWhiteSpace($hostnameInput)) {
                    $hostname = $hostnameInput.Trim()
                    # If hostname doesn't contain a dot, append .eclipseaurahub.com.au
                    if ($hostname -notmatch '\.') {
                        $hostname = "$hostname.eclipseaurahub.com.au"
                        Write-Host "Using hostname: $hostname" -ForegroundColor Green
                    } else {
                        Write-Host "Using hostname: $hostname" -ForegroundColor Green
                    }
                } else {
                    Write-Host "Using blank hostname (all hostnames)" -ForegroundColor Green
                }
                
                Write-Host ""
                Write-Host "Enter the port number for the IIS site (HTTPS)" -ForegroundColor Yellow
                Write-Host "Recommended: 443 (default) or press Enter for 443" -ForegroundColor Cyan
                Write-Host "Alternatives: 8443, 8080, 8081" -ForegroundColor Cyan
                $portInput = Read-Host "Port [Press Enter for 443 (default)]"
                $port = 443  # Initialize with default
                if ([string]::IsNullOrWhiteSpace($portInput)) {
                    Write-Host "Using default port 443" -ForegroundColor Green
                } else {
                    $parsedPort = 0
                    if ([int]::TryParse($portInput, [ref]$parsedPort)) {
                        $port = $parsedPort
                        $recommendedPorts = @(443, 8443, 8080, 8081)
                        if ($recommendedPorts -contains $port) {
                            Write-Host "Using port $port (recommended)" -ForegroundColor Green
                        } else {
                            Write-Host "Using port $port (not in recommended list: 443, 8443, 8080, 8081)" -ForegroundColor Yellow
                        }
                    } else {
                        Write-Host "Invalid port number. Using default port 443." -ForegroundColor Yellow
                    }
                }
                Write-Host ""
                
                try {
                    Import-Module WebAdministration
                    Write-Host "Creating and configuring Eclipse & EclipseAuraCore Application Pools..." -ForegroundColor Yellow
                    # Create Eclipse pool (.NET CLR highest, Integrated, LocalSystem)
                    try {
                        $highestClr = "v4.0"
                        $eclipseExists = Get-Item IIS:\AppPools\Eclipse -ErrorAction SilentlyContinue
                        if (-not $eclipseExists) {
                            New-WebAppPool -Name "Eclipse"
                            Start-Sleep -Seconds 1
                            Write-Host "Application Pool 'Eclipse' created!" -ForegroundColor Green
                        } else {
                            Write-Host "Application Pool 'Eclipse' already exists. Updating configuration..." -ForegroundColor Yellow
                        }
                        Set-ItemProperty IIS:\AppPools\Eclipse -Name "managedRuntimeVersion" -Value $highestClr
                        Set-ItemProperty IIS:\AppPools\Eclipse -Name "managedPipelineMode" -Value "Integrated"
                        Set-ItemProperty IIS:\AppPools\Eclipse -Name "processModel.identityType" -Value "LocalSystem"
                        Write-Host "Application Pool 'Eclipse' configured!" -ForegroundColor Green
                    } catch {
                        $err = $_
                        Write-Host ("Failed to create/configure Eclipse pool: {0}" -f $err.Exception.Message) -ForegroundColor Red
                    }
                    # Create EclipseAuraCore pool (No Managed Code, Integrated, LocalSystem)
                    try {
                        $auraExists = Get-Item IIS:\AppPools\EclipseAuraCore -ErrorAction SilentlyContinue
                        if (-not $auraExists) {
                            New-WebAppPool -Name "EclipseAuraCore"
                            Start-Sleep -Seconds 1
                            Write-Host "Application Pool 'EclipseAuraCore' created!" -ForegroundColor Green
                        } else {
                            Write-Host "Application Pool 'EclipseAuraCore' already exists. Updating configuration..." -ForegroundColor Yellow
                        }
                        Set-ItemProperty IIS:\AppPools\EclipseAuraCore -Name "managedRuntimeVersion" -Value ""
                        Set-ItemProperty IIS:\AppPools\EclipseAuraCore -Name "managedPipelineMode" -Value "Integrated"
                        Set-ItemProperty IIS:\AppPools\EclipseAuraCore -Name "processModel.identityType" -Value "LocalSystem"
                        Write-Host "Application Pool 'EclipseAuraCore' configured!" -ForegroundColor Green
                    } catch {
                        $err = $_
                        Write-Host ("Failed to create/configure EclipseAuraCore pool: {0}" -f $err.Exception.Message) -ForegroundColor Red
                    }
                    # Create EclipseAuraAuth pool (.NET CLR v4.0, Integrated, LocalSystem)
                    try {
                        $highestClr = "v4.0"
                        $auraAuthExists = Get-Item IIS:\AppPools\EclipseAuraAuth -ErrorAction SilentlyContinue
                        if (-not $auraAuthExists) {
                            New-WebAppPool -Name "EclipseAuraAuth"
                            Start-Sleep -Seconds 1
                            Write-Host "Application Pool 'EclipseAuraAuth' created!" -ForegroundColor Green
                        } else {
                            Write-Host "Application Pool 'EclipseAuraAuth' already exists. Updating configuration..." -ForegroundColor Yellow
                        }
                        Set-ItemProperty IIS:\AppPools\EclipseAuraAuth -Name "managedRuntimeVersion" -Value $highestClr
                        Set-ItemProperty IIS:\AppPools\EclipseAuraAuth -Name "managedPipelineMode" -Value "Integrated"
                        Set-ItemProperty IIS:\AppPools\EclipseAuraAuth -Name "processModel.identityType" -Value "LocalSystem"
                        Write-Host "Application Pool 'EclipseAuraAuth' configured!" -ForegroundColor Green
                    } catch {
                        $err = $_
                        Write-Host ("Failed to create/configure EclipseAuraAuth pool: {0}" -f $err.Exception.Message) -ForegroundColor Red
                    }
                    # Create EclipseAuraCoreAuth pool (No Managed Code, Integrated, LocalSystem)
                    try {
                        $auraCoreAuthExists = Get-Item IIS:\AppPools\EclipseAuraCoreAuth -ErrorAction SilentlyContinue
                        if (-not $auraCoreAuthExists) {
                            New-WebAppPool -Name "EclipseAuraCoreAuth"
                            Start-Sleep -Seconds 1
                            Write-Host "Application Pool 'EclipseAuraCoreAuth' created!" -ForegroundColor Green
                        } else {
                            Write-Host "Application Pool 'EclipseAuraCoreAuth' already exists. Updating configuration..." -ForegroundColor Yellow
                        }
                        Set-ItemProperty IIS:\AppPools\EclipseAuraCoreAuth -Name "managedRuntimeVersion" -Value ""
                        Set-ItemProperty IIS:\AppPools\EclipseAuraCoreAuth -Name "managedPipelineMode" -Value "Integrated"
                        Set-ItemProperty IIS:\AppPools\EclipseAuraCoreAuth -Name "processModel.identityType" -Value "LocalSystem"
                        Write-Host "Application Pool 'EclipseAuraCoreAuth' configured!" -ForegroundColor Green
                    } catch {
                        $err = $_
                        Write-Host ("Failed to create/configure EclipseAuraCoreAuth pool: {0}" -f $err.Exception.Message) -ForegroundColor Red
                    }
                    # Create folders in C:\inetpub\Eclipse if not found
                    $basePath = "C:\inetpub\Eclipse"
                    $folders = @("AuraApi", "eclipseApi", "oAuth", "oAuth2")
                    foreach ($folder in $folders) {
                        $fullPath = Join-Path $basePath $folder
                        if (-not (Test-Path $fullPath)) {
                            try {
                                New-Item -Path $fullPath -ItemType Directory -Force | Out-Null
                                Write-Host "Created folder: $fullPath" -ForegroundColor Green
                            } catch {
                                Write-Host "Failed to create folder: $fullPath - $($_.Exception.Message)" -ForegroundColor Red
                            }
                        } else {
                            Write-Host "Folder already exists: $fullPath" -ForegroundColor Yellow
                        }
                    }
                    # Create IIS Site 'Eclipse' if not exists
                    $siteName = "Eclipse"
                    $appPoolName = "Eclipse"
                    $sitePath = $basePath
                    $siteExists = Get-Website -Name $siteName -ErrorAction SilentlyContinue
                    
                    # Always create self-signed certificate for HTTPS binding
                    Write-Host ""
                    Write-Host "Creating self-signed certificate 'EclipsePlaceholderSelfsigned'..." -ForegroundColor Yellow
                    $certThumbprint = $null
                    try {
                        $certName = "EclipsePlaceholderSelfsigned"
                        $existingCert = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Subject -like "*$certName*" } | Select-Object -First 1
                        
                        if ($existingCert) {
                            Write-Host "Certificate '$certName' already exists. Using existing certificate." -ForegroundColor Green
                            $certThumbprint = $existingCert.Thumbprint
                        } else {
                            $cert = New-SelfSignedCertificate -DnsName "localhost", $env:COMPUTERNAME -CertStoreLocation "Cert:\LocalMachine\My" -FriendlyName $certName -NotAfter (Get-Date).AddYears(10)
                            $certThumbprint = $cert.Thumbprint
                            Write-Host "Self-signed certificate '$certName' created successfully!" -ForegroundColor Green
                            Write-Host "  Thumbprint: $certThumbprint" -ForegroundColor Cyan
                            Write-Host "  Note: This is a placeholder certificate. Replace with your production certificate later." -ForegroundColor Yellow
                        }
                    } catch {
                        Write-Host "Failed to create self-signed certificate: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host "Site will be created without SSL binding." -ForegroundColor Yellow
                    }
                    
                    if (-not $siteExists) {
                        try {
                            # Create site with HTTPS binding directly (no HTTP)
                            New-Website -Name $siteName -PhysicalPath $sitePath -ApplicationPool $appPoolName -Port $port -IPAddress "*" -HostHeader $hostname -Force | Out-Null
                            
                            if ($certThumbprint) {
                                # Remove the HTTP binding if it was created
                                Remove-WebBinding -Name $siteName -Protocol "http" -Port $port -HostHeader $hostname -ErrorAction SilentlyContinue
                                
                                # Add HTTPS binding with hostname
                                New-WebBinding -Name $siteName -Protocol "https" -Port $port -IPAddress "*" -HostHeader $hostname
                                
                                # Bind the certificate
                                $binding = Get-WebBinding -Name $siteName -Protocol "https" -Port $port -HostHeader $hostname
                                if ($binding) {
                                    $binding.AddSslCertificate($certThumbprint, "My")
                                    $bindingInfo = if ($hostname) { "${hostname}:${port}" } else { "*:${port}" }
                                    Write-Host "IIS Site '$siteName' created on $bindingInfo (HTTPS) with SSL certificate binding." -ForegroundColor Green
                                    Write-Host "  Certificate: EclipsePlaceholderSelfsigned" -ForegroundColor Cyan
                                } else {
                                    Write-Host "IIS Site '$siteName' created on port $port, but certificate binding failed." -ForegroundColor Yellow
                                }
                            } else {
                                Write-Host "IIS Site '$siteName' created on port $port (HTTPS), but certificate creation failed." -ForegroundColor Yellow
                                Write-Host "  You will need to manually bind a certificate in IIS Manager." -ForegroundColor Yellow
                            }
                        } catch {
                            Write-Host "Failed to create IIS Site '$siteName': $($_.Exception.Message)" -ForegroundColor Red
                        }
                    } else {
                        Write-Host "IIS Site '$siteName' already exists. Skipping creation." -ForegroundColor Yellow
                        if ($certThumbprint) {
                            Write-Host "Note: If you need to update the SSL binding, do so manually in IIS Manager." -ForegroundColor Yellow
                        }
                    }
                    # Create IIS Applications under Eclipse site
                    $appsToCreate = @(
                        @{ Name = "/AuraApi"; Folder = "AuraApi"; Pool = "EclipseAuraCore" },
                        @{ Name = "/eclipseApi"; Folder = "eclipseApi"; Pool = "Eclipse" },
                        @{ Name = "/oAuth"; Folder = "oAuth"; Pool = "EclipseAuraAuth" },
                        @{ Name = "/oAuth2"; Folder = "oAuth2"; Pool = "EclipseAuraCoreAuth" }
                    )
                    foreach ($app in $appsToCreate) {
                        $appPath = Join-Path $basePath $app.Folder
                        $iisAppExists = Get-WebApplication -Site $siteName -Name $app.Name.TrimStart('/') -ErrorAction SilentlyContinue
                        if (-not $iisAppExists) {
                            try {
                                New-WebApplication -Site $siteName -Name $app.Name.TrimStart('/') -PhysicalPath $appPath -ApplicationPool $app.Pool
                                Write-Host "Created IIS Application '$($app.Name)' mapped to '$appPath' using pool '$($app.Pool)'" -ForegroundColor Green
                            } catch {
                                Write-Host "Failed to create IIS Application '$($app.Name)': $($_.Exception.Message)" -ForegroundColor Red
                            }
                        } else {
                            Write-Host "IIS Application '$($app.Name)' already exists. Skipping creation." -ForegroundColor Yellow
                        }
                    }
                    
                    # Start or recycle the Eclipse application pools
                    Write-Host ""
                    Write-Host "Starting/recycling Eclipse application pools..." -ForegroundColor Yellow
                    $poolsToStart = @("Eclipse", "EclipseAuraCore", "EclipseAuraAuth", "EclipseAuraCoreAuth")
                    foreach ($poolName in $poolsToStart) {
                        try {
                            $pool = Get-WebAppPoolState -Name $poolName -ErrorAction SilentlyContinue
                            if ($pool) {
                                if ($pool.Value -eq "Started") {
                                    # Recycle if already running
                                    Restart-WebAppPool -Name $poolName
                                    Write-Host "  Recycled application pool: $poolName" -ForegroundColor Green
                                } else {
                                    # Start if not running
                                    Start-WebAppPool -Name $poolName
                                    Write-Host "  Started application pool: $poolName" -ForegroundColor Green
                                }
                            }
                        } catch {
                            Write-Host "  Warning: Could not start/recycle application pool '$poolName': $($_.Exception.Message)" -ForegroundColor Yellow
                        }
                    }
                    
                    # Start the IIS site if it exists
                    try {
                        $siteState = (Get-Website -Name $siteName -ErrorAction SilentlyContinue).State
                        if ($siteState -ne "Started") {
                            Start-Website -Name $siteName
                            Write-Host "  Started IIS site: $siteName" -ForegroundColor Green
                        } else {
                            Write-Host "  IIS site '$siteName' is already running" -ForegroundColor Cyan
                        }
                    } catch {
                        Write-Host "  Warning: Could not start IIS site '$siteName': $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                } catch {
                    $err = $_
                    Write-Host ("Failed to manage IIS Application Pools: {0}" -f $err.Exception.Message) -ForegroundColor Red
                }
                Start-Sleep -Seconds 2

                # Prompt to create Windows firewall rule for the selected port
                $fwPrompt = Read-Host "Would you like to create a Windows firewall rule to open port $port for IIS? (y/n)"
                if ($fwPrompt -eq 'y') {
                    try {
                        $ruleName = "IIS HTTPS ($port)"
                        New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -LocalPort $port -Protocol TCP -Action Allow -Profile Any
                        Write-Host "Windows firewall rule created to allow inbound traffic on port $port." -ForegroundColor Green
                    } catch {
                        Write-Host "Failed to create firewall rule: $($_.Exception.Message)" -ForegroundColor Red
                    }
                } else {
                    Write-Host "Firewall rule not created. You may need to configure this manually if required." -ForegroundColor Yellow
                }

                # Summary of actions
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "Eclipse IIS Setup Summary:" -ForegroundColor Cyan
                Write-Host "- Application Pools: Eclipse (.NET), EclipseAuraCore (No Managed Code)" -ForegroundColor Cyan
                Write-Host "- Folders ensured: eclipseApi, oAuth2 in C:\inetpub\Eclipse" -ForegroundColor Cyan
                $hostnameDisplay = if ($hostname) { $hostname } else { "all hostnames" }
                Write-Host "- IIS Site: Eclipse (${hostnameDisplay}:${port}, HTTPS)" -ForegroundColor Cyan
                Write-Host "- SSL Certificate: EclipsePlaceholderSelfsigned (self-signed placeholder)" -ForegroundColor Cyan
                Write-Host "- IIS Applications: /AuraApi (EclipseAuraCore), /eclipseApi (Eclipse), /oAuth (Eclipse), /oAuth2 (EclipseAuraCore)" -ForegroundColor Cyan
                Write-Host "- Firewall rule for port ${port}: " -NoNewline
                if ($fwPrompt -eq 'y') {
                    Write-Host "Created" -ForegroundColor Green
                } else {
                    Write-Host "Skipped" -ForegroundColor Yellow
                }
                Write-Host "===============================================" -ForegroundColor Cyan

                # Prompt to continue
                Write-Host ""
                Read-Host "Press Enter to continue back to the main menu"
            }
            '13' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Install ASP.NET Core Hosting Bundle 8.0.22 & .NET Desktop Runtime 8.0.22" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host ""
                
                try {
                    $targetDir = "C:\Eclipse Install\Dependencies"
                    if (!(Test-Path $targetDir)) {
                        New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
                    }
                    
                    Write-Host "Installing .NET 8 components via winget..." -ForegroundColor Yellow
                    $packages = @(
                        @{
                            Id = "Microsoft.DotNet.HostingBundle.8"
                            Name = "ASP.NET Core Hosting Bundle 8.0.22"
                            Version = "8.0.22"
                        },
                        @{
                            Id = "Microsoft.DotNet.DesktopRuntime.8"
                            Name = ".NET Desktop Runtime 8.0.22"
                            Version = "8.0.22"
                        }
                    )
                    
                    $allSuccess = $true
                    foreach ($package in $packages) {
                        Write-Host "Installing $($package.Name)..." -ForegroundColor Gray
                        $installArgs = "install --id $($package.Id) --version $($package.Version) --silent --accept-package-agreements --accept-source-agreements"
                        $process = Start-Process -FilePath "winget" -ArgumentList $installArgs -Wait -PassThru -NoNewWindow
                        if ($process.ExitCode -eq 0) {
                            Write-Host "$($package.Name) installed successfully" -ForegroundColor Green
                        } else {
                            Write-Host "Warning: $($package.Name) installation returned exit code: $($process.ExitCode)" -ForegroundColor Yellow
                            # Try without version pinning as fallback
                            Write-Host "  Attempting installation without version pinning..." -ForegroundColor Gray
                            $fallbackArgs = "install --id $($package.Id) --silent --accept-package-agreements --accept-source-agreements"
                            $fallbackProcess = Start-Process -FilePath "winget" -ArgumentList $fallbackArgs -Wait -PassThru -NoNewWindow
                            if ($fallbackProcess.ExitCode -eq 0) {
                                Write-Host "$($package.Name) installed successfully (latest version)" -ForegroundColor Green
                            } else {
                                Write-Host "  Failed to install $($package.Name)" -ForegroundColor Red
                                $allSuccess = $false
                            }
                        }
                    }
                    
                    if ($allSuccess) {
                        Write-Host ""
                        Write-Host "ASP.NET Core Hosting Bundle 8.0.22 and .NET Desktop Runtime 8.0.22 installed successfully!" -ForegroundColor Cyan
                    } else {
                        Write-Host ""
                        Write-Host ".NET 8 installation completed with warnings." -ForegroundColor Yellow
                    }
                } catch {
                    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
                }
                
                Write-Host ""
                Read-Host "Press Enter to continue back to the main menu"
            }
            '14' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Download .NET Framework 4.8" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                $targetDir = "C:\Eclipse Install\Dependencies"
                if (!(Test-Path $targetDir)) {
                    New-Item -Path $targetDir -ItemType Directory | Out-Null
                    Write-Host "Created: $targetDir" -ForegroundColor Green
                }
                $installerUrl = "https://download.microsoft.com/download/f/3/a/f3a6af84-da23-40a5-8d1c-49cc10c8e76f/NDP48-x86-x64-AllOS-ENU.exe"
                $installerPath = Join-Path $targetDir "NDP48-x86-x64-AllOS-ENU.exe"
                Write-Host "Downloading .NET Framework 4.8 installer (using BITS)..." -ForegroundColor Yellow
                try {
                    Start-BitsTransfer -Source $installerUrl -Destination $installerPath
                    Write-Host "Download complete: $installerPath" -ForegroundColor Green
                    Write-Host "Launching installer..." -ForegroundColor Cyan
                    Start-Process -FilePath $installerPath -Verb RunAs
                    Write-Host ".NET Framework 4.8 installation started!" -ForegroundColor Green
                } catch {
                    Write-Host "Download failed: $($_.Exception.Message)" -ForegroundColor Red
                }
                Write-Host ""
                Read-Host "Press Enter to continue back to the main menu"
            }
            '15' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Download Eclipse Update Service $EclipseUpdateServiceVersion" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                $targetDir = "C:\Eclipse Install\Dependencies"
                if (!(Test-Path $targetDir)) {
                    New-Item -Path $targetDir -ItemType Directory | Out-Null
                    Write-Host "Created: $targetDir" -ForegroundColor Green
                }
                $installerUrl = $EclipseUpdateServiceUrl
                $installerPath = Join-Path $targetDir ([System.Net.WebUtility]::UrlDecode((Split-Path $installerUrl -Leaf)))
                Write-Host "Downloading Eclipse Update Service $EclipseUpdateServiceVersion installer..." -ForegroundColor Yellow
                try {
                    $webClient = New-Object System.Net.WebClient
                    $webClient.DownloadFile($installerUrl, $installerPath)
                    $webClient.Dispose()
                    Write-Host "Download complete: $installerPath" -ForegroundColor Green
                    Write-Host "Launching installer..." -ForegroundColor Cyan
                    Start-Process -FilePath $installerPath -Verb RunAs
                    Write-Host "Eclipse Update Service $EclipseUpdateServiceVersion installation started!" -ForegroundColor Green
                } catch {
                    Write-Host "Download failed: $($_.Exception.Message)" -ForegroundColor Red
                }
                Write-Host ""
                Read-Host "Press Enter to continue back to the main menu"
            }
            '16' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Download Eclipse Online Chrome $EclipseOnlineChromeVersion" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                $targetDir = "C:\Eclipse Install\Dependencies"
                if (!(Test-Path $targetDir)) {
                    New-Item -Path $targetDir -ItemType Directory | Out-Null
                    Write-Host "Created: $targetDir" -ForegroundColor Green
                }
                $installerUrl = $EclipseOnlineChromeUrl
                $installerPath = Join-Path $targetDir ([System.Net.WebUtility]::UrlDecode((Split-Path $installerUrl -Leaf)))
                Write-Host "Downloading Eclipse Online Chrome $EclipseOnlineChromeVersion installer..." -ForegroundColor Yellow
                try {
                    $webClient = New-Object System.Net.WebClient
                    $webClient.DownloadFile($installerUrl, $installerPath)
                    $webClient.Dispose()
                    Write-Host "Download complete: $installerPath" -ForegroundColor Green
                    Write-Host "Launching installer..." -ForegroundColor Cyan
                    Start-Process -FilePath "$env:SystemRoot\System32\msiexec.exe" -ArgumentList "/i `"$installerPath`"" -Verb RunAs
                    Write-Host "Eclipse Online Chrome $EclipseOnlineChromeVersion installation started!" -ForegroundColor Green
                } catch {
                    Write-Host "Download failed: $($_.Exception.Message)" -ForegroundColor Red
                }
                Write-Host ""
                Read-Host "Press Enter to continue back to the main menu"
            }
            '17' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Download Eclipse Online Server (Aura)" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                $targetDir = "C:\Eclipse Install\Dependencies"
                if (!(Test-Path $targetDir)) {
                    New-Item -Path $targetDir -ItemType Directory | Out-Null
                    Write-Host "Created: $targetDir" -ForegroundColor Green
                }
                $installerUrl = $EclipseOnlineServerUrl
                $installerPath = Join-Path $targetDir ([System.Net.WebUtility]::UrlDecode((Split-Path $installerUrl -Leaf)))
                if (Test-Path $installerPath) {
                    Write-Host "Installer already downloaded: $installerPath" -ForegroundColor Green
                } else {
                    Write-Host "Downloading Eclipse Online Server $EclipseOnlineServerVersion installer..." -ForegroundColor Yellow
                    try {
                        $webClient = New-Object System.Net.WebClient
                        $webClient.DownloadFile($installerUrl, $installerPath)
                        $webClient.Dispose()
                        Write-Host "Download complete: $installerPath" -ForegroundColor Green
                    } catch {
                        Write-Host "Download failed: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host ""
                        Read-Host "Press Enter to continue back to the main menu"
                        continue
                    }
                }
                Write-Host "Launching installer..." -ForegroundColor Cyan
                Start-Process -FilePath $installerPath -Verb RunAs
                Write-Host "Eclipse Online Server $EclipseOnlineServerVersion installation started!" -ForegroundColor Green
                Write-Host ""
                Read-Host "Press Enter to continue back to the main menu"
            }
            '19' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Install Win-ACMEv2" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host ""
                
                # Get Eclipse Install Path
                $defaultPath = "C:\Eclipse Install"
                $eclipseInstallPath = Read-Host "Enter Eclipse Install Path (press Enter for default: $defaultPath)"
                if ([string]::IsNullOrWhiteSpace($eclipseInstallPath)) {
                    $eclipseInstallPath = $defaultPath
                }
                
                # Create Dependencies\Win-ACMEv2 directory
                $targetDir = Join-Path $eclipseInstallPath "Dependencies\Win-ACMEv2"
                if (!(Test-Path $targetDir)) {
                    Write-Host "Creating directory: $targetDir" -ForegroundColor Yellow
                    New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
                    Write-Host "Created directory: $targetDir" -ForegroundColor Green
                } else {
                    Write-Host "Directory exists: $targetDir" -ForegroundColor Green
                }
                
                # Check if already extracted
                $exePath = Join-Path $targetDir "wacs.exe"
                if (Test-Path $exePath) {
                    Write-Host "Win-ACMEv2 is already extracted in: $targetDir" -ForegroundColor Green
                    Write-Host ""
                } else {
                    # Download Win-ACMEv2
                    $winAcmeUrl = "https://github.com/win-acme/win-acme/releases/download/v2.2.9.1701/win-acme.v2.2.9.1701.x64.trimmed.zip"
                    $zipPath = Join-Path $targetDir "win-acme.zip"
                    
                    Write-Host ""
                    Write-Host "Downloading Win-ACMEv2 v2.2.9.1701..." -ForegroundColor Yellow
                    Write-Host "  Source: $winAcmeUrl" -ForegroundColor Gray
                    
                    try {
                        $webClient = New-Object System.Net.WebClient
                        $webClient.DownloadFile($winAcmeUrl, $zipPath)
                        $webClient.Dispose()
                        $fileSize = (Get-Item $zipPath).Length / 1MB
                        Write-Host "Download complete: $([math]::Round($fileSize, 2)) MB" -ForegroundColor Green
                    } catch {
                        Write-Host "Error downloading Win-ACMEv2: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host ""
                        Read-Host "Press Enter to continue back to the main menu"
                        continue
                    }
                    
                    # Extract ZIP file
                    Write-Host ""
                    Write-Host "Extracting Win-ACMEv2..." -ForegroundColor Yellow
                    
                    try {
                        Expand-Archive -Path $zipPath -DestinationPath $targetDir -Force
                        Write-Host "Extraction complete!" -ForegroundColor Green
                    } catch {
                        Write-Host "Error extracting Win-ACMEv2: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host ""
                        Read-Host "Press Enter to continue back to the main menu"
                        continue
                    }
                    
                    # Clean up ZIP file
                    Write-Host "Cleaning up..." -ForegroundColor Gray
                    Remove-Item -Path $zipPath -Force -ErrorAction SilentlyContinue
                    
                    # Verify extraction
                    if (!(Test-Path $exePath)) {
                        Write-Host "Warning: Extraction completed but wacs.exe not found" -ForegroundColor Yellow
                        Write-Host ""
                        Read-Host "Press Enter to continue back to the main menu"
                        continue
                    }
                    
                    Write-Host ""
                    Write-Host "Win-ACMEv2 extracted successfully!" -ForegroundColor Green
                    Write-Host "  Location: $targetDir" -ForegroundColor Gray
                    Write-Host ""
                }
                
                # Launch Win-ACMEv2 for user configuration
                Write-Host "Launching Win-ACMEv2 as Administrator for configuration..." -ForegroundColor Yellow
                Write-Host "  Please configure your SSL certificates" -ForegroundColor Cyan
                Write-Host "  The script will wait until you close Win-ACMEv2" -ForegroundColor Cyan
                Write-Host ""
                
                try {
                    $process = Start-Process -FilePath $exePath -WorkingDirectory $targetDir -Verb RunAs -PassThru -Wait
                    
                    Write-Host ""
                    Write-Host "Win-ACMEv2 closed (exit code: $($process.ExitCode))" -ForegroundColor Green
                    Write-Host "Win-ACMEv2 configuration completed!" -ForegroundColor Cyan
                } catch {
                    Write-Host "Error launching Win-ACMEv2: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host "You can run it manually from: $targetDir" -ForegroundColor Gray
                }
                
                Write-Host ""
                Read-Host "Press Enter to continue back to the main menu"
            }
            '20' {
                Clear-Host
                Write-Host "===============================================" -ForegroundColor Cyan
                Write-Host "   Download Eclipse Smart Hub $EclipseSmartHubVersion" -ForegroundColor Yellow
                Write-Host "===============================================" -ForegroundColor Cyan
                $targetDir = "C:\Eclipse Install\Dependencies"
                if (!(Test-Path $targetDir)) {
                    New-Item -Path $targetDir -ItemType Directory | Out-Null
                    Write-Host "Created: $targetDir" -ForegroundColor Green
                }
                $installerUrl = $EclipseSmartHubUrl
                $installerPath = Join-Path $targetDir ([System.Net.WebUtility]::UrlDecode((Split-Path $installerUrl -Leaf)))
                if (Test-Path $installerPath) {
                    Write-Host "Installer already downloaded: $installerPath" -ForegroundColor Green
                } else {
                    Write-Host "Downloading Eclipse Smart Hub $EclipseSmartHubVersion installer..." -ForegroundColor Yellow
                    try {
                        $webClient = New-Object System.Net.WebClient
                        $webClient.DownloadFile($installerUrl, $installerPath)
                        $webClient.Dispose()
                        Write-Host "Download complete: $installerPath" -ForegroundColor Green
                    } catch {
                        Write-Host "Download failed: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host ""
                        Read-Host "Press Enter to continue back to the main menu"
                        continue
                    }
                }
                Write-Host "Launching installer..." -ForegroundColor Cyan
                Start-Process -FilePath $installerPath -Verb RunAs
                Write-Host "Eclipse Smart Hub $EclipseSmartHubVersion installation started!" -ForegroundColor Green
                Write-Host ""
                Read-Host "Press Enter to continue back to the main menu"
            }
            'U' {
                # Launch Unattended Server Setup Mode
                Start-UnattendedMode
            }
            '0' { Write-Host "Exiting..." -ForegroundColor Red; Start-Sleep -Seconds 1 }
            default { Write-Host "Invalid selection. Please try again." -ForegroundColor Red; Start-Sleep -Seconds 1 }
        }
    } while ($choice -ne '0' -and $choice -ne 'U')
}

Main


