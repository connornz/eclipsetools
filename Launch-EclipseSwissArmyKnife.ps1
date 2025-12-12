# Eclipse Swiss Army Knife Launcher
# Automatically checks for and downloads the latest version from GitHub
# Created by Connor Brown for Eclipse Support Team

# Configuration
$GitHubRawUrl = "https://raw.githubusercontent.com/connornz/eclipsetools/refs/heads/main/EclipseSwissArmyKnife.ps1"
$ScriptName = "EclipseSwissArmyKnife.ps1"
$ScriptPath = Join-Path $PSScriptRoot $ScriptName

# Function to extract version from script content
function Get-ScriptVersion {
    param(
        [string]$ScriptContent
    )
    
    # Look for: $ScriptVersion = "2.1.0" or $ScriptVersion = '2.1.0'
    $versionPattern = '\$ScriptVersion\s*=\s*["'']([\d.]+)["'']'
    
    if ($ScriptContent -match $versionPattern) {
        return $Matches[1]
    }
    
    return $null
}

# Function to compare semantic versions
function Compare-Versions {
    param(
        [string]$Version1,
        [string]$Version2
    )
    
    if ([string]::IsNullOrWhiteSpace($Version1)) { return -1 }
    if ([string]::IsNullOrWhiteSpace($Version2)) { return 1 }
    
    # Remove 'v' prefix if present
    $Version1 = $Version1 -replace '^v', ''
    $Version2 = $Version2 -replace '^v', ''
    
    $v1Parts = $Version1.Split('.')
    $v2Parts = $Version2.Split('.')
    
    # Compare each part numerically
    $maxLength = [Math]::Max($v1Parts.Length, $v2Parts.Length)
    
    for ($i = 0; $i -lt $maxLength; $i++) {
        $v1Part = if ($i -lt $v1Parts.Length) { [int]$v1Parts[$i] } else { 0 }
        $v2Part = if ($i -lt $v2Parts.Length) { [int]$v2Parts[$i] } else { 0 }
        
        if ($v1Part -lt $v2Part) { return -1 }
        if ($v1Part -gt $v2Part) { return 1 }
    }
    
    return 0  # Versions are equal
}

# Function to download script from GitHub
function Get-RemoteScript {
    param(
        [string]$Url
    )
    
    try {
        Write-Host "Checking for updates..." -ForegroundColor Yellow
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "EclipseSwissArmyKnife-Launcher")
        $scriptContent = $webClient.DownloadString($Url)
        $webClient.Dispose()
        return $scriptContent
    } catch {
        Write-Host "Warning: Could not check for updates. Error: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }
}

# Main execution
try {
    # Check if local script exists
    if (-not (Test-Path $ScriptPath)) {
        Write-Host "Local script not found. Downloading from GitHub..." -ForegroundColor Yellow
        $remoteScript = Get-RemoteScript -Url $GitHubRawUrl
        
        if ($remoteScript) {
            $remoteScript | Out-File -FilePath $ScriptPath -Encoding UTF8
            Write-Host "Script downloaded successfully!" -ForegroundColor Green
        } else {
            Write-Host "ERROR: Could not download script. Please check your internet connection." -ForegroundColor Red
            exit 1
        }
    } else {
        # Read local script version
        $localScriptContent = Get-Content -Path $ScriptPath -Raw -ErrorAction Stop
        $localVersion = Get-ScriptVersion -ScriptContent $localScriptContent
        
        if ($localVersion) {
            Write-Host "Current local version: $localVersion" -ForegroundColor Gray
        } else {
            Write-Host "Warning: Could not determine local script version." -ForegroundColor Yellow
        }
        
        # Check for remote version
        $remoteScript = Get-RemoteScript -Url $GitHubRawUrl
        
        if ($remoteScript) {
            $remoteVersion = Get-ScriptVersion -ScriptContent $remoteScript
            
            if ($remoteVersion) {
                Write-Host "Latest available version: $remoteVersion" -ForegroundColor Gray
                
                # Compare versions
                $comparison = Compare-Versions -Version1 $localVersion -Version2 $remoteVersion
                
                if ($comparison -lt 0) {
                    # Remote is newer
                    Write-Host "New version available! Updating..." -ForegroundColor Green
                    $remoteScript | Out-File -FilePath $ScriptPath -Encoding UTF8
                    Write-Host "Script updated to version $remoteVersion" -ForegroundColor Green
                } elseif ($comparison -eq 0) {
                    Write-Host "You have the latest version." -ForegroundColor Green
                } else {
                    Write-Host "Local version is newer than remote (unusual). Using local version." -ForegroundColor Yellow
                }
            } else {
                Write-Host "Warning: Could not determine remote script version. Using local script." -ForegroundColor Yellow
            }
        } else {
            Write-Host "Using local script (could not check for updates)." -ForegroundColor Yellow
        }
    }
    
    # Execute the script
    Write-Host ""
    Write-Host "Launching Eclipse Swiss Army Knife..." -ForegroundColor Cyan
    Write-Host ""
    
    & $ScriptPath @args
    
} catch {
    Write-Host ""
    Write-Host "ERROR: An error occurred while launching the script." -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please ensure:" -ForegroundColor Yellow
    Write-Host "  - You have internet connectivity (for update checks)" -ForegroundColor Yellow
    Write-Host "  - The script file exists and is accessible" -ForegroundColor Yellow
    Write-Host "  - You have appropriate permissions" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}


