# Usage:
#   .\Download-SSMS.ps1                                   # Install SSMS, download components, and copy to installation
#   .\Download-SSMS.ps1 -DownloadPath "C:\Temp"           # Custom download location
#   .\Download-SSMS.ps1 -Verbose                          # Verbose output
#
# This script (REQUIRES ADMINISTRATOR PRIVILEGES):
# 1. Temporarily sets execution policy to allow script execution
# 2. Runs as administrator and closes Visual Studio Installer instances
# 3. Downloads and installs SSMS using vs_SSMS.exe --arch x64 --quiet
# 4. Downloads VSIX components from the SSMS channel manifest 
# 5. Extracts VSIX files and copies them to the SSMS installation directory
# 6. Uses VSSetup PowerShell module to detect SSMS installation path (with fallback methods)
# 7. Restores original execution policy

param(
    [string]$DownloadPath = ".\SSMS21-Downloads",
    [string]$ExtractPath = ".\Extracted",
    [switch]$Verbose
)

# Store original execution policies for restoration
$script:OriginalExecutionPolicies = @{}

# Function to write verbose output
function Write-VerboseOutput {
    param([string]$Message)
    if ($Verbose) {
        Write-Host $Message -ForegroundColor Green
    }
}


# Function to save current execution policies
function Save-ExecutionPolicies {
    try {
        Write-VerboseOutput "Saving current execution policies..."
        
        # Get execution policies for different scopes
        $scopes = @('Process', 'CurrentUser', 'LocalMachine')
        
        foreach ($scope in $scopes) {
            try {
                $policy = Get-ExecutionPolicy -Scope $scope -ErrorAction SilentlyContinue
                $script:OriginalExecutionPolicies[$scope] = $policy
                Write-VerboseOutput "Current $scope execution policy: $policy"
            }
            catch {
                Write-VerboseOutput "Could not get $scope execution policy: $($_.Exception.Message)"
            }
        }
    }
    catch {
        Write-VerboseOutput "Error saving execution policies: $($_.Exception.Message)"
    }
}

# Function to set execution policies to allow script execution
function Set-ExecutionPoliciesForScript {
    try {
        Write-Host "Temporarily setting execution policies to allow script execution..." -ForegroundColor Yellow
        
        # Set execution policy for current process (most permissive, least risky)
        Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue
        Write-VerboseOutput "Set Process execution policy to Bypass"
        
        # Try to set CurrentUser policy if Process scope isn't sufficient
        try {
            Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction SilentlyContinue
            Write-VerboseOutput "Set CurrentUser execution policy to RemoteSigned"
        }
        catch {
            Write-VerboseOutput "Could not set CurrentUser execution policy: $($_.Exception.Message)"
        }
        
        Write-Host "Execution policies configured for script execution." -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to set execution policies: $($_.Exception.Message)"
        Write-Host "You may need to manually set execution policy or run PowerShell with -ExecutionPolicy Bypass" -ForegroundColor Yellow
    }
}

# Function to restore original execution policies
function Restore-ExecutionPolicies {
    try {
        Write-VerboseOutput "Restoring original execution policies..."
        
        foreach ($scope in $script:OriginalExecutionPolicies.Keys) {
            $originalPolicy = $script:OriginalExecutionPolicies[$scope]
            
            if ($originalPolicy -and $originalPolicy -ne 'Undefined') {
                try {
                    Set-ExecutionPolicy -ExecutionPolicy $originalPolicy -Scope $scope -Force -ErrorAction SilentlyContinue
                    Write-VerboseOutput "Restored $scope execution policy to: $originalPolicy"
                }
                catch {
                    Write-VerboseOutput "Could not restore $scope execution policy: $($_.Exception.Message)"
                }
            }
        }
        
        Write-Host "Original execution policies restored." -ForegroundColor Green
    }
    catch {
        Write-VerboseOutput "Error restoring execution policies: $($_.Exception.Message)"
    }
}

# Save current execution policies and set new ones
Save-ExecutionPolicies
Set-ExecutionPoliciesForScript

# Set up trap to restore execution policies on unexpected exit
trap {
    Write-Warning "Script interrupted unexpectedly. Restoring execution policies..."
    Restore-ExecutionPolicies
    break
}

# Check if running as administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Restart script as administrator if needed
if (-not (Test-Administrator)) {
    Write-Host "This script requires administrator privileges. Restarting as administrator..." -ForegroundColor Yellow
    
    $arguments = ""
    if ($DownloadPath -ne ".\SSMS21-Downloads") { $arguments += " -DownloadPath $DownloadPath" }
    if ($ExtractPath -ne ".\Extracted") { $arguments += " -ExtractPath $ExtractPath" }
    if ($Verbose) { $arguments += " -Verbose" }
    
    Start-Process PowerShell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`"$arguments"
    exit
}

# Function to close Visual Studio Installer instances
function Stop-VisualStudioInstaller {
    Write-Host "Checking for Visual Studio Installer processes..." -ForegroundColor Yellow

    $installerProcesses = Get-Process -Name "vs_installer*", "VSIXInstaller*", "setup*" | Where-Object { $_.Path -match "\\Installer\\" } -ErrorAction SilentlyContinue

    if ($installerProcesses) {
        Write-Host "Found Visual Studio Installer processes. Closing them..." -ForegroundColor Yellow
        foreach ($process in $installerProcesses) {
            try {
                Write-VerboseOutput "Stopping process: $($process.ProcessName) (PID: $($process.Id))"
                $process.CloseMainWindow()
                Start-Sleep -Seconds 2
                
                if (-not $process.HasExited) {
                    $process.Kill()
                    Start-Sleep -Seconds 1
                }
                Write-VerboseOutput "Process $($process.ProcessName) stopped successfully"
            }
            catch {
                Write-Warning "Failed to stop process $($process.ProcessName): $($_.Exception.Message)"
            }
        }
        Write-Host "Visual Studio Installer processes closed." -ForegroundColor Green
    }
    else {
        Write-VerboseOutput "No Visual Studio Installer processes found."
    }
}

# Set default ExtractPath relative to DownloadPath if not specified
if ($ExtractPath -eq ".\Extracted") {
    $ExtractPath = Join-Path $DownloadPath "Extracted"
}

# Function to download file with progress
function Get-FileWithProgress {
    param(
        [string]$Url,
        [string]$OutputPath
    )
    
    try {
        Write-VerboseOutput "Downloading: $Url"
        Write-VerboseOutput "To: $OutputPath"
        $ProgressPreference = 'SilentlyContinue'  # Suppress progress bar in console
        Invoke-WebRequest -Uri $Url -OutFile $OutputPath -UseBasicParsing -ErrorAction Stop
        $ProgressPreference = 'Continue'  # Restore progress preference
        Write-Host "Successfully downloaded: $(Split-Path $OutputPath -Leaf)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to download $Url - Error: $($_.Exception.Message)"
        return $false
    }
}

# Function to extract VSIX files
function Expand-VsixFile {
    param(
        [string]$VsixPath,
        [string]$ExtractPath
    )
    
    try {
        Write-VerboseOutput "Extracting VSIX: $VsixPath"
        
        # Create temporary extraction directory
        $tempDir = Join-Path $env:TEMP "vsix_temp_$(Get-Random)"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        
        # Extract VSIX (it's a ZIP file)
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($VsixPath, $tempDir)
        
        # Read manifest.json
        $manifestPath = Join-Path $tempDir "manifest.json"
        if (-not (Test-Path $manifestPath)) {
            Write-Warning "No manifest.json found in $VsixPath"
            Remove-Item $tempDir -Recurse -Force
            return $false
        }
        
        $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
        Write-VerboseOutput "Read manifest for: $($manifest.id)"
        
        # Get extension directory from manifest
        $extensionDir = $manifest.extensionDir
        if (-not $extensionDir) {
            $extensionDir = ""  # Default to root if not specified
            Write-VerboseOutput "No extensionDir specified in manifest, defaulting to root"
        }
        
        # Remove [installdir]\ prefix if present
        $cleanExtensionDir = $extensionDir -replace '^\[installdir\]\\', ''
        Write-VerboseOutput "Extension directory: $cleanExtensionDir"
        
        # Create target directory structure
        $targetDir = Join-Path $ExtractPath $cleanExtensionDir
        if (-not (Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }
        
        # Process files from manifest
        if ($manifest.files) {
            foreach ($file in $manifest.files) {
                $sourceFile = Join-Path $tempDir $file.fileName
                
                if (Test-Path $sourceFile) {
                    # Create relative directory structure in target
                    $relativePath = ($file.fileName -replace '/Contents/', '')  # Remove leading /Content/ if present
                    $targetFile = Join-Path $targetDir $relativePath
                    
                    # Create directory if it doesn't exist
                    $targetFileDir = Split-Path $targetFile -Parent
                    if (-not (Test-Path $targetFileDir)) {
                        New-Item -ItemType Directory -Path $targetFileDir -Force | Out-Null
                    }
                    
                    # Copy file
                    Copy-Item $sourceFile $targetFile -Force
                    Write-VerboseOutput "Copied: $relativePath"
                }
                else {
                    Write-VerboseOutput "File not found in VSIX: $($file.fileName)"
                }
            }
        }
        
        # Also copy any additional files from the VSIX root
        $vsixFiles = Get-ChildItem $tempDir -File
        foreach ($file in $vsixFiles) {
            if ($file.Name -notin @("manifest.json", "catalog.json")) {
                $targetFile = Join-Path $targetDir $file.Name
                Copy-Item $file.FullName $targetFile -Force
                Write-VerboseOutput "Copied additional file: $($file.Name)"
            }
        }
        
        # Clean up temp directory
        Remove-Item $tempDir -Recurse -Force
        
        Write-Host "  Extracted: $(Split-Path $VsixPath -Leaf)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to extract $VsixPath - Error: $($_.Exception.Message)"
        if (Test-Path $tempDir) {
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        return $false
    }
}

# Function to extract all VSIX files
function Expand-AllVsixFiles {
    param(
        [string]$DownloadPath,
        [string]$ExtractPath
    )
    
    Write-Host "`n5. Extracting VSIX files..." -ForegroundColor Yellow
    
    # Create extract directory if it doesn't exist
    if (-not (Test-Path $ExtractPath)) {
        New-Item -ItemType Directory -Path $ExtractPath -Force | Out-Null
        Write-VerboseOutput "Created extraction directory: $ExtractPath"
    }
    
    # Find all VSIX files
    $vsixFiles = Get-ChildItem $DownloadPath -Filter "*.vsix"
    
    if ($vsixFiles.Count -eq 0) {
        Write-Warning "No VSIX files found in $DownloadPath"
        return 0
    }
    
    $extractedCount = 0
    foreach ($vsixFile in $vsixFiles) {
        if (Expand-VsixFile -VsixPath $vsixFile.FullName -ExtractPath $ExtractPath) {
            $extractedCount++
        }
    }
    
    Write-Host "Extraction Summary:" -ForegroundColor Cyan
    Write-Host "VSIX files found: $($vsixFiles.Count)" -ForegroundColor White
    Write-Host "Successfully extracted: $extractedCount" -ForegroundColor Green
    Write-Host "Extraction location: $((Resolve-Path $ExtractPath).Path)" -ForegroundColor White    
    return $extractedCount
}

# Function to find SSMS installation path using VSSetup PowerShell module
function Get-SSMSInstallationPath {
    try {
        Write-VerboseOutput "Searching for SSMS installation using VSSetup module..."
        
        # Check if VSSetup module is available, install if needed
        if (-not (Get-Module -ListAvailable -Name VSSetup)) {
            Write-Host "VSSetup module not found. Installing VSSetup module..." -ForegroundColor Yellow
            try {
                Install-Module -Name VSSetup -Force -Scope CurrentUser -AllowClobber
                Write-VerboseOutput "VSSetup module installed successfully"
            }
            catch {
                Write-Warning "Failed to install VSSetup module: $($_.Exception.Message)"
                Write-VerboseOutput "Falling back to registry and common paths method..."
                return Get-SSMSInstallationPathFallback
            }
        }
        
        # Import VSSetup module
        Import-Module VSSetup -Force
        
        # Get all Visual Studio instances
        $instances = Get-VSSetupInstance | Where-Object {$_.Product.Chip -eq "x64"}
        if ($instances.Count -eq 0) {
            Write-Warning "No Visual Studio instances found with x64 architecture"
            return Get-SSMSInstallationPathFallback
        }
        # Look for SSMS instance
        foreach ($instance in $instances) {

            $displayName = $instance.DisplayName
            $installationPath = $instance.InstallationPath
            Write-VerboseOutput "Found instance: $displayName at $installationPath"
            
            # Check if this is SSMS 21
            if ($displayName -match "SQL Server Management Studio" -and $displayName -match "21") {
                Write-Host "Found SSMS 21 installation: $displayName" -ForegroundColor Green
                Write-Host "Installation path: $installationPath" -ForegroundColor Green
                return $installationPath
            }
        }
        
        Write-Warning "SSMS 21 instance not found via VSSetup module"
        
        # Try fallback method
        Write-VerboseOutput "Trying fallback method to find SSMS installation..."
        return Get-SSMSInstallationPathFallback
    }
    catch {
        Write-VerboseOutput "VSSetup module failed: $($_.Exception.Message)"
        
        # Fallback: Try common SSMS installation paths
        Write-VerboseOutput "Trying fallback method to find SSMS installation..."
        return Get-SSMSInstallationPathFallback
    }
}

# Fallback function to find SSMS installation using registry and common paths
function Get-SSMSInstallationPathFallback {
    try {
        Write-VerboseOutput "Using fallback method to locate SSMS installation..."
        
        # Try registry first
        $registryPaths = @(
            "HKLM:\SOFTWARE\Microsoft\SQL Server Management Studio\21.0",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\SQL Server Management Studio\21.0",
            "HKLM:\SOFTWARE\Microsoft\VisualStudio\Packages",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\Packages"
        )
        
        foreach ($regPath in $registryPaths) {
            try {
                if (Test-Path $regPath) {
                    $regItems = Get-ChildItem $regPath -ErrorAction SilentlyContinue
                    foreach ($item in $regItems) {
                        $properties = Get-ItemProperty $item.PSPath -ErrorAction SilentlyContinue
                        if ($properties -and $properties.InstallDir) {
                            $installPath = $properties.InstallDir
                            if (Test-Path $installPath) {
                                $ssmsExe = Join-Path $installPath "Common7\IDE\Ssms.exe"
                                if (Test-Path $ssmsExe) {
                                    Write-Host "Found SSMS installation via registry: $installPath" -ForegroundColor Green
                                    return $installPath
                                }
                            }
                        }
                    }
                }
            }
            catch {
                Write-VerboseOutput "Registry path $regPath not accessible: $($_.Exception.Message)"
            }
        }
        
        # Try common SSMS installation paths
        $commonPaths = @(
            "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio 21",
            "${env:ProgramFiles}\Microsoft SQL Server Management Studio 21",
            "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\SQL",
            "${env:ProgramFiles}\Microsoft Visual Studio\2022\SQL",
            "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio 20",
            "${env:ProgramFiles}\Microsoft SQL Server Management Studio 20"
        )
        
        foreach ($path in $commonPaths) {
            if (Test-Path $path) {
                $ssmsExe = Join-Path $path "Common7\IDE\Ssms.exe"
                if (Test-Path $ssmsExe) {
                    Write-Host "Found SSMS installation at common path: $path" -ForegroundColor Green
                    return $path
                }
            }
        }
        
        Write-Warning "Could not locate SSMS installation directory using fallback methods"
        return $null
    }
    catch {
        Write-Warning "Fallback method failed: $($_.Exception.Message)"
        return $null
    }
}

# Function to copy extracted files to SSMS installation
function Copy-ExtractedFilesToSSMS {
    param(
        [string]$ExtractPath,
        [string]$SSMSPath
    )
    
    try {
        Write-Host "`n6. Copying extracted files to SSMS installation..." -ForegroundColor Yellow
        
        if (-not $SSMSPath) {
            Write-Warning "SSMS installation path not provided. Skipping file copy."
            return $false
        }
        
        if (-not (Test-Path $ExtractPath)) {
            Write-Warning "Extract path does not exist: $ExtractPath"
            return $false
        }
        
        # Get all subdirectories in the extract path
        $extractedDirs = Get-ChildItem $ExtractPath -Directory
        
        if ($extractedDirs.Count -eq 0) {
            Write-Warning "No extracted directories found in $ExtractPath"
            return $false
        }
        
        $copiedCount = 0
        foreach ($dir in $extractedDirs) {
            $targetPath = Join-Path $SSMSPath $dir.Name
            
            try {
                Write-VerboseOutput "Copying: $($dir.FullName) -> $targetPath"
                
                # Create target directory if it doesn't exist
                if (-not (Test-Path $targetPath)) {
                    New-Item -ItemType Directory -Path $targetPath -Force | Out-Null
                }
                
                # Copy all contents
                Copy-Item -Path "$($dir.FullName)\*" -Destination $targetPath -Recurse -Force
                Write-Host "  Copied: $($dir.Name)" -ForegroundColor Green
                $copiedCount++
            }
            catch {
                Write-Warning "Failed to copy $($dir.Name): $($_.Exception.Message)"
            }
        }
        
        Write-Host "Copy Summary:" -ForegroundColor Cyan
        Write-Host "Directories copied: $copiedCount of $($extractedDirs.Count)" -ForegroundColor Green
        Write-Host "Target location: $SSMSPath" -ForegroundColor White
        
        return $copiedCount -gt 0
    }
    catch {
        Write-Error "Error during file copy: $($_.Exception.Message)"
        return $false
    }
}

function Test-ProductPackages {
    param(
        [object]$package
    )

    # Check if package has x64 architecture for both productArch and machineArch
    $hasX64ProductArch = $package.productArch -eq "x64"
    $hasX64MachineArch = $package.machineArch -eq "x64"
    $isNeutral = $null -eq $package.productArch -and $null -eq $package.machineArch

    return (($hasX64ProductArch -and $hasX64MachineArch) -or ($hasX64ProductArch -and $null -eq $package.machineArch) -or $isNeutral)
}

# Target item IDs to download
$targetItemIds = @(
    "Microsoft.VisualStudio.MinShell.Targeted",
    "Microsoft.VisualStudio.ExtensionManager.x64",
    "Microsoft.DiagnosticsHub.Runtime.Targeted",
    "Microsoft.ServiceHub.Managed",
    "Microsoft.ServiceHub.amd64",
    "Microsoft.VisualStudio.Identity",
    "Microsoft.VisualStudio.ExtensionManager"
)

Write-Host "Starting SSMS Component Download Script" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan

# Main script execution with error handling
try {
    # Close Visual Studio Installer instances before starting
    Stop-VisualStudioInstaller

    # Step 0: Download and run SSMS installer
    Write-Host "`n0. Downloading and installing SSMS..." -ForegroundColor Yellow
    $ssmsInstallerUrl = "https://aka.ms/ssms/21/release/vs_SSMS.exe"
    $ssmsInstallerPath = Join-Path $DownloadPath "vs_SSMS.exe"

    # Create download directory if it doesn't exist
    New-Item -ItemType Directory -Path $DownloadPath -Force | Out-Null
    Write-VerboseOutput "Created download directory: $DownloadPath"


    # Download SSMS installer
    try {
        Write-Host "Downloading SSMS installer..." -ForegroundColor Yellow
        if (Get-FileWithProgress -Url $ssmsInstallerUrl -OutputPath $ssmsInstallerPath) {
            Write-Host "SSMS installer downloaded successfully" -ForegroundColor Green
        
            # Run SSMS installer
            Write-Host "Running SSMS installer with --arch x64 --quiet..." -ForegroundColor Yellow
            Write-VerboseOutput "Executing: $ssmsInstallerPath --arch x64 --quiet"
        
            $process = Start-Process -FilePath $ssmsInstallerPath -ArgumentList "--arch", "x64", "--quiet", "--wait" -Wait -PassThru -NoNewWindow
        
            if ($process.ExitCode -eq 0) {
                Write-Host "SSMS installation completed successfully" -ForegroundColor Green
            }
            else {
                Write-Warning "SSMS installer exited with code: $($process.ExitCode)"
                Write-Host "Continuing with component download..." -ForegroundColor Yellow
            }
        }
        else {
            Write-Error "Failed to download SSMS installer"
            #  exit 1
        }
    }
    catch {
        Write-Error "Error during SSMS installation: $($_.Exception.Message)"
        Sleep 5
        # exit 1
    }

    # Step 1: Get the channel manifest
    Write-Host "`n1. Fetching SSMS21 channel manifest..." -ForegroundColor Yellow
    $channelUrl = "https://aka.ms/ssms/21/release/channel"

    try {
        Write-VerboseOutput "Requesting: $channelUrl"
        $channelResponse = Invoke-RestMethod -Uri $channelUrl -Method Get
        Write-VerboseOutput "Channel manifest retrieved successfully"
    }
    catch {
        Write-Error "Failed to retrieve channel manifest from $channelUrl - Error: $($_.Exception.Message)"
        exit 1
    }

    # Step 2: Parse the channel manifest and get the first payload URL
    Write-Host "2. Parsing channel manifest for payload URL..." -ForegroundColor Yellow

    if (-not $channelResponse.channelItems -or $channelResponse.channelItems.Count -eq 0) {
        Write-Error "No channel items found in the manifest"
        exit 1
    }

    $firstPayload = $channelResponse.channelItems[0]
    if (-not $firstPayload.payloads -or $firstPayload.payloads.Count -eq 0) {
        Write-Error "No payloads found in the first channel item"
        exit 1
    }

    $catalogUrl = $firstPayload.payloads[0].url
    Write-Host "Found catalog URL: $catalogUrl" -ForegroundColor Green
    Write-VerboseOutput "Catalog URL: $catalogUrl"

    # Step 3: Get the catalog JSON
    Write-Host "`n3. Fetching catalog JSON..." -ForegroundColor Yellow

    try {
        Write-VerboseOutput "Requesting: $catalogUrl"
        $catalogResponse = Invoke-RestMethod -Uri $catalogUrl -Method Get
        Write-VerboseOutput "Catalog JSON retrieved successfully"
    }
    catch {
        Write-Error "Failed to retrieve catalog from $catalogUrl - Error: $($_.Exception.Message)"
        exit 1
    }

    # Step 4: Filter and download target components
    Write-Host "4. Filtering components and downloading..." -ForegroundColor Yellow

    $downloadCount = 0
    $totalFound = 0

    # Process packages in the catalog
    if ($catalogResponse.packages) {
        foreach ($package in $catalogResponse.packages | Where-Object { Test-ProductPackages $_ }) {
            # Check if this package matches our target item IDs
            if ($targetItemIds -contains $package.id) {
                $totalFound++
                Write-VerboseOutput "Found target package: $($package.id)"

                if ($package.payloads) {
                    foreach ($payload in $package.payloads) {

                        Write-Host "  Found payload for: $($package.id)" -ForegroundColor Cyan
                        
                        # Generate filename
                        $fileName = "$($package.id)_$($payload.sha256).vsix"
                        if ($payload.fileName) {
                            $fileName = $payload.fileName
                        }
                        
                        $outputPath = Join-Path $DownloadPath $fileName
                        # Download the payload
                        if (Get-FileWithProgress -Url $payload.url -OutputPath $outputPath) {
                            $downloadCount++
                        }
                        
                        break # Only download first matching payload per package
                    
                    }
                }
            }
        }
    }

    # Summary
    Write-Host "`n=========================================" -ForegroundColor Cyan
    Write-Host "Download Summary:" -ForegroundColor Cyan
    Write-Host "Target components found: $totalFound" -ForegroundColor White
    Write-Host "Successfully downloaded: $downloadCount" -ForegroundColor Green
    Write-Host "Download location: $((Resolve-Path $DownloadPath).Path)" -ForegroundColor White

    if ($downloadCount -eq 0) {
        Write-Warning "No files were downloaded. This could mean:"
        Write-Warning "- The target components don't have x64 architecture versions"
        Write-Warning "- The component IDs have changed"
        Write-Warning "- There was an issue with the catalog structure"
        Write-Warning "Run with -Verbose switch for more details"
    }

    # Extract VSIX files and copy to SSMS installation
    $extractedCount = 0
    if ($downloadCount -gt 0) {
        Write-Host "`nExtracting VSIX files for installation..." -ForegroundColor Yellow
        $extractedCount = Expand-AllVsixFiles -DownloadPath $DownloadPath -ExtractPath $ExtractPath
    
        # Copy extracted files to SSMS installation if extraction was successful
        if ($extractedCount -gt 0) {
            $ssmsPath = Get-SSMSInstallationPath
            if ($ssmsPath) {
                Copy-ExtractedFilesToSSMS -ExtractPath $ExtractPath -SSMSPath $ssmsPath
            }
            else {
                Write-Warning "Could not locate SSMS installation. Extracted files remain in: $ExtractPath"
            }    
        }
    }

    Write-Host "`n=========================================" -ForegroundColor Cyan
    Write-Host "Script completed." -ForegroundColor Cyan
    Sleep 5
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    Write-Host "Error details: $($_.Exception)" -ForegroundColor Red
}
finally {
    # Always restore execution policies, even if script fails
    Restore-ExecutionPolicies
}
