# Posh-SSMSInstaller

PowerShell script repository for automated SQL Server Management Studio (SSMS) 21+ download and installation on Windows x64 systems.

Always reference these instructions first and fallback to search or bash commands only when you encounter unexpected information that does not match the info here.

## Working Effectively

### Prerequisites - Windows Environment REQUIRED
- **CRITICAL**: This repository contains Windows-only PowerShell scripts. Do NOT attempt to run the main script on Linux/Mac.
- Windows 10/11 or Windows Server with PowerShell 5.1 or PowerShell Core 7+
- Administrator privileges required for script execution
- Internet connection for downloading SSMS installer and components

### Repository Validation and Testing
- Validate PowerShell syntax: `pwsh -Command '$ErrorActionPreference = "Stop"; try { $null = [System.Management.Automation.PSParser]::Tokenize((Get-Content -Path "./Download-SSMS.ps1" -Raw), [ref]$null); "PowerShell syntax is valid" } catch { "Syntax error: $($_.Exception.Message)" }'`
- Check script parameters: `Get-Help -Name "./Download-SSMS.ps1"`
- Verify required cmdlets available: `Get-Command Invoke-RestMethod, Invoke-WebRequest, Start-Process, Get-Process`
- Check execution policy: `Get-ExecutionPolicy -List`
- Test on Windows VM or Windows environment only - script will NOT work on Linux/Mac

### Script Execution - NEVER CANCEL OPERATIONS
- **CRITICAL TIMING**: SSMS download and installation takes 15-45 minutes. NEVER CANCEL. Set timeout to 60+ minutes.
- **VSIX component processing**: Takes 5-10 minutes additional. NEVER CANCEL. Set timeout to 30+ minutes.
- Run as Administrator in elevated PowerShell session:
  ```powershell
  # Basic installation
  .\Download-SSMS.ps1
  
  # Custom download path with verbose output
  .\Download-SSMS.ps1 -DownloadPath "C:\Temp\SSMS" -Verbose
  
  # All parameters example
  .\Download-SSMS.ps1 -DownloadPath "C:\Custom\Downloads" -ExtractPath "C:\Custom\Extracted" -Verbose
  ```

### Script Behavior and Validation
- Script automatically handles execution policy changes (temporarily sets to Bypass, then restores)
- Downloads SSMS installer from Microsoft (vs_SSMS.exe) - takes 5-15 minutes
- Installs SSMS with `--arch x64 --quiet` parameters - takes 15-30 minutes  
- Downloads and extracts VSIX components from Microsoft channel manifest - takes 5-10 minutes
- Automatically detects SSMS installation path using VSSetup PowerShell module
- Installs VSSetup module if not present (Install-Module -Name VSSetup -Force -Scope CurrentUser -AllowClobber)

### Manual Validation Steps
After running the script, always validate:
- SSMS 21 is installed in Program Files: Check `${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio 21\Common7\IDE\Ssms.exe`
- VSIX components are present in SSMS installation directory
- Launch SSMS to verify functionality: `& "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio 21\Common7\IDE\Ssms.exe"`

## Development and Modification

### Script Structure
- Main script: `Download-SSMS.ps1` (719 lines)
- Core functions:
  - `Save-ExecutionPolicies` / `Restore-ExecutionPolicies`: Manages PowerShell execution policy
  - `Test-Administrator`: Checks admin privileges
  - `Stop-VisualStudioInstaller`: Closes VS installer processes
  - `Get-FileWithProgress`: Downloads files with error handling
  - `Expand-VsixFile` / `Expand-AllVsixFiles`: Extracts VSIX packages
  - `Get-SSMSInstallationPath`: Locates SSMS using VSSetup module + fallback methods
  - `Copy-ExtractedFilesToSSMS`: Copies components to SSMS installation

### Target Components Downloaded
The script downloads these specific VSIX components:
- Microsoft.VisualStudio.MinShell.Targeted
- Microsoft.VisualStudio.ExtensionManager.x64
- Microsoft.DiagnosticsHub.Runtime.Targeted
- Microsoft.ServiceHub.Managed
- Microsoft.ServiceHub.amd64
- Microsoft.VisualStudio.Identity
- Microsoft.VisualStudio.ExtensionManager

### Making Changes
- Always validate PowerShell syntax before committing changes
- Test on Windows VM with fresh SSMS installation
- Verify all error handling paths work correctly
- Test with different Windows versions (Windows 10/11, Server)
- Validate both success and failure scenarios
- **IMPORTANT**: Do NOT commit temporary download folders (SSMS21-Downloads, Extracted) - add to .gitignore if needed

### Common Development Tasks
- **Adding new VSIX components**: Modify `$targetItemIds` array around line 532
- **Changing SSMS version**: Update `$ssmsInstallerUrl` around line 552
- **Modifying installation detection**: Edit `Get-SSMSInstallationPath` and `Get-SSMSInstallationPathFallback` functions
- **Updating download paths**: Check registry paths in `Get-SSMSInstallationPathFallback` function

## Repository Information

### File Structure
```
/
├── .github/
│   └── copilot-instructions.md     # This file
├── .git/                          # Git repository data
├── .gitignore                     # Standard Visual Studio .gitignore
├── Download-SSMS.ps1              # Main PowerShell script (29KB)
├── LICENSE                        # MIT License
└── README.md                      # Basic repository description
```

### Dependencies
- Windows PowerShell 5.1+ or PowerShell Core 7+
- VSSetup PowerShell module (auto-installed by script)
- Administrator privileges
- Internet access to aka.ms Microsoft domains

### URLs Used by Script
- SSMS installer: `https://aka.ms/ssms/21/release/vs_SSMS.exe`
- Channel manifest: `https://aka.ms/ssms/21/release/channel`
- VSIX payloads: Retrieved dynamically from Microsoft catalog

## Troubleshooting

### Common Issues
- **"Execution policy restricted"**: Script handles this automatically, but manual override: `Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force`
- **"Access denied"**: Must run as Administrator
- **"SSMS installation not found"**: Script has fallback detection methods, manually check common paths
- **Download failures**: Check internet connectivity and Microsoft service availability
- **VSIX extraction fails**: Verify disk space and file permissions

### Debug Mode
Always use `-Verbose` parameter for detailed troubleshooting output:
```powershell
.\Download-SSMS.ps1 -Verbose
```

### Expected Execution Time
- Total script execution: 25-55 minutes (NEVER CANCEL)
- SSMS download: 5-15 minutes  
- SSMS installation: 15-30 minutes
- VSIX processing: 5-10 minutes
- Always set timeouts to 60+ minutes for full execution

## Security Notes
- Script temporarily modifies execution policy (restored on completion)
- Requires Administrator privileges for SSMS installation
- Downloads from trusted Microsoft domains only
- Uses HTTPS for all downloads
- No sensitive data stored or transmitted beyond Microsoft's standard telemetry