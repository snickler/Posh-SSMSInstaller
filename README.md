# Posh-SSMSInstaller
Repo containing script to download and install SSMS 21+ for x64

## Usage

### Basic Usage
```powershell
# Install SSMS 21 (release channel) - default behavior  
.\Download-SSMS.ps1

# Install SSMS 22 (release channel)
.\Download-SSMS.ps1 -Version 22

# Install SSMS 21 (preview channel)
.\Download-SSMS.ps1 -Channel preview

# Install SSMS 22 (preview channel) 
.\Download-SSMS.ps1 -Version 22 -Channel preview
```

### Advanced Options
```powershell
# Custom download location
.\Download-SSMS.ps1 -DownloadPath "C:\Temp"

# Verbose output
.\Download-SSMS.ps1 -Verbose

# All options combined
.\Download-SSMS.ps1 -Version 22 -Channel preview -DownloadPath "C:\Temp" -Verbose
```

## Parameters

- **Version** (int): SSMS version to install. Supports `21` or `22`. Default: `21`
- **Channel** (string): Release channel to use. Supports `release` or `preview`. Default: `release`  
- **DownloadPath** (string): Custom download location. Default: `.\SSMS{Version}-Downloads`
- **ExtractPath** (string): Path for extracting VSIX files. Default: `.\Extracted`
- **Verbose** (switch): Enable verbose output

## Requirements

- Windows PowerShell 5.1 or PowerShell Core 6+
- Administrator privileges (script will auto-elevate)
- Internet connection
