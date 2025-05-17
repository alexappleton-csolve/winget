# Winget PowerShell Module

This repository contains a PowerShell module that wraps the Windows Package Manager (`winget`) command line tool. The module provides helper functions for installing, upgrading, and uninstalling applications and captures detailed log output during operations.

## Installation

1. Copy the `winget.psm1` file to a folder included in your PowerShell module path or import it directly using `Import-Module` with the full path.
2. Ensure you have the Windows Package Manager installed. You can use the `Enable-WG` function to download and install it if necessary.

```powershell
# Import the module (adjust the path as needed)
Import-Module .\winget.psm1
```

## Usage

The module exposes several functions to help automate software management:

- `Enable-WG [-Preview]` – Install the latest stable or preview version of Winget.
- `Get-WGList [-appid <ID>]` – List installed applications. Optionally filter by application ID.
- `Get-WGUpgrade` – Display applications with available updates.
- `Start-WGUpgrade [-appid <ID>] [-All]` – Upgrade a single application or all outdated applications.
- `Start-WGInstall -appid <ID>` – Install an application by ID.
- `Start-WGUninstall -appid <ID>` – Uninstall an application by ID.
- `Get-WGver` – Show the installed winget version.

All activity is logged to `C:\Windows\Temp\ps_winget.log`. The log is rotated whenever the module loads. The `Process-WingetResults` function normalizes winget output, removing blank lines, collapsing whitespace, and stripping non‑basic Latin characters so the log remains concise.

## Example

```powershell
# Update all applications that have updates available
Start-WGUpgrade -All

# Install a specific package
Start-WGInstall -appid Microsoft.VisualStudioCode
```

## Notes

The module attempts to keep interactions with winget unattended by accepting package and source agreements on your behalf. Review the code and ensure this behavior meets your organization’s requirements before running in production.

