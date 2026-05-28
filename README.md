# SharePointVersionCleaner

## Description
SharePointVersionCleaner helps manage document versions within SharePoint document libraries. It connects to a specified SharePoint site, iterates through all document libraries, and prunes older versions of files, retaining only a user-defined number of the most recent versions. This is ideal for cleaning up storage without touching the current version of any file.

Two scripts are provided with identical parameters and behavior. Choose the one that works for your environment:

| Script | Method | Requires |
|---|---|---|
| `SharePointVersionCleaner_PnP.ps1` | PnP PowerShell module | `SharePointPnPPowerShellOnline` installed; web login |
| `SharePointVersionCleaner_Graph.ps1` | Microsoft Graph API | No modules; device code login via browser |

**Recommended:** Use the Graph version if possible — it has no external dependencies and is more likely to remain stable long-term. Use the PnP version if the Graph version cannot authenticate in your tenant.

## Requirements

### SharePointVersionCleaner_PnP.ps1
- PowerShell 5.1 or higher
- `SharePointPnPPowerShellOnline` module (archived, but still functional)
- Access to a SharePoint Online site

### SharePointVersionCleaner_Graph.ps1
- PowerShell 5.1 or higher
- No modules required
- Access to a SharePoint Online site

## Installation

For the PnP version, install the module if not already present:

```powershell
Install-Module SharePointPnPPowerShellOnline -AllowClobber
```

The Graph version requires no installation.

## Usage

Both scripts share the same parameters:

```powershell
.\SharePointVersionCleaner_Graph.ps1 `
  -SiteURL "https://yourdomain.sharepoint.com/sites/YourSite" `
  -VersionsToKeep 5
```

By default, all document libraries are processed recursively from their root.

To process only a specific folder and everything under it, pass `-RelativePath`. The path is relative to each document library root, not the site URL:

```powershell
.\SharePointVersionCleaner_Graph.ps1 `
  -SiteURL "https://yourdomain.sharepoint.com/sites/YourSite" `
  -VersionsToKeep 5 `
  -RelativePath "Reports/2024"
```

To target a single file:

```powershell
.\SharePointVersionCleaner_Graph.ps1 `
  -SiteURL "https://yourdomain.sharepoint.com/sites/YourSite" `
  -VersionsToKeep 5 `
  -RelativePath "report.docx"
```

Use `/` to explicitly process the whole library:

```powershell
.\SharePointVersionCleaner_Graph.ps1 `
  -SiteURL "https://yourdomain.sharepoint.com/sites/YourSite" `
  -VersionsToKeep 5 `
  -RelativePath "/"
```

The PnP version uses the same syntax — just replace the script name.

### Authentication

The PnP version opens a web login window automatically.

The Graph version uses device code flow: after launching the script, open a browser, go to `https://login.microsoft.com/device`, and enter the code shown in the terminal. The script resumes automatically once authenticated.

## Troubleshooting

### Graph version fails to authenticate

Some tenants restrict which applications can request tokens, including Microsoft first-party apps. If the Graph version shows `AADSTS65002` during login, your tenant administrator has not enabled this flow. In that case, use the PnP version instead.

### PnP version: OneDrive document library not found

Standard SharePoint document libraries use `BaseTemplate = 101`. OneDrive for Business personal sites store files in a library with `BaseTemplate = 700`. If the PnP script runs without processing any files, check which libraries exist and their template numbers:

```powershell
Connect-PnPOnline -Url "https://yourdomain-my.sharepoint.com/personal/username" -UseWebLogin
Get-PnPList | Select-Object Title, BaseTemplate | Format-Table -AutoSize
```

A OneDrive site typically shows output like this:

```
Title              BaseTemplate
-----              ------------
文档 / Documents        700
样式库                   101
PersonalCacheLibrary     101
表单模板                 101
...
```

The script already handles this by filtering for both template types:

```powershell
$documentLibraries = Get-PnPList | Where-Object { $_.BaseTemplate -in @(101, 700) }
```

If you encounter a similar issue on other SharePoint variants, use the diagnostic command above to identify the correct `BaseTemplate` number and add it to the filter accordingly.