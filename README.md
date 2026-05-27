# SharePointVersionCleaner

## Description
The SharePointVersionCleaner script is designed to help manage document versions within SharePoint document libraries. It connects to a specified SharePoint site, iterates through all document libraries, and prunes older versions of files, retaining only a user-defined number of the most recent versions. This script is ideal for SharePoint site administrators looking to optimize storage and maintain version control.

## Requirements
- PowerShell 5.1 or higher
- SharePointPnPPowerShellOnline module
- Access to a SharePoint Online site

## Installation
Before running the script, ensure that the SharePointPnPPowerShellOnline module is installed on your system. If it is not installed, run the following PowerShell command to install it:

```powershell
Install-Module SharePointPnPPowerShellOnline -AllowClobber
```

## Usage
Open PowerShell and navigate to the directory containing the SharePointVersionCleaner.ps1 script.
Run the script using the following command: .\SharePointVersionCleaner.ps1

When prompted, enter the number of latest versions you wish to keep for each document.
Note: The script requires you to log in to your SharePoint site. A web login prompt will appear for authentication.

## Configuration
To target a specific SharePoint site, modify the $SiteURL variable at the beginning of the script with your SharePoint site URL:

```powershell
$SiteURL = "https://yourdomain.sharepoint.com/sites/YourSite"
```

## Troubleshooting

### OneDrive for Business: Document Library Not Found

Standard SharePoint document libraries use `BaseTemplate = 101`. However, **OneDrive for Business personal sites** store files in a `文档`（Documents）library with `BaseTemplate = 700`, which is a special template number specific to OneDrive.

If the script runs without processing any files, first check which libraries exist on the site and their template numbers:

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

If you encounter a similar issue on other SharePoint variants, use the diagnostic command above to identify the correct `BaseTemplate` number, and add it to the filter accordingly.