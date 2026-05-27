param(
    [Parameter(Mandatory = $true, HelpMessage = "SharePoint site URL, e.g. https://yourdomain.sharepoint.com/sites/YourSite")]
    [string]$SiteURL,

    [Parameter(Mandatory = $true, HelpMessage = "Number of latest versions to keep (minimum 1)")]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$VersionsToKeep
)

# Connect to the SharePoint site
Connect-PnPOnline -Url $SiteURL -UseWebLogin

# Retrieve all document libraries in the site
$documentLibraries = Get-PnPList | Where-Object { $_.BaseTemplate -in @(101, 700) }

foreach ($lib in $documentLibraries) {
    Write-Host "Processing Document Library:" $lib.Title
    
    # Retrieve all items from the document library
    $items = Get-PnPListItem -List $lib -PageSize 500
    
    foreach ($item in $items) {
        if ($item.FileSystemObjectType -eq "File") {
            try {
                $file = Get-PnPFile -Url $item["FileRef"] -AsListItem
                $versions = Get-PnPProperty -ClientObject $file -Property Versions
                
                # Calculate the number of versions to delete
                $versionsToDelete = $versions.Count - $VersionsToKeep

                # Skip files with only one version or when no versions need to be deleted
                if ($versions.Count -le 1 -or $versionsToDelete -le 0) {
                    Write-Host "Skipping file with only one version or no extra versions to delete:" $item["FileRef"]
                    continue
                }

                Write-Host "Processing file:" $item["FileRef"]
                
                # Delete older versions, preserving the specified number of latest versions
                for ($i = $versions.Count - 1; $i -ge $VersionsToKeep; $i--) {
                    Write-Host "Deleting older version $($versions[$i].VersionLabel) of $($item["FileRef"])"
                    $versions[$i].DeleteObject()
                    Invoke-PnPQuery
                }

            } catch {
                Write-Host "Error accessing versions for $($item["FileRef"]): $_"
            }
        }
    }
}

Write-Host "Version cleanup complete."

# Disconnect the session

Disconnect-PnPOnline
