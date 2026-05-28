param(
    [Parameter(Mandatory = $true, HelpMessage = "SharePoint site URL, e.g. https://yourdomain.sharepoint.com/sites/YourSite")]
    [string]$SiteURL,

    [Parameter(Mandatory = $true, HelpMessage = "Number of latest versions to keep (minimum 1)")]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$VersionsToKeep,

    [Parameter(Mandatory = $false, HelpMessage = "Path relative to each document library root. It can point to a folder or file. Use '/' to process the whole library.")]
    [string]$RelativePath = "/"
)

# Connect to the SharePoint site
Connect-PnPOnline -Url $SiteURL -UseWebLogin

# Retrieve all document libraries in the site
$documentLibraries = Get-PnPList | Where-Object { $_.BaseTemplate -in @(101, 700) }
$normalizedRelativePath = $RelativePath.Replace("\", "/").Trim("/")

foreach ($lib in $documentLibraries) {
    Write-Host "Processing Document Library:" $lib.Title

    $rootFolder = Get-PnPProperty -ClientObject $lib -Property RootFolder
    $libraryRootUrl = $rootFolder.ServerRelativeUrl.TrimEnd("/")

    if ([string]::IsNullOrWhiteSpace($normalizedRelativePath)) {
        $targetPathUrl = $libraryRootUrl
    }
    else {
        $targetPathUrl = "$libraryRootUrl/$normalizedRelativePath"
    }

    Write-Host "Processing path recursively:" $targetPathUrl

    # Retrieve files under the target folder recursively, or the exact file when the path points to a file.
    $camlQuery = @"
<View Scope='RecursiveAll'>
    <Query>
        <Where>
            <And>
                <Eq>
                    <FieldRef Name='FSObjType' />
                    <Value Type='Integer'>0</Value>
                </Eq>
                <Or>
                    <Eq>
                        <FieldRef Name='FileRef' />
                        <Value Type='Text'>$targetPathUrl</Value>
                    </Eq>
                    <BeginsWith>
                        <FieldRef Name='FileDirRef' />
                        <Value Type='Text'>$targetPathUrl</Value>
                    </BeginsWith>
                </Or>
            </And>
        </Where>
    </Query>
</View>
"@

    $items = Get-PnPListItem -List $lib -PageSize 500 -Query $camlQuery
    
    foreach ($item in $items) {
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

        }
        catch {
            Write-Host "Error accessing versions for $($item["FileRef"]): $_"
        }
    }
}

Write-Host "Version cleanup complete."

# Disconnect the session

Disconnect-PnPOnline
