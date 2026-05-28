param(
    [Parameter(Mandatory = $true, HelpMessage = "SharePoint site URL, e.g. https://yourdomain.sharepoint.com/sites/YourSite")]
    [string]$SiteURL,

    [Parameter(Mandatory = $true, HelpMessage = "Number of latest versions to keep (minimum 1)")]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$VersionsToKeep,

    [Parameter(Mandatory = $false, HelpMessage = "Path relative to each document library root. It can point to a folder or file. Use '/' to process the whole library.")]
    [string]$RelativePath = "/"
)

# -----------------------------------------------------------------------
# Authentication — device code flow, no module or app registration needed
# -----------------------------------------------------------------------
$clientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c"  # Microsoft Office first-party app
$scopes = "https://graph.microsoft.com/.default"

$deviceCodeResponse = Invoke-RestMethod -Method POST `
    -Uri "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode" `
    -Body @{ client_id = $clientId; scope = $scopes }

Write-Host $deviceCodeResponse.message

$interval = $deviceCodeResponse.interval
$expiry = (Get-Date).AddSeconds($deviceCodeResponse.expires_in)
$token = $null

while ((Get-Date) -lt $expiry) {
    Start-Sleep -Seconds $interval
    try {
        $token = Invoke-RestMethod -Method POST `
            -Uri "https://login.microsoftonline.com/common/oauth2/v2.0/token" `
            -Body @{
            client_id   = $clientId
            grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
            device_code = $deviceCodeResponse.device_code
        }
        Write-Host "Authenticated successfully."
        break
    }
    catch {
        $err = ($_.ErrorDetails.Message | ConvertFrom-Json).error
        if ($err -eq "authorization_pending") { continue }
        Write-Error "Authentication failed: $err"
        exit 1
    }
}

if (-not $token) {
    Write-Error "Authentication timed out."
    exit 1
}

$headers = @{ Authorization = "Bearer $($token.access_token)" }

# -----------------------------------------------------------------------
# Helper — Graph paged requests
# -----------------------------------------------------------------------
function Invoke-GraphGetAll {
    param([string]$Uri)
    $results = @()
    $next = $Uri
    while ($next) {
        $page = Invoke-RestMethod -Uri $next -Headers $headers
        if ($page.value) { $results += $page.value }
        $next = $page.'@odata.nextLink'
    }
    return $results
}

# -----------------------------------------------------------------------
# Resolve site ID from URL
# -----------------------------------------------------------------------
$uri = [System.Uri]$SiteURL
$hostname = $uri.Host                                 # e.g. contoso.sharepoint.com
$sitePath = $uri.AbsolutePath.TrimStart("/")          # e.g. sites/YourSite  or  personal/user

$siteInfo = Invoke-RestMethod `
    -Uri "https://graph.microsoft.com/v1.0/sites/${hostname}:/${sitePath}" `
    -Headers $headers
$siteId = $siteInfo.id
Write-Host "Resolved site: $($siteInfo.displayName) ($siteId)"

# -----------------------------------------------------------------------
# Enumerate drives (document libraries)
# -----------------------------------------------------------------------
$drives = Invoke-GraphGetAll "https://graph.microsoft.com/v1.0/sites/$siteId/drives"

$normalizedRelativePath = $RelativePath.Replace("\", "/").Trim("/")

# -----------------------------------------------------------------------
# Process versions for a single file item
# -----------------------------------------------------------------------
function Remove-OldVersions {
    param([string]$driveId, [string]$itemId, [string]$itemPath)

    $versionsUri = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$itemId/versions"
    $versions = Invoke-GraphGetAll $versionsUri

    if ($versions.Count -le $VersionsToKeep) {
        Write-Host "  Skipping (only $($versions.Count) versions): $itemPath"
        return
    }

    # Graph returns versions newest-first; delete the oldest ones first for safety
    $toDelete = ($versions | Select-Object -Last ($versions.Count - $VersionsToKeep)) |
    Sort-Object -Property lastModifiedDateTime

    Write-Host "  Processing: $itemPath — $($versions.Count) versions, keeping $VersionsToKeep"

    foreach ($v in $toDelete) {
        # The current (latest) version cannot be deleted; Graph returns 405 for it — skip
        if ($v.lastModifiedDateTime -eq $versions[0].lastModifiedDateTime -and
            $v.id -eq $versions[0].id) { continue }

        Write-Host "    Deleting version $($v.id)"
        try {
            Invoke-RestMethod -Method DELETE `
                -Uri "https://graph.microsoft.com/v1.0/drives/$driveId/items/$itemId/versions/$($v.id)" `
                -Headers $headers | Out-Null
        }
        catch {
            Write-Host "    Error deleting version $($v.id): $_"
        }
    }
}

# -----------------------------------------------------------------------
# Recursively list all files under a folder path
# -----------------------------------------------------------------------
function Get-AllFiles {
    param([string]$driveId, [string]$folderId)

    $children = Invoke-GraphGetAll `
        "https://graph.microsoft.com/v1.0/drives/$driveId/items/$folderId/children"

    foreach ($child in $children) {
        if ($child.folder) {
            Get-AllFiles -driveId $driveId -folderId $child.id
        }
        elseif ($child.file) {
            Remove-OldVersions -driveId $driveId -itemId $child.id -itemPath $child.name
        }
    }
}

# -----------------------------------------------------------------------
# Main loop over drives
# -----------------------------------------------------------------------
foreach ($drive in $drives) {
    Write-Host "Processing drive: $($drive.name)"

    if ([string]::IsNullOrWhiteSpace($normalizedRelativePath)) {
        # Process entire drive from root
        $rootItem = Invoke-RestMethod `
            -Uri "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root" `
            -Headers $headers
        Get-AllFiles -driveId $drive.id -folderId $rootItem.id
    }
    else {
        # Resolve the target path — could be a file or a folder
        try {
            $targetItem = Invoke-RestMethod `
                -Uri "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root:/$normalizedRelativePath" `
                -Headers $headers
        }
        catch {
            # Path doesn't exist in this drive — skip silently
            continue
        }

        if ($targetItem.folder) {
            Get-AllFiles -driveId $drive.id -folderId $targetItem.id
        }
        elseif ($targetItem.file) {
            Remove-OldVersions -driveId $drive.id -itemId $targetItem.id -itemPath $normalizedRelativePath
        }
    }
}

Write-Host "Version cleanup complete."