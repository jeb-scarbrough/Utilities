<# ========================
   CONFIGURATION
   ======================== #>

# Root folder where the archive will be stored
$ExportRoot = "C:\TeamsArchive"

# Your tenant short name (for SharePoint URLs)
$TenantName = "salesvista"

# Your PnP Entra app registration IDs
# From Entra ID → App registrations → "PnP Archive App" → Overview
$ClientId = "e05585c0-d531-44aa-a494-b93113a189ce"   # Application (client) ID
$TenantId = "82cc1d78-5e41-4edf-87ec-3fd8759fdec4"   # Directory (tenant) ID

<# ========================
   SETUP
   ======================== #>

if (-not (Test-Path $ExportRoot)) {
    New-Item -ItemType Directory -Path $ExportRoot -Force | Out-Null
}

Import-Module MicrosoftTeams
Import-Module PnP.PowerShell

Write-Host "Connecting to Microsoft Teams..."
Connect-MicrosoftTeams   # interactive sign-in

Write-Host "Getting all Teams..."
$teams = Get-Team

<# ========================
   EXPORT CHANNEL LISTS
   ======================== #>

foreach ($team in $teams) {
    $safeTeamName = $team.DisplayName -replace '[\\/:*?"<>|]', '_'
    $teamFolder   = Join-Path $ExportRoot ("Team_" + $safeTeamName)

    if (-not (Test-Path $teamFolder)) {
        New-Item -ItemType Directory -Path $teamFolder -Force | Out-Null
    }

    Write-Host "Exporting channels for Team: $($team.DisplayName)"
    $channels = Get-TeamChannel -GroupId $team.GroupId -ErrorAction SilentlyContinue

    if ($channels) {
        $channels | Export-Csv -Path (Join-Path $teamFolder "Channels.csv") -NoTypeInformation -Encoding UTF8
    }
}

<# ========================
   DOWNLOAD FILES & FOLDERS (RECURSIVE) FOR ALL TEAMS
   ======================== #>

$adminUrl = "https://$TenantName-admin.sharepoint.com"
Write-Host "Connecting to SharePoint Admin: $adminUrl ..."
Connect-PnPOnline -Url $adminUrl -ClientId $ClientId -Tenant $TenantId -Interactive

Write-Host "Retrieving group-connected SharePoint sites (Teams)..."
$groupSites = Get-PnPTenantSite -Template GROUP#0 -IncludeOneDriveSites:$false

foreach ($team in $teams) {
    $safeTeamName = $team.DisplayName -replace '[\\/:*?"<>|]', '_'
    $teamFolder   = Join-Path $ExportRoot ("Team_" + $safeTeamName)
    $filesFolder  = Join-Path $teamFolder "Files"

    if (-not (Test-Path $teamFolder)) {
        New-Item -ItemType Directory -Path $teamFolder -Force | Out-Null
    }
    if (-not (Test-Path $filesFolder)) {
        New-Item -ItemType Directory -Path $filesFolder -Force | Out-Null
    }

    # Find the SharePoint site that backs this Team, via GroupId
    $site = $groupSites | Where-Object { $_.GroupId -eq $team.GroupId }

    if (-not $site) {
        Write-Warning "No SharePoint site found for Team '$($team.DisplayName)' (GroupId: $($team.GroupId)). Skipping files."
        continue
    }

    Write-Host ""
    Write-Host "==== Team: $($team.DisplayName) ===="
    Write-Host "Site: $($site.Url)"

    try {
        # Connect to the team site
        Connect-PnPOnline -Url $site.Url -ClientId $ClientId -Tenant $TenantId -Interactive

        # Try to get the main document library (Shared Documents or first visible document library)
        $docLib = Get-PnPList -Identity "Shared Documents" -ErrorAction SilentlyContinue

        if (-not $docLib) {
            # Fallback: first non-hidden document library (BaseTemplate 101)
            $docLib = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden } | Select-Object -First 1
        }

        if (-not $docLib) {
            Write-Warning "No document library found for site $($site.Url). Skipping."
            continue
        }

        Write-Host "Using document library: $($docLib.Title)"

        # Server-relative root of the library, used for relative path calc
        $libRootUrl = $docLib.RootFolder.ServerRelativeUrl.TrimEnd('/')

        # Get ALL list items (folders + files) from the library
        Write-Host "Retrieving list items (this may take a while for large libraries)..."
        $items = Get-PnPListItem -List $docLib -PageSize 2000 -Fields "FileRef","FileDirRef","FSObjType","FileLeafRef"

        if (-not $items) {
            Write-Host "No items found in '$($docLib.Title)' for Team '$($team.DisplayName)'."
            continue
        }

        # --------- FOLDERS (recursive view) ---------
        $folders = $items | Where-Object { $_["FSObjType"] -eq 1 }

        $folderInfos = foreach ($folder in $folders) {
            $fileRef = [string]$folder["FileRef"]
            if (-not $fileRef.StartsWith($libRootUrl, [System.StringComparison]::OrdinalIgnoreCase)) {
                continue
            }

            $relativePath = $fileRef.Substring($libRootUrl.Length).TrimStart('/')
            if ([string]::IsNullOrWhiteSpace($relativePath)) {
                $relativePath = "/"   # root of library
            }

            [PSCustomObject]@{
                FolderRelativePath = $relativePath
            }
        }

        if ($folderInfos) {
            $folderInfos |
                Sort-Object FolderRelativePath |
                Export-Csv -Path (Join-Path $teamFolder "Folders.csv") -NoTypeInformation -Encoding UTF8
            Write-Host ("Folders discovered: {0}" -f $folderInfos.Count)
        }
        else {
            Write-Host "No folders found (library may contain only files at root)."
        }

        # --------- FILES (download everything) ---------
        $files = $items | Where-Object { $_["FSObjType"] -eq 0 }
        Write-Host ("Files discovered: {0}" -f $files.Count)

        foreach ($file in $files) {
            $fileRef   = [string]$file["FileRef"]      # e.g. /sites/Corporate/Shared Documents/01 - Administration/...
            $fileName  = [string]$file["FileLeafRef"]

            if (-not $fileRef.StartsWith($libRootUrl, [System.StringComparison]::OrdinalIgnoreCase)) {
                continue
            }

            # Path relative to the library root
            $relativePath = $fileRef.Substring($libRootUrl.Length).TrimStart('/')

            # Preserve folder hierarchy under Files\
            $targetPath = Join-Path $filesFolder $relativePath
            $targetDir  = Split-Path $targetPath -Parent

            if (-not (Test-Path $targetDir)) {
                New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
            }

            Get-PnPFile -Url $fileRef -Path $targetDir -FileName $fileName -AsFile -Force | Out-Null
        }

        Write-Host "Downloaded all files for Team: $($team.DisplayName)"
    }
    catch {
        Write-Warning "Error processing Team '$($team.DisplayName)': $_"
    }
}

Write-Host ""
Write-Host "Archive complete. Root folder: $ExportRoot"
