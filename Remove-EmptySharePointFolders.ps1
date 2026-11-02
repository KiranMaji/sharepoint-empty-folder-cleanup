<#
.SYNOPSIS
    Validates and deletes approved empty-folder candidates from a
    specified SharePoint document library directory.

.DESCRIPTION
    Reads a previously generated and reviewed empty-folder CSV report.
    Connects to SharePoint Online through Microsoft Graph PowerShell
    and validates each candidate against its current item metadata,
    folder status, parent hierarchy, and child-item count.

    Supports dry-run mode and requires a single explicit confirmation
    before executing deletion operations. Eligible empty folders are
    processed automatically without requiring individual folder-name
    confirmations.

    Successful deletion operations move folders to the SharePoint
    recycle bin, subject to the applicable SharePoint and Microsoft
    Graph behavior and permissions. Each candidate's outcome is
    recorded in a timestamped CSV audit log.

    The script processes only candidates listed in the configured CSV
    and restricts processing to the configured target directory.
    Candidates that fail validation or contain items are skipped.

.PARAMETER ExecuteDeletion
    Enables deletion mode. Without this switch, the script performs
    a dry run and makes no changes to SharePoint content.

.OUTPUTS
    Timestamped CSV audit log containing the processing status and
    outcome for each candidate.

.NOTES
    Requires PowerShell 7 and the Microsoft Graph PowerShell modules.
    Requires appropriate Microsoft Graph delegated write permissions.

    Run the scan script first, review the candidate and error reports,
    and configure the deletion script to use the approved candidate CSV.
    Always perform a dry run before enabling deletion.
#>


param(
    [switch]$ExecuteDeletion
)

$ErrorActionPreference = "Stop"

# ============================================================
# CONFIGURATION
# ============================================================

$SiteUrl      = "https://example.com/sites/example"
$LibraryName  = "Documents"
$TargetFolder = "Directory/Folder Name"

# Use the latest reviewed, successful read-only audit CSV.
$CandidateCsv = "C:\Scripting\EmptyFolders_20261015_100040.csv"

$OutputFolder = "C:\Scripting\"
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogCsv = Join-Path $OutputFolder "DeleteEmptyFolders_$Timestamp.csv"

# ============================================================
# PREPARE
# ============================================================

New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null

foreach ($Module in @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Sites",
    "Microsoft.Graph.Files"
)) {
    if (-not (Get-Module -ListAvailable -Name $Module)) {
        throw "Required module missing: $Module"
    }

    Import-Module $Module -ErrorAction Stop
}

if (-not (Test-Path -LiteralPath $CandidateCsv)) {
    throw "Candidate CSV not found: $CandidateCsv"
}

$Candidates = @(Import-Csv -LiteralPath $CandidateCsv)

if ($Candidates.Count -eq 0) {
    throw "Candidate CSV is empty. Nothing to process."
}

# Validate every candidate before connecting or deleting.
$ExpectedPrefix = "$LibraryName/$TargetFolder/"

foreach ($Candidate in $Candidates) {
    if ([string]::IsNullOrWhiteSpace($Candidate.FolderId) -or
        [string]::IsNullOrWhiteSpace($Candidate.FolderPath)) {
        throw "Candidate CSV has a missing FolderId or FolderPath."
    }

    if (-not $Candidate.FolderPath.StartsWith(
        $ExpectedPrefix,
        [StringComparison]::OrdinalIgnoreCase
    )) {
        throw "Candidate is outside the permitted subtree: $($Candidate.FolderPath)"
    }
}

Write-Host "`nCandidates loaded: $($Candidates.Count)" -ForegroundColor Cyan

$Candidates |
    Select-Object FolderName, FolderPath, FolderId |
    Format-Table -AutoSize

# ============================================================
# AUTHENTICATION
# ============================================================

$Context = Get-MgContext

if (-not $Context -or
    $Context.Scopes -notcontains "Sites.ReadWrite.All") {

    if ($Context) {
        Disconnect-MgGraph | Out-Null
    }

    Connect-MgGraph -Scopes "Sites.ReadWrite.All" -NoWelcome
}

$Context = Get-MgContext

if (-not $Context -or
    $Context.Scopes -notcontains "Sites.ReadWrite.All") {
    throw "Sites.ReadWrite.All was not granted. Stopping safely."
}

# ============================================================
# RESOLVE SITE AND DOCUMENTS DRIVE
# ============================================================

$Uri = [Uri]$SiteUrl
$SitePath = $Uri.AbsolutePath.TrimEnd("/")

$Site = Get-MgSite -SiteId "$($Uri.Host):$SitePath" -ErrorAction Stop

if (-not $Site -or [string]::IsNullOrWhiteSpace($Site.Id)) {
    throw "Site ID could not be resolved."
}

$Drives = @(Get-MgSiteDrive -SiteId $Site.Id -All -ErrorAction Stop)

$Drive = $Drives |
    Where-Object { $_.Name.Trim() -ieq $LibraryName } |
    Select-Object -First 1

if (-not $Drive -or [string]::IsNullOrWhiteSpace($Drive.Id)) {
    throw "Documents drive ID could not be resolved. No deletion performed."
}

Write-Host "`nSite: $($Site.WebUrl)" -ForegroundColor Cyan
Write-Host "Library: $($Drive.Name)"
Write-Host "Drive ID resolved successfully." -ForegroundColor Green

# ============================================================
# GRAPH HELPER FUNCTIONS
# ============================================================

function Get-Field {
    param(
        [Parameter(Mandatory)][object]$Object,
        [Parameter(Mandatory)][string]$Name
    )

    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) {
            return $Object[$Name]
        }
        return $null
    }

    $Property = $Object.PSObject.Properties[$Name]

    if ($null -ne $Property) {
        return $Property.Value
    }

    return $null
}

function Get-ItemById {
    param([Parameter(Mandatory)][string]$ItemId)

    if ([string]::IsNullOrWhiteSpace($Drive.Id)) {
        throw "Drive ID is empty. Refusing request."
    }

    $D = [Uri]::EscapeDataString($Drive.Id)
    $I = [Uri]::EscapeDataString($ItemId)

    $RequestUri = "https://graph.microsoft.com/v1.0/drives/$D/items/$I"

    return Invoke-MgGraphRequest `
        -Method GET `
        -Uri $RequestUri `
        -ErrorAction Stop
}

function Get-Children {
    param([Parameter(Mandatory)][string]$ItemId)

    if ([string]::IsNullOrWhiteSpace($Drive.Id)) {
        throw "Drive ID is empty. Refusing request."
    }

    $D = [Uri]::EscapeDataString($Drive.Id)
    $I = [Uri]::EscapeDataString($ItemId)

    $RequestUri = "https://graph.microsoft.com/v1.0/drives/$D/items/$I/children?`$select=id,name,folder,file,parentReference&`$top=200"

    $Items = [System.Collections.Generic.List[object]]::new()

    while ($RequestUri) {
        $Response = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $RequestUri `
            -ErrorAction Stop

        foreach ($Child in @(Get-Field $Response "value")) {
            if ($null -ne $Child) {
                $Items.Add($Child)
            }
        }

        $NextLink = Get-Field $Response "@odata.nextLink"

        if ($NextLink) {
            $RequestUri = [string]$NextLink
        }
        else {
            $RequestUri = $null
        }
    }

    return $Items.ToArray()
}

function Test-IsFolder {
    param([Parameter(Mandatory)][object]$Item)

    return ($null -ne (Get-Field $Item "folder"))
}

function Test-WithinTarget {
    param(
        [Parameter(Mandatory)][string]$ItemId,
        [Parameter(Mandatory)][string]$TargetId
    )

    $Seen = [System.Collections.Generic.HashSet[string]]::new()
    $Current = Get-ItemById -ItemId $ItemId

    for ($Depth = 0; $Depth -lt 50; $Depth++) {
        $CurrentId = [string](Get-Field $Current "id")

        # Never delete the target root itself.
        if ($CurrentId -eq $TargetId) {
            return $false
        }

        if (-not $Seen.Add($CurrentId)) {
            throw "Cycle detected in parent chain."
        }

        $ParentRef = Get-Field $Current "parentReference"
        $ParentId = [string](Get-Field $ParentRef "id")

        if ([string]::IsNullOrWhiteSpace($ParentId)) {
            return $false
        }

        if ($ParentId -eq $TargetId) {
            return $true
        }

        $Current = Get-ItemById -ItemId $ParentId
    }

    throw "Could not verify folder ancestry within 50 levels."
}

# ============================================================
# RESOLVE TARGET ROOT
# ============================================================

$Root = Get-MgDriveRoot -DriveId $Drive.Id -ErrorAction Stop
$RootItems = @(Get-Children -ItemId $Root.Id)

$Target = $RootItems |
    Where-Object {
        (Get-Field $_ "name") -ieq $TargetFolder -and
        (Test-IsFolder $_)
    } |
    Select-Object -First 1

if (-not $Target) {
    throw "Target folder '$TargetFolder' was not found at the library root."
}

$TargetId = [string](Get-Field $Target "id")

Write-Host "Permitted subtree: $LibraryName/$TargetFolder" -ForegroundColor Green

# ============================================================
# SINGLE CONFIRMATION
# ============================================================

if (-not $ExecuteDeletion) {
    Write-Host "`nDRY RUN: no content will be deleted." -ForegroundColor Yellow
}
else {
    Write-Host "`nWARNING: DELETION MODE ENABLED" -ForegroundColor Red
    Write-Host "Candidates in approved CSV: $($Candidates.Count)"
    Write-Host "Scope: $LibraryName/$TargetFolder"
    Write-Host "Only candidates still empty and within this subtree will be deleted."
    Write-Host "Deleted folders are moved to the SharePoint recycle bin."

    $Confirmation = Read-Host "`nType DELETE to process the entire candidate list"

    if ($Confirmation -cne "DELETE") {
        throw "Confirmation failed. No deletion performed."
    }
}

# ============================================================
# PROCESS CANDIDATES
# ============================================================

$Results = [System.Collections.Generic.List[object]]::new()

$Index = 0

foreach ($Candidate in $Candidates) {
    $Index++

    $Status = ""
    $Message = ""

    Write-Progress `
        -Activity "Processing empty-folder candidates" `
        -Status "$Index of $($Candidates.Count): $($Candidate.FolderPath)" `
        -PercentComplete (($Index / $Candidates.Count) * 100)

    try {
        # Re-fetch current item metadata.
        $Item = Get-ItemById -ItemId ([string]$Candidate.FolderId)

        $ItemId = [string](Get-Field $Item "id")

        if ($ItemId -ne [string]$Candidate.FolderId) {
            throw "Item ID mismatch."
        }

        if (-not (Test-IsFolder $Item)) {
            throw "Item is no longer a folder."
        }

        # Verify actual SharePoint ancestry.
        if (-not (Test-WithinTarget `
            -ItemId $ItemId `
            -TargetId $TargetId)) {
            throw "Item is outside the permitted target subtree."
        }

        # Recheck the folder immediately before processing.
        $CurrentChildren = @(Get-Children -ItemId $ItemId)

        if ($CurrentChildren.Count -ne 0) {
            throw "Folder is no longer empty. Skipping."
        }

        if (-not $ExecuteDeletion) {
            $Status = "DRY RUN - verified empty"
            $Message = "No changes made."

            Write-Host "[DRY RUN] $($Candidate.FolderPath)" `
                -ForegroundColor Yellow
        }
        else {
            $D = [Uri]::EscapeDataString($Drive.Id)
            $I = [Uri]::EscapeDataString($ItemId)

            $DeleteUri = "https://graph.microsoft.com/v1.0/drives/$D/items/$I"

            $ETag = Get-Field $Item "eTag"
            $Headers = @{}

            if ($ETag) {
                $Headers["If-Match"] = [string]$ETag
            }

            Invoke-MgGraphRequest `
                -Method DELETE `
                -Uri $DeleteUri `
                -Headers $Headers `
                -ErrorAction Stop | Out-Null

            $Status = "Deleted to recycle bin"
            $Message = "Graph DELETE request succeeded."

            Write-Host "[DELETED] $($Candidate.FolderPath)" `
                -ForegroundColor Green
        }
    }
    catch {
        $Status = "ERROR - not confirmed deleted"
        $Message = $_.Exception.Message

        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            $Message += " | " + $_.ErrorDetails.Message
        }

        Write-Warning "$($Candidate.FolderPath): $Message"
    }

    $Results.Add([PSCustomObject]@{
        FolderName = $Candidate.FolderName
        FolderPath = $Candidate.FolderPath
        FolderId   = $Candidate.FolderId
        Status     = $Status
        Message    = $Message
    })
}

Write-Progress -Activity "Processing empty-folder candidates" -Completed

# ============================================================
# EXPORT AUDIT LOG
# ============================================================

$Results |
    Export-Csv -LiteralPath $LogCsv -NoTypeInformation -Encoding utf8

Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "               FINAL SUMMARY                " -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

$Results | Format-Table FolderPath, Status -AutoSize

Write-Host "`nAudit log: $LogCsv" -ForegroundColor Green

if (-not $ExecuteDeletion) {
    Write-Host "`nDRY RUN COMPLETE. Nothing was deleted." `
        -ForegroundColor Yellow
}
else {
    $DeletedCount = @(
        $Results | Where-Object { $_.Status -eq "Deleted to recycle bin" }
    ).Count

    $ErrorCount = @(
        $Results | Where-Object { $_.Status -like "ERROR*" }
    ).Count

    Write-Host "`nSuccessfully deleted: $DeletedCount"
    Write-Host "Errors: $ErrorCount"

    Write-Host "Review the audit CSV for the outcome of every candidate." `
        -ForegroundColor Cyan
}
