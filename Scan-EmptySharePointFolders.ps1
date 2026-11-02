<#
.SYNOPSIS
    Performs a read-only recursive scan of a specified SharePoint
    document library folder to identify empty-folder candidates.

.DESCRIPTION
    Connects to SharePoint Online through Microsoft Graph PowerShell
    and recursively enumerates folders within the configured target
    directory.

    Identifies folders containing no child items and exports the
    results to a timestamped CSV report. Scan errors are recorded
    separately to ensure that folders with failed scans are not
    incorrectly classified as empty.

    The script does not delete, move, or modify SharePoint content.

.OUTPUTS
    Empty-folder candidate CSV and scan-error CSV.

.NOTES
    Requires PowerShell 7 and the Microsoft Graph PowerShell modules.
    Requires appropriate Microsoft Graph delegated permissions.

    Review the error report and candidate CSV before proceeding
    with any deletion operation.
#>


$ErrorActionPreference = "Stop"

# ============================================================
# CONFIGURATION
# ============================================================

$SiteUrl      = "https://example.com/sites/example"
$LibraryName  = "Documents"
$TargetFolder = "Directory/Folder Name"
$OutputFolder = "C:\Scripting\"

# ============================================================
# INITIALIZATION
# ============================================================

New-Item -Path $OutputFolder -ItemType Directory -Force |
    Out-Null

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

$EmptyCsv = Join-Path $OutputFolder `
    "EmptyFolders_$Timestamp.csv"

$ErrorCsv = Join-Path $OutputFolder `
    "EmptyFolders_Errors_$Timestamp.csv"

$EmptyFolders = [System.Collections.Generic.List[object]]::new()
$ScanErrors   = [System.Collections.Generic.List[object]]::new()

# ============================================================
# REQUIRED MODULES AND AUTHENTICATION
# ============================================================

$RequiredModules = @(
    "Microsoft.Graph.Authentication"
    "Microsoft.Graph.Sites"
    "Microsoft.Graph.Files"
)

foreach ($ModuleName in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        throw "Missing module: $ModuleName. Install it before running this script."
    }

    Import-Module $ModuleName -ErrorAction Stop
}

if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes "Sites.Read.All" -NoWelcome
}

if (-not (Get-MgContext)) {
    throw "Microsoft Graph authentication was not established."
}

# ============================================================
# HELPER: READ A FIELD FROM A GRAPH RESPONSE
# Supports both dictionaries and PowerShell objects.
# ============================================================

function Get-GraphField {
    param(
        [Parameter(Mandatory)]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name)) {
            return $InputObject[$Name]
        }

        return $null
    }

    $Property = $InputObject.PSObject.Properties[$Name]

    if ($null -ne $Property) {
        return $Property.Value
    }

    return $null
}

# ============================================================
# HELPER: DETECT A REAL FOLDER FACET
# A file must never be added to the folder queue.
# ============================================================

function Test-GraphFolder {
    param(
        [Parameter(Mandatory)]
        [object]$Item
    )

    $FolderFacet = Get-GraphField -InputObject $Item -Name "folder"

    return ($null -ne $FolderFacet)
}

# ============================================================
# HELPER: GET ALL CHILDREN WITH PAGINATION
# Uses raw Graph responses and follows @odata.nextLink.
# Errors are thrown to the caller, never treated as empty.
# ============================================================

function Get-GraphChildren {
    param(
        [Parameter(Mandatory)]
        [string]$DriveId,

        [Parameter(Mandatory)]
        [string]$ItemId
    )

    $EncodedDriveId = [System.Uri]::EscapeDataString($DriveId)
    $EncodedItemId  = [System.Uri]::EscapeDataString($ItemId)

    $RequestUri = "https://graph.microsoft.com/v1.0/drives/$EncodedDriveId/items/$EncodedItemId/children?`$select=id,name,webUrl,folder,file,parentReference&`$top=200"

    $AllItems = [System.Collections.Generic.List[object]]::new()

    while (-not [string]::IsNullOrWhiteSpace($RequestUri)) {

        $Response = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $RequestUri `
            -ErrorAction Stop

        $PageItems = Get-GraphField `
            -InputObject $Response `
            -Name "value"

        foreach ($Item in @($PageItems)) {
            if ($null -ne $Item) {
                $AllItems.Add($Item)
            }
        }

        $NextLink = Get-GraphField `
            -InputObject $Response `
            -Name "@odata.nextLink"

        if ($null -ne $NextLink -and
            -not [string]::IsNullOrWhiteSpace([string]$NextLink)) {

            $RequestUri = [string]$NextLink
        }
        else {
            $RequestUri = $null
        }
    }

    return $AllItems.ToArray()
}

# ============================================================
# RESOLVE SITE AND LIBRARY
# ============================================================

$Uri = [System.Uri]$SiteUrl
$SitePath = $Uri.AbsolutePath.TrimEnd("/")

Write-Host "`nResolving SharePoint site..." -ForegroundColor Cyan

$Site = Get-MgSite -SiteId "$($Uri.Host):$SitePath"

if (-not $Site) {
    throw "Could not resolve site: $SiteUrl"
}

Write-Host "Site: $($Site.WebUrl)" -ForegroundColor Green

$Drives = Get-MgSiteDrive -SiteId $Site.Id -All

$Drive = $Drives |
    Where-Object { $_.Name.Trim() -ieq $LibraryName } |
    Select-Object -First 1

if (-not $Drive) {
    Write-Host "`nAvailable libraries:" -ForegroundColor Yellow
    $Drives | Select-Object Name, Id | Format-Table -AutoSize

    throw "Library '$LibraryName' was not found."
}

Write-Host "Library: $($Drive.Name)" -ForegroundColor Green

# ============================================================
# FIND TARGET FOLDER AT LIBRARY ROOT
# ============================================================

Write-Host "`nLocating target folder..." -ForegroundColor Cyan

$Root = Get-MgDriveRoot -DriveId $Drive.Id

$RootChildren = @(
    Get-GraphChildren -DriveId $Drive.Id -ItemId $Root.Id
)

$Target = $RootChildren |
    Where-Object {
        (Get-GraphField -InputObject $_ -Name "name") -ieq $TargetFolder -and
        (Test-GraphFolder -Item $_)
    } |
    Select-Object -First 1

if (-not $Target) {
    Write-Host "`nRoot items returned by Graph:" -ForegroundColor Yellow

    $RootItemsReport = foreach ($Item in $RootChildren) {
    [PSCustomObject]@{
        Name = Get-GraphField -InputObject $Item -Name "name"
        Type = if (Test-GraphFolder -Item $Item) {
            "Folder"
        }
        else {
            "Not a folder"
        }
    }
}

$RootItemsReport | Format-Table -AutoSize

    throw "Target folder '$TargetFolder' was not found at the library root."
}

$TargetId   = [string](Get-GraphField $Target "id")
$TargetName = [string](Get-GraphField $Target "name")
$TargetUrl  = [string](Get-GraphField $Target "webUrl")
$TargetPath = "$LibraryName/$TargetName"

Write-Host "Target: $TargetPath" -ForegroundColor Green

# ============================================================
# READ-ONLY RECURSIVE SCAN
# ============================================================

$Queue = [System.Collections.Generic.Queue[object]]::new()

$Queue.Enqueue([PSCustomObject]@{
    Id   = $TargetId
    Name = $TargetName
    Path = $TargetPath
    Url  = $TargetUrl
})

$FoldersInspected = 0
$FolderErrors     = 0
$FilesSeen        = 0

Write-Host "`nStarting read-only scan..." -ForegroundColor Cyan

while ($Queue.Count -gt 0) {

    $Current = $Queue.Dequeue()
    $FoldersInspected++

    Write-Progress `
        -Activity "Scanning SharePoint folders" `
        -Status "Inspected: $FoldersInspected | Pending: $($Queue.Count) | $($Current.Path)"

    # SAFETY CHECK: Never call the children endpoint unless
    # the queued item's metadata confirms it is a folder.
    #
    # The target was confirmed as a folder, and children are
    # added to the queue only after checking their folder facet.

    try {
        $Children = @(
            Get-GraphChildren `
                -DriveId $Drive.Id `
                -ItemId $Current.Id
        )
    }
    catch {
        $FolderErrors++

        $ScanErrors.Add([PSCustomObject]@{
            FolderName = $Current.Name
            FolderPath = $Current.Path
            FolderId   = $Current.Id
            Error      = $_.Exception.Message
        })

        Write-Warning "Scan failed; NOT marking as empty: $($Current.Path)"
        continue
    }

    # A successfully retrieved zero-item collection is empty.
    # Any returned item, including Forms or a file, means the
    # folder is not empty.

    if ($Children.Count -eq 0) {

        $EmptyFolders.Add([PSCustomObject]@{
            FolderName = $Current.Name
            FolderPath = $Current.Path
            FolderId   = $Current.Id
            WebUrl     = $Current.Url
        })

        Write-Host "Empty-folder candidate: $($Current.Path)" `
            -ForegroundColor Yellow

        continue
    }

    foreach ($Child in $Children) {

        $ChildName = [string](Get-GraphField $Child "name")
        $ChildId   = [string](Get-GraphField $Child "id")
        $ChildUrl  = [string](Get-GraphField $Child "webUrl")

        # Files count as content and are never queued.
        if (-not (Test-GraphFolder -Item $Child)) {
            $FilesSeen++
            continue
        }

        # Skip recursion into the SharePoint Forms system folder.
        # It still counts as content in its parent folder.
        if ($ChildName -ieq "Forms") {
            continue
        }

        # Only verified folders enter the queue.
        $ChildPath = "$($Current.Path)/$ChildName"

        $Queue.Enqueue([PSCustomObject]@{
            Id   = $ChildId
            Name = $ChildName
            Path = $ChildPath
            Url  = $ChildUrl
        })
    }
}

Write-Progress -Activity "Scanning SharePoint folders" -Completed

# ============================================================
# EXPORT REPORTS
# ============================================================

if ($EmptyFolders.Count -gt 0) {
    $EmptyFolders |
        Export-Csv -Path $EmptyCsv -NoTypeInformation -Encoding utf8
}
else {
    '"FolderName","FolderPath","FolderId","WebUrl"' |
        Set-Content -Path $EmptyCsv -Encoding utf8
}

if ($ScanErrors.Count -gt 0) {
    $ScanErrors |
        Export-Csv -Path $ErrorCsv -NoTypeInformation -Encoding utf8
}
else {
    '"FolderName","FolderPath","FolderId","Error"' |
        Set-Content -Path $ErrorCsv -Encoding utf8
}

# ============================================================
# SUMMARY
# ============================================================

Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "          READ-ONLY SCAN SUMMARY            " -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

Write-Host "Site                : $($Site.WebUrl)"
Write-Host "Library             : $($Drive.Name)"
Write-Host "Target folder       : $TargetPath"
Write-Host "Folders inspected   : $FoldersInspected"
Write-Host "Files encountered   : $FilesSeen"
Write-Host "Folders with errors : $FolderErrors"
Write-Host "Empty candidates    : $($EmptyFolders.Count)"

Write-Host "`nEmpty-folder report:" -ForegroundColor Green
Write-Host $EmptyCsv

Write-Host "`nError report:" -ForegroundColor Green
Write-Host $ErrorCsv

if ($FolderErrors -gt 0) {
    Write-Host "`nAUDIT INCOMPLETE: Resolve errors before relying on results." `
        -ForegroundColor Red
}
else {
    Write-Host "`nScan completed without recorded errors." `
        -ForegroundColor Green
}

Write-Host "`nREAD-ONLY: No SharePoint content was deleted or modified." `
    -ForegroundColor Cyan
