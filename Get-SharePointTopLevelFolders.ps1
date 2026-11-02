
# Requires an existing Microsoft Graph connection
# READ-ONLY: Lists top-level folders only

$ErrorActionPreference = "Stop"

$SiteUrl     = "https://example.com/sites/example"
$LibraryName = "Documents"
$Timestamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputCsv   = ".\SharePointTopLevelFolders_$Timestamp.csv"

# Resolve site
$Uri = [System.Uri]$SiteUrl
$SitePath = $Uri.AbsolutePath.TrimEnd('/')

$Site = Get-MgSite -SiteId "$($Uri.Host):$SitePath" -ErrorAction Stop

# Resolve library
$Drives = Get-MgSiteDrive -SiteId $Site.Id -All -ErrorAction Stop

$Drive = $Drives |
    Where-Object { $_.Name.Trim() -ieq $LibraryName } |
    Select-Object -First 1

if (-not $Drive) {
    throw "Library '$LibraryName' not found."
}

# Get library root
$Root = Get-MgDriveRoot -DriveId $Drive.Id -ErrorAction Stop

# Enumerate immediate children
$Children = @(
    Get-MgDriveItemChild `
        -DriveId $Drive.Id `
        -DriveItemId $Root.Id `
        -All `
        -Property "id,name,webUrl,folder,file,parentReference" `
        -ErrorAction Stop
)

# Record top-level folders only
$Inventory = @(
    foreach ($Item in $Children) {
        if ($null -ne $Item.Folder -and $Item.Name -ine "Forms") {
            [PSCustomObject]@{
                FolderName = $Item.Name
                FolderId   = $Item.Id
                WebUrl     = $Item.WebUrl
                DriveId    = $Drive.Id
                SiteUrl    = $SiteUrl
                Library    = $LibraryName
            }
        }
    }
)

$Inventory |
    Sort-Object FolderName |
    Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "Top-level folder inventory completed." -ForegroundColor Green
Write-Host "Folders found: $($Inventory.Count)"
Write-Host "Report: $((Resolve-Path $OutputCsv).Path)"
Write-Host "No SharePoint items were modified." -ForegroundColor Green

$Inventory | Format-Table FolderName, FolderId -AutoSize
