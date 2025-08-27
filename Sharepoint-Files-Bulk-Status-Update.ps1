# --- Modern PnP.PowerShell Script for Bulk Metadata Updates (2025) ---

# Parameters
$SiteUrl      = "https://tenant.sharepoint.com/sites/sitename"
$LibraryName  = "<LibraryName>"
$appId = "<EntraAppClientID>"
$CsvPath      = "./documents-to-update.csv"

# Connect using secure interactive login with new app
Connect-PnPOnline -Url $SiteUrl -ClientId $appId -Interactive

# Import CSV
$ItemsToUpdate = Import-Csv -Path $CsvPath
if (-not $ItemsToUpdate) {
    Write-Error "CSV file empty or invalid. Exiting."
    exit
}

# Get list reference
$List = Get-PnPList -Identity $LibraryName

# Track progress
$Total   = $ItemsToUpdate.Count
$Counter = 0

Write-Host "Processing $Total items..." -ForegroundColor Cyan

foreach ($Row in $ItemsToUpdate) {
    $Counter++
    $FileName  = $Row.FileName
    $NewStatus = $Row.NewStatus

    # Find the list item
    $ListItem = Get-PnPListItem -List $List -Fields "FileLeafRef" `
                | Where-Object { $_["FileLeafRef"] -eq $FileName }

    if ($ListItem) {
        # Direct update without batching
        Set-PnPListItem -List $List -Identity $ListItem.Id -Values @{ "Status" = $NewStatus }
        Write-Host "[$Counter/$Total] Updated: $FileName -> Status: $NewStatus" -ForegroundColor Green
    }
    else {
        Write-Warning "File not found: $FileName"
    }
}

Disconnect-PnPOnline
Write-Host "Bulk update complete." -ForegroundColor Cyan
