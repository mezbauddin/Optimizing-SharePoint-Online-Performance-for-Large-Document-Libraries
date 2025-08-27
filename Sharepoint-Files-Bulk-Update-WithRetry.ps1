param( 
    [string]$SiteUrl = "https://tenant.sharepoint.com/sites/sitename", 
    [string]$LibraryName = "<LibraryName>", 
    [string]$CsvPath = "./documents-to-update.csv", 
    [int]$BatchSize = 10, 
    [int]$MaxRetries = 5, 
    [string]$ClientId = "<EntraAppClientID>",
    [string]$FolderServerRelativeUrl = "<FolderPath>"
) 

# Log file setup
$LogFile = "./BulkUpdate_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param($Message, $Level = "INFO")
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Level - $Message"
    Add-Content -Path $LogFile -Value $logMessage
    switch ($Level) {
        "WARNING" { Write-Warning $Message }
        "ERROR"   { Write-Error $Message }
        default    { Write-Host $Message }
    }
}

try {
    # Connect to SharePoint
    Write-Log "Connecting to SharePoint site: $SiteUrl"
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive -ErrorAction Stop
    
    # Import CSV
    Write-Log "Reading CSV file: $CsvPath"
    $rows = Import-Csv -Path $CsvPath -ErrorAction Stop
    if (-not $rows) { throw "CSV is empty or cannot be read." }
    
    # Get list reference
    $list = Get-PnPList -Identity $LibraryName -ErrorAction Stop
    
    $updated = 0 
    $notFound = 0
    $errors = 0
    
    Write-Log "Starting bulk update of $($rows.Count) items..."
    
    # Process in batches
    for ($i = 0; $i -lt $rows.Count; $i += $BatchSize) { 
        $slice = $rows[$i..([Math]::Min($i + $BatchSize - 1, $rows.Count - 1))] 
        $batch = New-PnPBatch 
        $queued = 0 
        
        # Process each item in the current batch
        foreach ($r in $slice) { 
            $fileName = $r.FileName.Trim()
            $newStatus = $r.NewStatus
            $fileUrl = "$FolderServerRelativeUrl/$fileName"
            
            try {
                # Try to get the file directly by URL first (most efficient)
                $file = Get-PnPFile -Url $fileUrl -ErrorAction SilentlyContinue
                
                if ($file) {
                    $item = Get-PnPFile -Url $fileUrl -AsListItem -ErrorAction Stop
                    Set-PnPListItem -List $list -Identity $item.Id -Values @{ "Status" = $newStatus } -Batch $batch -ErrorAction Stop
                    $queued++
                    Write-Log "Queued for update: $fileName -> $newStatus"
                }
                else {
                    $notFound++
                    Write-Log "File not found: $fileName" "WARNING"
                }
            }
            catch {
                $errors++
                Write-Log "Error processing $fileName : $_" "ERROR"
            }
        }
        
        # Process the batch with retry logic
        if ($queued -gt 0) {
            $attempt = 0 
            $success = $false
            
            while (-not $success -and $attempt -lt $MaxRetries) { 
                try { 
                    $attempt++
                    Write-Log "Processing batch $([int]($i / $BatchSize) + 1) (attempt $attempt of $MaxRetries) with $queued items..."
                    
                    Invoke-PnPBatch -Batch $batch -ErrorAction Stop
                    $updated += $queued
                    $success = $true
                    
                    Write-Log "Successfully updated batch $([int]($i / $BatchSize) + 1). Total updated: $updated"
                    
                    # Add a small delay between batches
                    if (($i + $BatchSize) -lt $rows.Count) {
                        Start-Sleep -Milliseconds 1000
                    }
                } 
                catch { 
                    $delay = 60  # Default delay in seconds
                    $m = [regex]::Match($_.Exception.Message, 'Retry-After[\s:]+(\d+)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase) 
                    if ($m.Success) { 
                        $delay = [int]$m.Groups[1].Value 
                    }
                    
                    Write-Log "Throttling detected. Waiting $delay seconds before retry $attempt of $MaxRetries." "WARNING"
                    Start-Sleep -Seconds $delay
                    
                    if ($attempt -ge $MaxRetries) { 
                        Write-Log "Max retries reached for batch. Moving to next batch." "WARNING"
                        $errors += $queued
                    } 
                }
            }
        }
    } 
    
    # Generate summary
    $summary = @"

=== BULK UPDATE SUMMARY ===
Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Site: $SiteUrl
Library: $LibraryName
Folder: $FolderServerRelativeUrl

TOTAL ITEMS: $($rows.Count)
- Successfully updated: $updated
- Not found: $notFound
- Errors: $errors

Log file: $((Get-Item $LogFile).FullName)
"@
    
    Write-Log $summary
    
    if ($errors -gt 0) {
        Write-Log "There were $errors errors during processing. Please check the log file for details." "WARNING"
        exit 1
    }
}
catch {
    $errorMsg = "Fatal error: $_"
    Write-Log $errorMsg "ERROR"
    Write-Log $_.ScriptStackTrace "ERROR"
    exit 1
}
finally {
    # Disconnect from SharePoint
    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
    catch {
        Write-Log "Error disconnecting from SharePoint: $_" "WARNING"
    }
}
