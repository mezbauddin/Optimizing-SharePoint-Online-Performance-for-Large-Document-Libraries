# ==============================
# Bulk Metadata Update using Raw Microsoft Graph API
# ==============================

# Config
$tenantId      = "<tenant-id>"
$clientId      = "<app-client-id>"
$clientSecret  = "<app-client-secret>"
$siteHost      = "contoso.sharepoint.com"
$sitePath      = "/sites/Research"
$listName      = "LargeDocs"
$batchSize     = 20
$maxRetries    = 3
$retryDelaySec = 5
$CsvPath       = "./documents-to-update.csv"

# Log helper
$LogFile = "./GraphAPI_BulkUpdate_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Level - $Message"
    Add-Content -Path $LogFile -Value $logMessage
    switch ($Level) {
        "WARNING" { Write-Warning $Message }
        "ERROR"   { Write-Error $Message }
        default   { Write-Host $Message }
    }
}

# -----------------------------
# Get OAuth2 token
# -----------------------------

function Get-GraphToken {
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }
    $resp = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $body
    return $resp.access_token
}

$Token = Get-GraphToken
$Headers = @{
    Authorization = "Bearer $Token"
    "Content-Type" = "application/json"
}

# -----------------------------
# Get Site and List IDs
# -----------------------------

$siteResp = Invoke-RestMethod -Headers $Headers -Uri "https://graph.microsoft.com/v1.0/sites/$siteHost:$sitePath"
$siteId = $siteResp.id

$listResp = Invoke-RestMethod -Headers $Headers -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists?$filter=displayName eq '$listName'"
$listId = $listResp.value[0].id

Write-Log "Resolved Site ID: $siteId, List ID: $listId"

# -----------------------------
# Load CSV
# -----------------------------

$rows = Import-Csv -Path $CsvPath
if (-not $rows) { throw "CSV empty or invalid." }

# -----------------------------
# Retry wrapper
# -----------------------------

function Invoke-WithRetry {
    param(
        [scriptblock]$Action,
        [int]$Retries = $maxRetries
    )
    for ($i = 1; $i -le $Retries; $i++) {
        try {
            return & $Action
        } catch {
            Write-Log "Attempt $i failed: $($_.Exception.Message)" "WARNING"
            Start-Sleep -Seconds $retryDelaySec
        }
    }
    throw "Action failed after $Retries retries."
}

# -----------------------------
# Prepare batches for Graph $batch endpoint
# -----------------------------

function Build-GraphBatch {
    param(
        [array]$Items,
        [string]$siteId,
        [string]$listId
    )
    $requests = @()
    $id = 1
    foreach ($row in $Items) {
        $fileName = $row.FileName
        # Use OData filter to find item by filename
        $uri = "/sites/$siteId/lists/$listId/items?`$filter=fields/FileLeafRef eq '$fileName'"
        $requests += @{
            id = "$id"
            method = "GET"
            url = $uri
        }
        $id++
    }
    return @{ requests = $requests }
}

# -----------------------------
# Execute batch request
# -----------------------------

function Invoke-GraphBatch {
    param(
        [hashtable]$BatchBody
    )
    $jsonBody = $BatchBody | ConvertTo-Json -Depth 10
    $resp = Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/$batch" -Headers $Headers -Body $jsonBody
    return $resp
}

# -----------------------------
# Main Loop
# -----------------------------

$total = $rows.Count
$counter = 0

for ($i = 0; $i -lt $total; $i += $batchSize) {
    $slice = $rows[$i..([Math]::Min($i + $batchSize - 1, $total - 1))]
    Write-Log "Processing batch $([int]($i / $batchSize) + 1) with $($slice.Count) items"

    # Step 1: Build batch to retrieve list items
    $batchBody = Build-GraphBatch -Items $slice -siteId $siteId -listId $listId
    $batchResp = Invoke-WithRetry { Invoke-GraphBatch -BatchBody $batchBody }

    # Step 2: Update each item
    foreach ($respItem in $batchResp.responses) {
        if ($respItem.status -eq 200 -and $respItem.body.value.Count -gt 0) {
            $itemId = $respItem.body.value[0].id
            $newStatus = $slice[$respItem.id - 1].NewStatus
            $patchBody = @{ fields = @{ Status = $newStatus } } | ConvertTo-Json
            Invoke-WithRetry {
                Invoke-RestMethod -Method Patch -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$listId/items/$itemId" -Headers $Headers -Body $patchBody
            }
            Write-Log "Updated $($slice[$respItem.id - 1].FileName) -> $newStatus"
        } else {
            Write-Log "Item not found for batch request ID $($respItem.id)" "WARNING"
        }
        $counter++
        Write-Log "Progress: $counter / $total"
    }
}
