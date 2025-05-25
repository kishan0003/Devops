# ========== Configuration ==========
$InputFile = "ObjectId.txt"
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputFile = "SignInoutput_$Timestamp.csv"

# Environment Variables (set in Azure DevOps pipeline)
$TenantId = $env:TenantId
$ClientId = $env:ClientId
$ClientSecret = $env:ClientSecret

# Microsoft Graph scope
$scope = "https://graph.microsoft.com/.default"
# ===================================

# ========== Authentication ==========
# Request access token using client credentials
$tokenBody = @{
    client_id     = $ClientId
    scope         = $scope
    client_secret = $ClientSecret
    grant_type    = "client_credentials"
}
try {
    $tokenResponse = Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -Body $tokenBody -ErrorAction Stop
    $accessToken = $tokenResponse.access_token
} catch {
    Write-Error "❌ Failed to retrieve access token. $_"
    exit 1
}

# Connect to Microsoft Graph with access token
Connect-MgGraph -AccessToken $accessToken
# =============================================

# ========== Input Validation ==========
if (-Not (Test-Path $InputFile)) {
    Write-Error "❌ Input file not found: $InputFile"
    Disconnect-MgGraph
    exit 1
}
$userIds = Get-Content -Path $InputFile | Where-Object { $_ -match '^[a-f0-9\-]{36}$' }
if (-not $userIds) {
    Write-Error "❌ No valid Object IDs found in input file."
    Disconnect-MgGraph
    exit 1
}
# ======================================

# Prepare output CSV file with headers
[PSCustomObject]@{
    UserId      = ''
    DisplayName = ''
    SignInTime  = ''
    IP          = ''
    ClientApp   = ''
    Status      = ''
} | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

# Initial Graph API URL (last 30 days filter)
$startDate = (Get-Date).AddDays(-30).ToString("yyyy-MM-dd")
$uri = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=createdDateTime ge $startDate"

# Process logs page by page
do {
    try {
        $response = Invoke-MgGraphRequest -Uri $uri -Method GET
    } catch {
        Write-Error "❌ Failed to fetch sign-in logs from URI: $uri. Error: $_"
        Disconnect-MgGraph
        exit 1
    }

    # Filter logs for specific user Object IDs
    $filtered = $response.value | Where-Object { $userIds -contains $_.userId }

    # Format filtered logs for output
    $entries = $filtered | ForEach-Object {
        [PSCustomObject]@{
            UserId      = $_.userId
            DisplayName = $_.userDisplayName
            SignInTime  = $_.createdDateTime
            IP          = $_.ipAddress
            ClientApp   = $_.clientAppUsed
            Status      = $_.status.errorCode
        }
    }

    # Append entries to output file
    if ($entries.Count -gt 0) {
        $entries | Export-Csv -Path $OutputFile -Append -NoTypeInformation -Encoding UTF8
        Write-Host "✅ Processed and saved $($entries.Count) log(s)..."
    }

    $uri = $response.'@odata.nextLink'
} while ($uri)

Write-Host "`n✅ Finished exporting filtered sign-in logs to: $OutputFile" -ForegroundColor Green

# Disconnect from Microsoft Graph
Disconnect-MgGraph
