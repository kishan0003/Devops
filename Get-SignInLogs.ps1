# ========== Configuration ==========
$InputFile = "ObjectId.txt"
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputFile = "SignInoutput_$Timestamp.csv"

# Environment Variables from Azure DevOps
$TenantId = $env:TenantId
$ClientId = $env:ClientId
$ClientSecret = $env:ClientSecret
$Scopes = @("https://graph.microsoft.com/.default")
# ===================================

# Install & import Microsoft.Graph if not present
if (-not (Get-Module -ListAvailable -Name "Microsoft.Graph")) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}
Import-Module Microsoft.Graph

# Authenticate using client credentials
$secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -ClientSecret $secureSecret

# Validate input file
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

# Sign-ins API URI (optional filter by date)
$startDate = (Get-Date).AddDays(-30).ToString("yyyy-MM-dd")
$uri = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=createdDateTime ge $startDate"

# Initialize CSV with headers
[PSCustomObject]@{
    UserId      = ''
    DisplayName = ''
    SignInTime  = ''
    IP          = ''
    ClientApp   = ''
    Status      = ''
} | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

# Stream and filter logs page by page
do {
    try {
        $response = Invoke-MgGraphRequest -Uri $uri -Method GET
    } catch {
        Write-Error "❌ Failed to fetch sign-in logs from URI: $uri. Error: $_"
        Disconnect-MgGraph
        exit 1
    }

    $filtered = $response.value | Where-Object { $userIds -contains $_.userId }

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

    if ($entries.Count -gt 0) {
        $entries | Export-Csv -Path $OutputFile -Append -NoTypeInformation -Encoding UTF8
        Write-Host "✅ Processed and saved $($entries.Count) log(s)..."
    }

    $uri = $response.'@odata.nextLink'
} while ($uri)

Write-Host "`n✅ Finished exporting filtered sign-in logs to: $OutputFile" -ForegroundColor Green

# Cleanup
Disconnect-MgGraph
