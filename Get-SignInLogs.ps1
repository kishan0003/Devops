# ========== Configuration ==========
$InputFile = "ObjectId.txt"  # Make sure this file is in the same directory or update path
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputFile = "SignInoutput_$Timestamp.csv"

# Read from environment variables (passed from Azure DevOps pipeline)
$TenantId = $env:TenantId
$ClientId = $env:ClientId
$ClientSecret = $env:ClientSecret

# Microsoft Graph permission scope
$Scopes = @("https://graph.microsoft.com/.default")
# ===================================

# Install Microsoft.Graph if not already present
if (-not (Get-Module -ListAvailable -Name "Microsoft.Graph")) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}
Import-Module Microsoft.Graph

# Authenticate using client credentials (non-interactive)
$secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -ClientSecret $secureSecret

# Read and validate Object IDs from input file
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

# Fetch all sign-in logs with pagination
$logs = @()
$uri = "https://graph.microsoft.com/v1.0/auditLogs/signIns"

do {
    try {
        $response = Invoke-MgGraphRequest -Uri $uri -Method GET
    } catch {
        Write-Error "❌ Failed to fetch sign-in logs from URI: $uri. Error: $_"
        Disconnect-MgGraph
        exit 1
    }
    $logs += $response.value
    $uri = $response.'@odata.nextLink'
} while ($uri)

if (-not $logs) {
    Write-Host "❌ No sign-in logs returned from Microsoft Graph." -ForegroundColor Red
    Disconnect-MgGraph
    exit 1
}

# Filter logs for user IDs present in input list
$filteredLogs = $logs | Where-Object { $userIds -contains $_.userId }

# Get latest log per user
$latestLogs = $filteredLogs |
    Sort-Object createdDateTime -Descending |
    Group-Object userId |
    ForEach-Object { $_.Group | Select-Object -First 1 }

# Prepare final output
$final = $latestLogs | ForEach-Object {
    [PSCustomObject]@{
        UserId       = $_.userId
        DisplayName  = $_.userDisplayName
        SignInTime   = $_.createdDateTime
        IP           = $_.ipAddress
        ClientApp    = $_.clientAppUsed
        Status       = $_.status.errorCode
    }
}

# Export to CSV if results exist
if ($final.Count -gt 0) {
    $final | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
    Write-Host "`n✅ Exported latest sign-ins to: $OutputFile" -ForegroundColor Green
} else {
    Write-Host "`n⚠️ No sign-in logs found for the provided Object IDs." -ForegroundColor Yellow
}

# Disconnect session
Disconnect-MgGraph