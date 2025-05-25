# ========== Configuration ==========
$InputFile = "ObjectId.txt"  # Ensure this is in the same directory or update the path
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputFile = "SignInOutput_$Timestamp.csv"

# Read environment variables (set via Azure DevOps)
$TenantId = $env:TenantId
$ClientId = $env:ClientId
$ClientSecret = $env:ClientSecret

# Required modules
if (-not (Get-Module -ListAvailable -Name "Microsoft.Graph")) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}
Import-Module Microsoft.Graph

# Authenticate using Azure.Identity.ClientSecretCredential
Add-Type -Path "$env:HOME/.local/share/powershell/Modules/Azure.Identity/*/lib/netstandard2.0/Azure.Identity.dll"
$credential = [Azure.Identity.ClientSecretCredential]::new($TenantId, $ClientId, $ClientSecret)
Connect-MgGraph -ClientSecretCredential $credential -Scopes "https://graph.microsoft.com/.default"

# ===== Read Input IDs =====
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

# ===== Fetch Sign-In Logs =====
$startDate = (Get-Date).AddDays(-30).ToString("yyyy-MM-dd")
$logs = @()
$uri = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=createdDateTime ge $startDate"

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

# ===== Filter logs and export =====
$filteredLogs = $logs | Where-Object { $userIds -contains $_.userId }

$latestLogs = $filteredLogs |
    Sort-Object createdDateTime -Descending |
    Group-Object userId |
    ForEach-Object { $_.Group | Select-Object -First 1 }

$final = $latestLogs | ForEach-Object {
    [PSCustomObject]@{
        UserId      = $_.userId
        DisplayName = $_.userDisplayName
        SignInTime  = $_.createdDateTime
        IP          = $_.ipAddress
        ClientApp   = $_.clientAppUsed
        Status      = $_.status.errorCode
    }
}

if ($final.Count -gt 0) {
    $final | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
    Write-Host "`n✅ Exported latest sign-ins to: $OutputFile" -ForegroundColor Green
} else {
    Write-Host "`n⚠️ No sign-in logs found for the provided Object IDs." -ForegroundColor Yellow
}

Disconnect-MgGraph
