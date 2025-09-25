[CmdletBinding()]
param(
  [int]$LookbackDays = 90,
  [string]$OutCsv = ".\Entra_Device_Agg.csv"
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Minimal dependencies: Authentication + Reports only
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
  Write-Host "Installing Microsoft.Graph.Authentication..." -ForegroundColor Yellow
  Install-PackageProvider NuGet -Force | Out-Null
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
}
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Reports)) {
  Write-Host "Installing Microsoft.Graph.Reports..." -ForegroundColor Yellow
  Install-Module Microsoft.Graph.Reports -Scope CurrentUser -Force -AllowClobber
}

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Reports -ErrorAction Stop

$scopes = @("AuditLog.Read.All")
Write-Host "Connecting to Microsoft Graph (scopes: $($scopes -join ', '))..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $scopes | Out-Null

# Build a UTC timestamp without fractions (Graph requires this)
$startUtc = (Get-Date).AddDays(-$LookbackDays).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
Write-Host "Querying sign-ins since $startUtc (UTC) ..." -ForegroundColor Cyan

# Pull all sign-ins in window (tenant-wide)
$signins = Get-MgAuditLogSignIn -Filter "createdDateTime ge $startUtc" -All


if (-not $signins) {
  Write-Warning "No sign-in records returned for the lookback window."
  @() | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
  return
}

# Normalize → aggregate per UPN
$norm = $signins | ForEach-Object {
  [PSCustomObject]@{
    UPN               = $_.userPrincipalName
    DisplayName       = $_.userDisplayName
    Time              = [datetime]$_.createdDateTime
    DeviceDisplayName = $_.deviceDetail.displayName
    DeviceOS          = $_.deviceDetail.operatingSystem
  }
}

$agg = $norm | Group-Object UPN | ForEach-Object {
  $u = $_.Name
  $g = $_.Group
  $byDevice = $g | Group-Object DeviceDisplayName | Sort-Object Count -Descending
  $most = $byDevice | Select-Object -First 1
  $last = $g | Sort-Object Time -Descending | Select-Object -First 1

  [PSCustomObject]@{
    UPN                 = $u
    DisplayName         = ($g | Select-Object -First 1).DisplayName
    MostUsedDevice      = $most.Name
    MostUsedDeviceCount = $most.Count
    MostUsedDeviceOS    = ($g | Where-Object { $_.DeviceDisplayName -eq $most.Name } | Select-Object -First 1).DeviceOS
    LastSeenDevice      = $last.DeviceDisplayName
    LastSeenDeviceOS    = $last.DeviceOS
    LastSeenAt          = $last.Time
    TotalSignIns        = $g.Count
  }
}

$agg | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
Write-Host "Entra sign-in aggregation written: $OutCsv" -ForegroundColor Green
