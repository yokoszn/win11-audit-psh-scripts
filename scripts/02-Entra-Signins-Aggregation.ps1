[CmdletBinding()]
param(
  [int]$LookbackDays = 90,
  [string]$OutCsv = ".\Entra_Device_Agg.csv"
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# --- Modules (Auth + Reports + Directory) ---
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
  Write-Host "Installing Microsoft.Graph.Authentication..." -ForegroundColor Yellow
  Install-PackageProvider NuGet -Force | Out-Null
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
}
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Reports)) {
  Write-Host "Installing Microsoft.Graph.Reports..." -ForegroundColor Yellow
  Install-Module Microsoft.Graph.Reports -Scope CurrentUser -Force -AllowClobber
}
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Identity.DirectoryManagement)) {
  Write-Host "Installing Microsoft.Graph.Identity.DirectoryManagement..." -ForegroundColor Yellow
  Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser -Force -AllowClobber
}

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Reports -ErrorAction Stop
Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

# --- Connect with least-priv scopes needed ---
$scopes = @("AuditLog.Read.All","Directory.Read.All")
Write-Host "Connecting to Microsoft Graph (scopes: $($scopes -join ', '))..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $scopes | Out-Null

# --- Build UTC timestamp without fractions (Graph filter requirement) ---
$startUtc = (Get-Date).AddDays(-$LookbackDays).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
Write-Host "Querying sign-ins since $startUtc (UTC) ..." -ForegroundColor Cyan

# --- Sign-ins tenant-wide in window ---
try {
  $signins = Get-MgAuditLogSignIn -Filter "createdDateTime ge $startUtc" -All -ErrorAction Stop
} catch {
  throw "Graph sign-in query failed. Filter used: createdDateTime ge $startUtc . $($_.Exception.Message)"
}

if (-not $signins) {
  Write-Warning "No sign-in records returned for the lookback window."
  @() | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
  return
}

# --- Directory devices (for real OS + version) ---
Write-Host "Fetching directory devices for OS enrichment..." -ForegroundColor Cyan
$dirDevices = Get-MgDevice -All -Property "id,deviceId,displayName,operatingSystem,operatingSystemVersion"

# Hash by deviceId (GUID shown in sign-in deviceDetail.deviceId)
$byDeviceId = @{}
foreach ($d in $dirDevices) {
  if ($d.DeviceId) { $byDeviceId[$d.DeviceId] = $d }
}

function Get-DetectedOS {
  param(
    [string]$rawOS,        # from sign-in: deviceDetail.operatingSystem
    [string]$dirOS,        # from directory device
    [string]$dirOSVersion  # e.g. "10.0.22631.4037"
  )
  if ($dirOS -and $dirOSVersion) {
    try {
      $v = [version]$dirOSVersion
      if ($v.Build -ge 22000) { return "Windows 11" }
      else { return "Windows 10" }
    } catch {
      if ($dirOS) { return $dirOS } else { return $rawOS }
    }
  }

  if ($rawOS -match 'Windows\s*11') { return "Windows 11" }
  if ($rawOS -match 'Windows\s*10') { return "Windows 10" }
  if ($rawOS -match 'Windows')      { return "Windows" }
  return $rawOS
}

# --- Normalize → include corrected OS fields ---
$norm = $signins | ForEach-Object {
  $devId  = $_.deviceDetail.deviceId
  $rawOS  = $_.deviceDetail.operatingSystem
  $dirDev = $null
  if ($devId -and $byDeviceId.ContainsKey($devId)) { $dirDev = $byDeviceId[$devId] }

  $dirOS  = if ($dirDev) { $dirDev.operatingSystem } else { $null }
  $dirVer = if ($dirDev) { $dirDev.operatingSystemVersion } else { $null }
  $detOS  = Get-DetectedOS -rawOS $rawOS -dirOS $dirOS -dirOSVersion $dirVer

  [PSCustomObject]@{
    UPN                 = $_.userPrincipalName
    DisplayName         = $_.userDisplayName
    Time                = [datetime]$_.createdDateTime
    DeviceId            = $devId
    DeviceDisplayName   = $_.deviceDetail.displayName

    # Original OS from sign-in (often wrong on Win11)
    DeviceOS            = $rawOS

    # Directory truth
    Dir_DeviceOS        = $dirOS
    Dir_OSVersion       = $dirVer

    # Corrected classification (Win11 if build >= 22000)
    DetectedOS          = $detOS
  }
}

# --- Aggregate per UPN (use corrected OS for Most/Last) ---
$agg = $norm | Group-Object UPN | ForEach-Object {
  $u = $_.Name
  $g = $_.Group

  $byDevice = $g | Group-Object DeviceDisplayName | Sort-Object Count -Descending
  $most = $byDevice | Select-Object -First 1
  $last = $g | Sort-Object Time -Descending | Select-Object -First 1

  $mostRow = $null
  if ($most -and $most.Name) {
    $mostRow = $g | Where-Object { $_.DeviceDisplayName -eq $most.Name } | Select-Object -First 1
  }

  [PSCustomObject]@{
    UPN                       = $u
    DisplayName               = ($g | Select-Object -First 1).DisplayName

    MostUsedDevice            = if ($most) { $most.Name } else { $null }
    MostUsedDeviceCount       = if ($most) { $most.Count } else { 0 }
    MostUsedDeviceOS          = if ($mostRow) { $mostRow.DeviceOS } else { $null }               # raw
    MostUsedDeviceOSCorrected = if ($mostRow) { $mostRow.DetectedOS } else { $null }             # corrected

    LastSeenDevice            = $last.DeviceDisplayName
    LastSeenDeviceOS          = $last.DeviceOS                                                      # raw
    LastSeenOSCorrected       = $last.DetectedOS                                                    # corrected
    LastSeenOSVersion         = $last.Dir_OSVersion                                                 # directory version
    LastSeenAt                = $last.Time

    TotalSignIns              = $g.Count
  }
}

$agg | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
Write-Host "Entra sign-in aggregation written: $OutCsv" -ForegroundColor Green
