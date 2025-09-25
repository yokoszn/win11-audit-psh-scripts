[CmdletBinding()]
param(
  [string]$OutCsv = ".\AD_Device_Inventory.csv",
  [int]$TimeoutSec = 6,
  [switch]$SkipUnreachable,
  [PSCredential]$Credential
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Get-AdComputerDns {
  # Prefer RSAT ActiveDirectory if installed; else pure LDAP
  if (Get-Module -ListAvailable -Name ActiveDirectory) {
    try {
      Import-Module ActiveDirectory -ErrorAction Stop
      return (Get-ADComputer -Filter * -Properties DNSHostName |
        Where-Object { $_.DNSHostName } |
        Select-Object -ExpandProperty DNSHostName)
    } catch {
      # fall through to LDAP
    }
  }

  # Fallback: LDAP
  try {
    $root = [ADSI]"LDAP://RootDSE"
    $base = "LDAP://{0}" -f $root.defaultNamingContext
    $ds = New-Object System.DirectoryServices.DirectorySearcher([ADSI]$base)
    $ds.Filter = "(&(objectClass=computer)(dnsHostName=*))"
    $null = $ds.PropertiesToLoad.Add("dnshostname")
    $ds.PageSize = 1000
    $results = $ds.FindAll()
    $list = foreach ($r in $results) { $r.Properties['dnshostname'][0] }
    $list | Where-Object { $_ }
  } catch {
    throw "Failed to enumerate computers from AD/LDAP. $($_.Exception.Message)"
  }
}

function New-BestCimSession {
  param(
    [Parameter(Mandatory)][string]$ComputerName,
    [int]$TimeoutSec = 6,
    [PSCredential]$Credential
  )
  # Try WSMan first
  try {
    if ($Credential) {
      return New-CimSession -ComputerName $ComputerName -Credential $Credential -OperationTimeoutSec $TimeoutSec -ErrorAction Stop
    } else {
      return New-CimSession -ComputerName $ComputerName -OperationTimeoutSec $TimeoutSec -ErrorAction Stop
    }
  } catch {
    # Fallback to DCOM (no WinRM needed)
    $opt = New-CimSessionOption -Protocol Dcom
    if ($Credential) {
      return New-CimSession -ComputerName $ComputerName -SessionOption $opt -Credential $Credential -OperationTimeoutSec $TimeoutSec -ErrorAction Stop
    } else {
      return New-CimSession -ComputerName $ComputerName -SessionOption $opt -OperationTimeoutSec $TimeoutSec -ErrorAction Stop
    }
  }
}

function Get-CimSnapshot {
  param(
    [Parameter(Mandatory)][string]$ComputerName,
    [int]$TimeoutSec = 6,
    [PSCredential]$Credential
  )

  $s = $null
  try {
    $s = New-BestCimSession -ComputerName $ComputerName -TimeoutSec $TimeoutSec -Credential $Credential
  } catch {
    if ($script:SkipUnreachable) { return $null }
    $msg = $_.Exception.Message
    # Emit an "unreachable" row so the CSV is still useful
    return [PSCustomObject]@{
      Hostname       = $ComputerName
      Domain         = $null
      Manufacturer   = $null
      Model          = $null
      CPUModel       = $null
      Cores          = $null
      RAMGB          = $null
      DiskGB         = $null
      SysDriveFreeGB = $null
      OSName         = $null
      OSEdition      = $null
      OSVersion      = $null
      OSBuild        = $null
      UEFI           = $null
      SecureBoot     = $null
      TPMPresent     = $null
      TPMVersion     = $null
      Reachable      = $false
      Error          = $msg
    }
  }

  try {
    $os  = Get-CimInstance -CimSession $s -Class Win32_OperatingSystem -ErrorAction SilentlyContinue
    $cs  = Get-CimInstance -CimSession $s -Class Win32_ComputerSystem -ErrorAction SilentlyContinue
    $cpu = Get-CimInstance -CimSession $s -Class Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    $c   = Get-CimInstance -CimSession $s -Class Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
    $sb  = Get-CimInstance -CimSession $s -Namespace root\wmi -Class MS_SecureBoot -ErrorAction SilentlyContinue
    $tpm = Get-CimInstance -CimSession $s -Namespace root\CIMV2\Security\MicrosoftTpm -Class Win32_Tpm -ErrorAction SilentlyContinue
  } finally {
    if ($s) { Remove-CimSession $s }
  }

  if (-not $os) {
    if ($script:SkipUnreachable) { return $null }
    return [PSCustomObject]@{
      Hostname       = $ComputerName
      Domain         = $null
      Manufacturer   = $null
      Model          = $null
      CPUModel       = $null
      Cores          = $null
      RAMGB          = $null
      DiskGB         = $null
      SysDriveFreeGB = $null
      OSName         = $null
      OSEdition      = $null
      OSVersion      = $null
      OSBuild        = $null
      UEFI           = $null
      SecureBoot     = $null
      TPMPresent     = $null
      TPMVersion     = $null
      Reachable      = $false
      Error          = "Connected but no OS data (permissions or WMI service issue)"
    }
  }

  # Derive fields that need branching
  $uefi = $false
  $secureBoot = $null
  if ($sb) {
    $uefi = $true
    try { $secureBoot = [bool]$sb.SecureBootEnabled } catch { $secureBoot = $null }
  }

  $tpmPresent = $false
  $tpmVersion = $null
  if ($tpm) {
    $tpmPresent = $true
    try { $tpmVersion = ($tpm.SpecVersion -join ',') } catch { $tpmVersion = $null }
  }

  [PSCustomObject]@{
    Hostname       = $ComputerName
    Domain         = $cs.Domain
    Manufacturer   = $cs.Manufacturer
    Model          = $cs.Model
    CPUModel       = $cpu.Name
    Cores          = $cpu.NumberOfCores
    RAMGB          = [math]::Round(($cs.TotalPhysicalMemory / 1GB), 1)
    DiskGB         = if ($c) { [math]::Round(($c.Size / 1GB), 0) } else { $null }
    SysDriveFreeGB = if ($c) { [math]::Round(($c.FreeSpace / 1GB), 0) } else { $null }
    OSName         = $os.Caption
    OSEdition      = $os.OperatingSystemSKU
    OSVersion      = $os.Version
    OSBuild        = $os.BuildNumber
    UEFI           = $uefi
    SecureBoot     = $secureBoot
    TPMPresent     = $tpmPresent
    TPMVersion     = $tpmVersion
    Reachable      = $true
    Error          = $null
  }
}

Write-Host "Enumerating AD computers..." -ForegroundColor Cyan
$hosts = @(Get-AdComputerDns | Sort-Object -Unique)
Write-Host ("Computers found: {0}" -f $hosts.Count) -ForegroundColor Cyan

# Collect results (list -> fast append)
$res = New-Object System.Collections.Generic.List[object]
foreach ($h in $hosts) {
  $snap = Get-CimSnapshot -ComputerName $h -TimeoutSec $TimeoutSec -Credential $Credential
  if ($snap) { [void]$res.Add($snap) }
}

# Ensure output directory exists
$dir = Split-Path -Path $OutCsv -Parent
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }

# Always export CSV (you'll see Reachable=$false rows if hosts failed)
if ($res.Count -gt 0) {
  $res | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
} else {
  # No hosts found at all: still emit headers
  [pscustomobject]@{
    Hostname=''; Domain=''; Manufacturer=''; Model=''; CPUModel=''; Cores=0; RAMGB=0;
    DiskGB=0; SysDriveFreeGB=0; OSName=''; OSEdition=''; OSVersion=''; OSBuild='';
    UEFI=$false; SecureBoot=$null; TPMPresent=$false; TPMVersion=''; Reachable=$false; Error='No hosts found'
  } |
  Select-Object Hostname,Domain,Manufacturer,Model,CPUModel,Cores,RAMGB,DiskGB,SysDriveFreeGB,OSName,OSEdition,OSVersion,OSBuild,UEFI,SecureBoot,TPMPresent,TPMVersion,Reachable,Error |
  Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
}

Write-Host ("AD device inventory written: {0}" -f $OutCsv) -ForegroundColor Green
