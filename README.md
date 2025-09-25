
```
├─ README.md
└─ scripts/
 ├─ 01-AD-DeviceInventory.ps1
 └─ 02-Entra-SignInsAgg.ps1
```



# Two read-only scripts:



1\) `01-AD-DeviceInventory.ps1`  

 • Enumerates AD computers and snapshots OS/build, CPU, RAM, disk, TPM, Secure Boot, UEFI via CIM.  

 • No changes. Works with or without RSAT `ActiveDirectory` module.



2\) `02-Entra-SignInsAgg.ps1`  

 • Pulls Entra (Azure AD) sign-in logs and aggregates per-user \*\*MostUsedDevice\*\* and \*\*LastSeenDevice\*\*.  

• Read-only scopes: `AuditLog.Read.All` and `Directory.Read.All`.



\## Quick start

```powershell

\# 1) AD device + OS inventory (read-only)

.\\scripts\\01-AD-DeviceInventory.ps1 -OutCsv .\\AD\_Device\_Inventory.csv



\# 2) Entra sign-in aggregation (read-only; you’ll be prompted to sign in)

.\\scripts\\02-Entra-SignInsAgg.ps1 -LookbackDays 90 -OutCsv .\\Entra\_Device\_Agg.csv
```

- If you’re a domain admin on the server:
```
.\01-AD-DeviceInventory.ps1
```
If you’re a local admin (not domain):
```
$cred = Get-Credential "PRIMEDESIGN\SomeAdmin"
.\01-AD-DeviceInventory.ps1 -Credential $cred
```
To hide unreachable rows:
```
.\01-AD-DeviceInventory.ps1 -SkipUnreachable
```

This avoids WinRM setup. If you still get mostly Reachable = False, it means RPC/DCOM is blocked or the remote WMI service is disabled. In that case, you can still ship the CSV (it will list which hosts failed and why) without touching settings, or run it from a host with firewall rules that allow RPC/WMI to clients.

## Output schemas


### AD\_Device\_Inventory.csv
```
Hostname,Domain,Manufacturer,Model,CPUModel,Cores,RAMGB,DiskGB,SysDriveFreeGB,OSName,OSEdition,OSVersion,OSBuild,UEFI,SecureBoot,TPMPresent,TPMVersion
```

### Entra\_Device\_Agg.csv
```
UPN,DisplayName,MostUsedDevice,MostUsedDeviceOS,MostUsedDeviceCount,LastSeenDevice,LastSeenDeviceOS,LastSeenAt,TotalSignIns
```



