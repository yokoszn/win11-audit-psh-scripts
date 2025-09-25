\# Windows 11 Migration — Minimal Audit



```

win11-audit-min/

├─ README.md

└─ scripts/

&nbsp;  ├─ 01-AD-DeviceInventory.ps1

&nbsp;  └─ 02-Entra-SignInsAgg.ps1

```



Two read-only scripts:



1\) `01-AD-DeviceInventory.ps1`  

&nbsp;  • Enumerates AD computers and snapshots OS/build, CPU, RAM, disk, TPM, Secure Boot, UEFI via CIM.  

&nbsp;  • No changes. Works with or without RSAT `ActiveDirectory` module.



2\) `02-Entra-SignInsAgg.ps1`  

&nbsp;  • Pulls Entra (Azure AD) sign-in logs and aggregates per-user \*\*MostUsedDevice\*\* and \*\*LastSeenDevice\*\*.  

&nbsp;  • Read-only scopes: `AuditLog.Read.All` and `Directory.Read.All`.



\## Quick start

```powershell

\# 1) AD device + OS inventory (read-only)

.\\scripts\\01-AD-DeviceInventory.ps1 -OutCsv .\\AD\_Device\_Inventory.csv



\# 2) Entra sign-in aggregation (read-only; you’ll be prompted to sign in)

.\\scripts\\02-Entra-SignInsAgg.ps1 -LookbackDays 90 -OutCsv .\\Entra\_Device\_Agg.csv





Output schemas



AD\_Device\_Inventory.csv

Hostname,Domain,Manufacturer,Model,CPUModel,Cores,RAMGB,DiskGB,SysDriveFreeGB,OSName,OSEdition,OSVersion,OSBuild,UEFI,SecureBoot,TPMPresent,TPMVersion



Entra\_Device\_Agg.csv

UPN,DisplayName,MostUsedDevice,MostUsedDeviceOS,MostUsedDeviceCount,LastSeenDevice,LastSeenDeviceOS,LastSeenAt,TotalSignIns



Notes



Run 01 as a domain user with read access; CIM requires the account to query remote WMI/CIM (no writes).



02 requires Global Admin only to consent once to read scopes; ongoing runs are read-only.



Keep timeouts short; unreachable hosts are skipped.

