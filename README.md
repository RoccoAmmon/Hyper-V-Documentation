# Hyper-V-Documentation

PowerShell-Skript zur automatisierten Dokumentation einer Microsoft Hyper-V Umgebung
(Windows Server 2016/2019/2022/2025). Unterstützt Standalone-Hosts und
Hyper-V Failover-Cluster.

Das Ergebnis ist eine vollständige Umgebungsdokumentation als HTML, PDF oder Markdown.
Das HTML-Dokument lässt sich direkt in Microsoft Word importieren.

## Funktionen

- 58 auswählbare Dokumentations-Sektionen in 9 Kategorien
- Grafische Oberfläche (WPF) zur Auswahl von Hosts, Sektionen und Ausgabeformaten
- Ausgabeformate: **HTML**, **PDF**, **Markdown**
- Verbindung über CIM-Sessions mit automatischem **DCOM-Fallback** – funktioniert
  auch, wenn WinRM/PowerShell-Remoting nicht konfiguriert ist
- Automatische Cluster-Erkennung
- Automatischer Neustart mit erhöhten Rechten, wenn nicht als Administrator gestartet
- Vollständige Protokollierung in eine Logdatei inkl. Fehler- und Warnungszähler

## Voraussetzungen

| Komponente | Status | Hinweis |
| --- | --- | --- |
| Windows PowerShell 5.1 | erforderlich | |
| Hyper-V PowerShell-Modul (`RSAT-Hyper-V-Tools`) | erforderlich | |
| Administratorrechte auf den Hyper-V-Hosts | erforderlich | |
| RPC/DCOM oder WinRM erreichbar | erforderlich | WinRM wird bevorzugt, DCOM dient als Fallback |
| `FailoverClusters`-Modul | optional | für die Cluster-Sektionen |
| `ActiveDirectory`-Modul | optional | für AD- und FSMO-Sektionen |
| Microsoft Word **oder** Edge/Chrome | optional | für den PDF-Export |

## Verwendung

### Mit grafischer Oberfläche

```powershell
.\HyperV_Documentation.ps1
```

Ohne Angabe von `-HyperVServers` startet das Skript automatisch die GUI.
Alternativ lässt sie sich mit `-ShowGui` erzwingen.

### Ohne grafische Oberfläche

```powershell
.\HyperV_Documentation.ps1 -HyperVServers @("HV01","HV02") -CompanyName "Contoso GmbH"
```

```powershell
.\HyperV_Documentation.ps1 -HyperVServers @("HV01") `
    -CompanyName "Contoso GmbH" `
    -OutputPath "D:\Doku" `
    -OutputFormats @("HTML","PDF") `
    -NoGui
```

Nur bestimmte Sektionen dokumentieren:

```powershell
.\HyperV_Documentation.ps1 -HyperVServers @("HV01") `
    -Sections @("Hardware","VMOverview","VMSwitch","Cluster") -NoGui
```

## Parameter

| Parameter | Typ | Standard | Beschreibung |
| --- | --- | --- | --- |
| `-HyperVServers` | `string[]` | – | Hyper-V Hostnamen, z. B. `@("HV01","HV02")` |
| `-OutputPath` | `string` | `C:\HyperVDoku` | Ausgabeverzeichnis für Dokument und Log |
| `-CompanyName` | `string` | `Meine Organisation` | Firmenname im Dokumentkopf |
| `-OutputFormats` | `string[]` | `@("HTML")` | `HTML`, `PDF`, `Markdown` |
| `-Sections` | `string[]` | alle | Schlüssel der zu erstellenden Sektionen |
| `-ShowGui` | `switch` | – | GUI erzwingen |
| `-NoGui` | `switch` | – | GUI unterdrücken |

## Verfügbare Sektionen

**Hardware & OS**
`Hardware`, `CpuVirtualization`, `Numa`, `WindowsFeatures`, `Software`, `PowerPlan`,
`PendingReboot`, `Patch`, `DiskSpace`, `EventLog`, `TimeSync`, `Licensing`

**Hyper-V Host**
`VMHost`, `HostPaths`, `LiveMigration`, `EnhancedSession`, `HyperVServices`,
`VMVersions`, `HostUsage`

**Virtuelle Maschinen**
`VMOverview`, `VMProcessor`, `VMMemory`, `VMStorage`, `VMIntegration`, `VMCheckpoints`,
`VMFirmware`, `VMAutomatic`, `VMController`, `VMGroups`, `VMGuestOS`, `VMMetering`

**Virtuelle Netzwerke**
`VMSwitch`, `VMSwitchExt`, `VMNetwork`, `VMNetworkAdvanced`, `HostNIC`, `NetQoS`

**Speicher**
`Storage`, `VHDDetails`, `StorageQoS`, `SMBStorage`, `MPIO`

**Failover Cluster**
`Cluster`, `ClusterNetwork`, `CSV`, `ClusterQuorum`

**Replikation & Backup**
`Replication`, `Backup`

**Sicherheit**
`VMSecurity`, `HyperVAdmins`, `CredentialGuard`, `HGS`, `Delegation`,
`AntivirusExclusions`, `Firewall`, `SMBv1`

**Active Directory**
`AD`, `FSMO`

## Ausgabedateien

Im Ausgabeverzeichnis werden mit Zeitstempel abgelegt:

- `HyperV_Dokumentation_<yyyyMMdd_HHmmss>.html`
- `HyperV_Dokumentation_<yyyyMMdd_HHmmss>.pdf` (bei `-OutputFormats PDF`)
- `HyperV_Dokumentation_<yyyyMMdd_HHmmss>.md` (bei `-OutputFormats Markdown`)
- `HyperV-Dokumentation_<yyyyMMdd_HHmmss>.log`

## Changelog

Siehe [CHANGELOG.md](CHANGELOG.md).

## Lizenz

MIT – siehe [LICENSE](LICENSE).

## Autor

Rocco Ammon
