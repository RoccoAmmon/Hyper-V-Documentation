# 🖥️ Hyper-V Dokumentations-Tool / Hyper-V Documentation Tool

> **DE:** Automatische Erfassung und Dokumentation von Microsoft Hyper-V Umgebungen
> (Windows Server 2016, 2019, 2022 & 2025) – Standalone-Hosts und Failover-Cluster.
>
> **EN:** Automated inventory and documentation of Microsoft Hyper-V environments
> (Windows Server 2016, 2019, 2022 & 2025) – standalone hosts and failover clusters.

[![Status](https://img.shields.io/badge/Status-Stable-brightgreen)](CHANGELOG.md)
[![Version](https://img.shields.io/badge/Version-1.0-blue)](CHANGELOG.md)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?logo=powershell&logoColor=white)](https://www.microsoft.com/en-us/download/details.aspx?id=50395)
[![License](https://img.shields.io/github/license/RoccoAmmon/Hyper-V-Documentation)](LICENSE)
[![GitHub Issues](https://img.shields.io/github/issues/RoccoAmmon/Hyper-V-Documentation)](https://github.com/RoccoAmmon/Hyper-V-Documentation/issues)
[![Downloads](https://img.shields.io/github/downloads/RoccoAmmon/Hyper-V-Documentation/total)](https://github.com/RoccoAmmon/Hyper-V-Documentation/releases)

---

## 🎯 Übersicht / Overview

**DE:** Das Hyper-V Dokumentations-Tool ist ein umfassendes PowerShell-Skript, das
automatisch eine detaillierte Dokumentation von Microsoft Hyper-V Umgebungen erstellt.
Es erfasst Hardware, Host-Konfiguration, virtuelle Maschinen, virtuelle Netzwerke,
Speicher, Failover-Cluster, Replikation und Sicherheitsaspekte – ideal für
Administratoren, Consultants, Auditoren und Compliance-Teams.

Das Tool unterstützt Standalone-Hosts sowie Hyper-V Failover-Cluster und bietet
intelligente Fallback-Mechanismen für robuste Netzwerk-Kommunikation. Ein Cluster wird
automatisch erkannt und der Dokumenttitel entsprechend angepasst.

**EN:** The Hyper-V Documentation Tool is a comprehensive PowerShell script that
automatically generates detailed documentation of Microsoft Hyper-V environments.
It captures hardware, host configuration, virtual machines, virtual networks, storage,
failover clusters, replication and security aspects – ideal for administrators,
consultants, auditors and compliance teams.

The tool supports standalone hosts as well as Hyper-V failover clusters and provides
intelligent fallback mechanisms for robust network communication. Clusters are detected
automatically and the document title is adjusted accordingly.

---

## ✨ Features / Funktionen

### 🖥️ Hardware & Systeminfos / System Information

- CPU, RAM, Festplatten, Betriebssystem / CPU, RAM, disks, operating system
- Virtualisierungs-Unterstützung: SLAT, VT-x/AMD-V, DEP, Hypervisor-Status
- NUMA-Topologie des Hosts und NUMA-Spanning / NUMA topology and spanning
- Installierte Software, Windows Features & Rollen, Patch-Stand
- Power Plan, ausstehende Neustarts, Zeitsynchronisierung (w32tm), Lizenzierung
- Event Logs der letzten 7 Tage / event logs of the last 7 days

### 📡 Hyper-V Host-Konfiguration / Host Configuration

- Alle `Get-VMHost` Parameter / all `Get-VMHost` parameters
- Standardpfade für VHDs und VM-Konfigurationen / default VHD and VM paths
- Live Migration & Storage Migration Konfiguration
- Enhanced Session Mode und Resource Metering
- Hyper-V Dienststatus und unterstützte VM-Konfigurationsversionen
- Ressourcenauslastung & Überbuchung (CPU/RAM) / resource usage & overcommitment

### 💻 Virtuelle Maschinen / Virtual Machines

- VM-Übersicht (State, Generation, Version, Uptime)
- Prozessor-Konfiguration (vCPU, Reserve, Limit, Gewichtung, NUMA, HwThreads)
- Arbeitsspeicher & Dynamic Memory (Startup/Min/Max, Buffer, Priorität)
- Speicher (VHD/VHDX, Typ, Größe, Fragmentierung, Differencing, QoS)
- Integrationsdienste, Prüfpunkte, Firmware & Secure Boot
- Automatische Start-/Stop-Aktionen, Controller, DVD & COM-Ports
- VM-Gruppen, Gast-Betriebssysteme (KVP), Ressourcenmessung

### 🌐 Virtuelle Netzwerke / Virtual Networking

- Virtuelle Switches (Typ, Teaming, SR-IOV, Bandbreitenreservierung)
- Virtual Switch Extensions
- VM-Netzwerkadapter (Switch, MAC, IP, VLAN, Bandbreite)
- VLAN, Portsicherheit & Port-ACLs / VLAN, port security & ACLs
- Host-NICs (VMQ, RSS, SR-IOV, RDMA, Jumbo Frames, Teaming)
- Netzwerk-QoS / DCB

### 💾 Speicher / Storage

- Storage-Konfiguration (Disks, Volumes, Storage Spaces)
- VHD-Detailanalyse inkl. verwaister Dateien / VHD analysis incl. orphaned files
- Storage QoS Policies
- SMB 3.0 Storage (Shares, Multichannel, Verschlüsselung)
- MPIO (Multipath I/O)

### 🔗 Failover Cluster & Replikation / Cluster & Replication

- Cluster-Nodes, Rollen und Cluster-Eigenschaften
- Cluster-Netzwerke und Live-Migration-Netzwerke
- Cluster Shared Volumes (CSV)
- Cluster-Quorum
- Hyper-V Replica (Server-Konfiguration, VM-Replikationen, Autorisierung)
- Backup-Konfiguration & VSS Writer

### 🔒 Sicherheit & Compliance / Security & Compliance

- VM Sicherheit (vTPM, Shielded VM, Verschlüsselung)
- Hyper-V Administratoren / Hyper-V administrators
- Credential Guard und Host Guardian Service (HGS)
- Kerberos-Delegierung für Live Migration
- Antivirus-Ausschlüsse nach Best Practice / antivirus exclusions
- Firewall-Konfiguration und SMBv1-Status
- Active Directory & FSMO-Rollen

### 🖱️ GUI

- WPF-Oberfläche zur Auswahl von Hosts, Firmenname und Ausgabepfad
- Kategorisierte Auswahl aller 58 Dokumentationsbereiche
- Auswahl der Ausgabeformate direkt in der Oberfläche
- Automatische Erstellung des Ausgabeverzeichnisses

### 📄 Export & Output / Ausgabe

- HTML-Export mit formatiertem Inhaltsverzeichnis / HTML export with TOC
- PDF-Export (Microsoft Word COM, Fallback Edge/Chrome Headless)
- Markdown-Export / Markdown export
- Word-kompatible HTML-Struktur / Word-compatible HTML structure
- Detaillierte Protokollierung mit Fehler- und Warnungszähler

---

## 🚀 Schnellstart / Quick Start

### Voraussetzungen / Requirements

| Komponente / Component | Anforderung / Requirement |
| --- | --- |
| PowerShell | 5.1+ |
| Hyper-V Tools | Hyper-V PowerShell-Modul (`RSAT-Hyper-V-Tools`) |
| Cluster Modul / module | `FailoverClusters` (optional, für Cluster-Sektionen) |
| AD Modul / module | `ActiveDirectory` (optional, für AD- & FSMO-Sektionen) |
| Netzwerk / Network | CIM/RPC (DCOM) oder WinRM zu den Zielhosts |
| Rechte / Permissions | Administrative Rechte auf den Hyper-V-Hosts |
| PDF-Export (optional) | Microsoft Word **oder** Edge/Chrome |

### Installation

```powershell
# 1. Repository klonen / clone repository
git clone https://github.com/RoccoAmmon/Hyper-V-Documentation.git
cd Hyper-V-Documentation

# 2. Execution policy anpassen (falls nötig) / adjust if needed
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser

# 3. Skript ausführen / run script
.\HyperV_Documentation.ps1
```

> 💡 Das Skript startet sich bei Bedarf automatisch mit erhöhten Rechten neu. /
> The script automatically restarts itself with elevated privileges if required.

---

## 📥 Downloads

### Direkte Download-Optionen / Direct Download Options

| Methode / Method | Ziel / Target | Beschreibung / Description |
| --- | --- | --- |
| 🔗 Raw-Download (Main) | [HyperV_Documentation.ps1](https://raw.githubusercontent.com/RoccoAmmon/Hyper-V-Documentation/main/HyperV_Documentation.ps1) | Aktuelle Version direkt aus dem Repository |
| 📦 Git Clone | `git clone https://github.com/RoccoAmmon/Hyper-V-Documentation.git` | Vollständiges Repository mit Versionskontrolle |
| 📋 Releases | [All Releases](https://github.com/RoccoAmmon/Hyper-V-Documentation/releases) | Veröffentlichte Versionen |
| ⬇️ ZIP Download | [Download ZIP](https://github.com/RoccoAmmon/Hyper-V-Documentation/archive/refs/heads/main.zip) | Repository als ZIP-Archiv |

> 💡 **Empfehlung / Recommendation:** `git clone` für regelmäßige Updates verwenden!

### Schneller Download der PS1-Datei / Quick PS1 Download

```powershell
# Direkt herunterladen und speichern / download and save directly
$ScriptUrl = "https://raw.githubusercontent.com/RoccoAmmon/Hyper-V-Documentation/main/HyperV_Documentation.ps1"
$OutFile   = "C:\Scripts\HyperV_Documentation.ps1"
Invoke-WebRequest -Uri $ScriptUrl -OutFile $OutFile
Write-Host "Downloaded: $OutFile"
```

---

## 📖 Verwendungsbeispiele / Usage Examples

### Einfaches Beispiel (GUI-Auswahl) / Basic Example (GUI Selection)

```powershell
# Startet das Skript mit grafischer Oberfläche zur Host-Auswahl
# Starts the script with GUI for host selection
.\HyperV_Documentation.ps1
```

Ohne Angabe von `-HyperVServers` startet das Skript automatisch die GUI.
Alternativ lässt sie sich mit `-ShowGui` erzwingen.

### Mehrere Hosts dokumentieren / Document Multiple Hosts

```powershell
.\HyperV_Documentation.ps1 `
  -HyperVServers @('HV01','HV02','HV03') `
  -CompanyName 'Contoso GmbH' `
  -OutputPath 'D:\HyperVInventory'
```

### Ohne GUI, mit allen Formaten / Without GUI, All Formats

```powershell
.\HyperV_Documentation.ps1 `
  -HyperVServers @('HV01') `
  -CompanyName 'My Company' `
  -OutputPath 'C:\Reports' `
  -OutputFormats @('HTML','PDF','Markdown') `
  -NoGui
```

### Nur spezifische Dokumentationsbereiche / Specific Documentation Sections

```powershell
$sections = @(
    'Hardware',
    'VMHost',
    'VMOverview',
    'VMSwitch',
    'Cluster'
)

.\HyperV_Documentation.ps1 `
  -HyperVServers @('HV01') `
  -Sections $sections `
  -OutputPath 'C:\Reports' `
  -NoGui
```

---

## ⚙️ Parameter-Referenz / Parameter Reference

| Parameter | Typ / Type | Standard / Default | Beschreibung / Description |
| --- | --- | --- | --- |
| `HyperVServers` | `string[]` | (GUI) | Zu dokumentierende Hyper-V Hosts / Hyper-V hosts to document |
| `CompanyName` | `string` | `Meine Organisation` | Firmenname im Report / company name in report |
| `OutputPath` | `string` | `C:\HyperVDoku` | Ausgabeverzeichnis / output directory |
| `OutputFormats` | `string[]` | `@('HTML')` | Ausgabeformate: `HTML`, `PDF`, `Markdown` |
| `Sections` | `string[]` | (alle / all) | Zu erstellende Dokumentationsbereiche |
| `ShowGui` | `switch` | - | GUI erzwingen / force GUI |
| `NoGui` | `switch` | - | GUI unterdrücken / suppress GUI |

---

## 📋 Dokumentierte Inhalte / Documented Contents

**DE:** Insgesamt stehen **58 Dokumentationsbereiche in 9 Kategorien** zur Verfügung.

**EN:** A total of **58 documentation sections in 9 categories** are available.

| Kategorie / Category | Sektions-Schlüssel / Section Keys |
| --- | --- |
| 🖥️ Hardware & OS | `Hardware`, `CpuVirtualization`, `Numa`, `WindowsFeatures`, `Software`, `PowerPlan`, `PendingReboot`, `Patch`, `DiskSpace`, `EventLog`, `TimeSync`, `Licensing` |
| 📡 Hyper-V Host | `VMHost`, `HostPaths`, `LiveMigration`, `EnhancedSession`, `HyperVServices`, `VMVersions`, `HostUsage` |
| 💻 Virtuelle Maschinen | `VMOverview`, `VMProcessor`, `VMMemory`, `VMStorage`, `VMIntegration`, `VMCheckpoints`, `VMFirmware`, `VMAutomatic`, `VMController`, `VMGroups`, `VMGuestOS`, `VMMetering` |
| 🌐 Virtuelle Netzwerke | `VMSwitch`, `VMSwitchExt`, `VMNetwork`, `VMNetworkAdvanced`, `HostNIC`, `NetQoS` |
| 💾 Speicher | `Storage`, `VHDDetails`, `StorageQoS`, `SMBStorage`, `MPIO` |
| 🔗 Failover Cluster | `Cluster`, `ClusterNetwork`, `CSV`, `ClusterQuorum` |
| ♻️ Replikation & Backup | `Replication`, `Backup` |
| 🔒 Sicherheit | `VMSecurity`, `HyperVAdmins`, `CredentialGuard`, `HGS`, `Delegation`, `AntivirusExclusions`, `Firewall`, `SMBv1` |
| 🏢 Active Directory | `AD`, `FSMO` |

---

## 🌐 Netzwerkverbindung / Network Connection

**DE:** Das Tool verwendet intelligente Verbindungsmechanismen:

**EN:** The tool uses intelligent connection mechanisms:

1. **WsMan (WinRM)** – bevorzugter Standard / preferred standard
2. **DCOM RPC Fallback** – funktioniert ohne WinRM / works without WinRM
3. **Remote Registry** – .NET-basiert, kein WinRM nötig / .NET-based, no WinRM required
4. **Invoke-Command** – für lokal verfügbare Cmdlets / for locally available cmdlets

> ✅ **Vorteil / Advantage:** Funktioniert auch in Umgebungen mit deaktiviertem WinRM!

---

## 📝 Output-Beispiel

Die erzeugten Dateien werden mit Zeitstempel im Ausgabeverzeichnis abgelegt:

- `HyperV_Dokumentation_<yyyyMMdd_HHmmss>.html`
- `HyperV_Dokumentation_<yyyyMMdd_HHmmss>.pdf` (bei `-OutputFormats PDF`)
- `HyperV_Dokumentation_<yyyyMMdd_HHmmss>.md` (bei `-OutputFormats Markdown`)
- `HyperV-Dokumentation_<yyyyMMdd_HHmmss>.log`

Das generierte HTML-Dokument hat diese Struktur:

```
Hyper-V Umgebungsdokumentation
├─ Inhaltsverzeichnis
├─ Zusammenfassung
│  ├─ Dokumentierte Hosts
│  ├─ Erstellungsdatum
│  └─ Fehler- und Warnung-Zähler
│
├─ Hardware & OS
│  ├─ Hardware & Host-Details
│  ├─ Virtualisierungs-Unterstützung (SLAT)
│  ├─ NUMA-Topologie
│  └─ Patch-Stand & Event Logs
│
├─ Hyper-V Host
│  ├─ Host-Konfiguration
│  ├─ Standardpfade
│  ├─ Live Migration
│  └─ Ressourcenauslastung
│
├─ Virtuelle Maschinen
│  ├─ VM-Übersicht
│  ├─ Prozessor & Arbeitsspeicher
│  ├─ Speicher (VHD/VHDX)
│  └─ Prüfpunkte & Firmware
│
└─ [weitere Sektionen...]
```

---

## 🐛 Troubleshooting

### Problem: "Hyper-V PowerShell-Modul nicht verfügbar"

```powershell
# Hyper-V Verwaltungstools installieren / install Hyper-V management tools
Install-WindowsFeature -Name RSAT-Hyper-V-Tools -IncludeAllSubFeature

# Auf Windows-Clients alternativ / alternatively on Windows clients
Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-Management-PowerShell
```

### Problem: "Keine Verbindung zu Host XY"

```powershell
# Netzwerk-Erreichbarkeit prüfen / check network reachability
Test-Connection -ComputerName 'HV01' -Count 2

# DCOM/RPC aktivieren / enable DCOM/RPC
Enable-NetFirewallRule -DisplayGroup "Windows Management Instrumentation (WMI)"
```

### Problem: "Zugriff verweigert"

```powershell
# Mit Admin-Rechten ausführen / run with admin privileges
Start-Process powershell -ArgumentList "& '.\HyperV_Documentation.ps1'" -Verb RunAs
```

### Problem: "Cluster-Sektionen bleiben leer"

```powershell
# FailoverClusters-Modul installieren / install FailoverClusters module
Install-WindowsFeature -Name RSAT-Clustering-PowerShell
```

### Problem: "PDF-Export nicht möglich"

```powershell
# PDF benötigt Word oder Edge/Chrome - alternativ HTML/Markdown erzeugen
# PDF requires Word or Edge/Chrome - alternatively create HTML/Markdown
.\HyperV_Documentation.ps1 -HyperVServers @('HV01') -OutputFormats @('HTML','Markdown') -NoGui
```

---

## 📓 Changelog

Alle Änderungen sind in [CHANGELOG.md](CHANGELOG.md) dokumentiert. /
All changes are documented in [CHANGELOG.md](CHANGELOG.md).

---

## 🤝 Mitarbeit / Contributing

Beiträge sind willkommen! Bitte:

1. Fork das Repository
2. Feature Branch erstellen (`git checkout -b feature/MyFeature`)
3. Änderungen committen (`git commit -am 'Add MyFeature'`)
4. Branch pushen (`git push origin feature/MyFeature`)
5. Pull Request öffnen

---

## ❓ FAQ

**F: Funktioniert das Skript auch mit non-Administrator-Konten?**
A: Nein, administrative Rechte sind erforderlich. Das Skript startet sich bei Bedarf
automatisch erhöht neu.

**F: Kann ich das Dokument in Word bearbeiten?**
A: Ja, die HTML-Datei kann direkt in Word geöffnet werden.

**F: Muss das Skript auf einem Hyper-V-Host laufen?**
A: Nein, es kann auch von einer Management-Station mit RSAT-Hyper-V-Tools ausgeführt werden.

**F: Werden Failover-Cluster unterstützt?**
A: Ja, Cluster werden automatisch erkannt und mit Nodes, Netzwerken, CSV und Quorum
dokumentiert.

**F: Welche Daten werden übertragen?**
A: Nur Lesevorgänge auf die Hyper-V-Hosts. Keine Daten verlassen die Umgebung.

**F: Wie lange dauert die Dokumentation?**
A: Je nach Anzahl Hosts, VMs und gewählten Sektionen typischerweise 5-30 Minuten.

---

## 📄 Lizenz

Dieses Projekt ist unter der MIT-Lizenz lizenziert – siehe [LICENSE](LICENSE) für Details.

---

## 👤 Autor

**Rocco Ammon**

- 🔗 GitHub: [@RoccoAmmon](https://github.com/RoccoAmmon)
- 📧 Kontakt via GitHub Issues
- 🌍 Repository: [Hyper-V-Documentation](https://github.com/RoccoAmmon/Hyper-V-Documentation)

---

## 📞 Support

- 📋 **Issues:** [GitHub Issues](https://github.com/RoccoAmmon/Hyper-V-Documentation/issues)
- 💬 **Discussions:** [GitHub Discussions](https://github.com/RoccoAmmon/Hyper-V-Documentation/discussions)

> **Hinweis:** Vor dem Einsatz in produktiven Umgebungen die Skripte in einer
> Testumgebung prüfen.

