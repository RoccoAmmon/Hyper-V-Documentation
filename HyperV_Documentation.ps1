<#
.SYNOPSIS
    Hyper-V Dokumentations-Skript (v1.1 - Windows Server 2016/2019/2022/2025)
.DESCRIPTION
    Erstellt eine umfassende HTML-Dokumentation der gesamten Hyper-V Umgebung.
    Unterstützt Standalone-Hosts und Failover-Cluster (Hyper-V Cluster).
    Das HTML-Dokument kann direkt in Microsoft Word importiert werden.

    WICHTIG: Diese Version verwendet CIM-Sessions mit automatischem DCOM-Fallback,
    sodass das Skript auch funktioniert, wenn WinRM (PowerShell Remoting) nicht
    korrekt konfiguriert ist. Für einige Hyper-V-Detailabfragen wird WinRM
    bevorzugt, es existiert jedoch immer ein Fallback über CIM/WMI.

    Dokumentiert werden u.a.:
    - Hardware-Informationen (OS, RAM, CPU, Festplatten, Netzwerk)
    - Virtualisierungs-Unterstützung (SLAT, VT-x/AMD-V, DEP, Hypervisor-Status)
    - NUMA-Topologie des Hosts und NUMA-Spanning
    - Installierte Software, Windows Features & Rollen, Patches
    - Power Plan, ausstehende Neustarts, Zeitsynchronisierung
    - Hyper-V Host-Konfiguration (alle Get-VMHost Parameter)
    - Standardpfade für VHDs und VM-Konfigurationen
    - Live Migration & Storage Migration Konfiguration
    - Enhanced Session Mode, Resource Metering
    - Hyper-V Dienste & Ereignisprotokolle
    - Unterstützte VM-Konfigurationsversionen
    - Virtuelle Maschinen Übersicht (State, Generation, Version, Uptime)
    - VM Prozessor-Konfiguration (vCPU, Reserve, Limit, Gewichtung, NUMA, HwThreads)
    - VM Arbeitsspeicher (Dynamic Memory, Startup/Min/Max, Buffer, Priorität)
    - VM Speicher (VHD/VHDX, Typ, Größe, Fragmentierung, Differencing, QoS)
    - VM Netzwerkadapter (Switch, MAC, IP, VLAN, Bandbreite, Security-Settings)
    - Integrationsdienste je VM
    - Prüfpunkte / Checkpoints (Typ, Alter, Kette)
    - VM Firmware & Secure Boot (Generation 2)
    - VM Security (vTPM, Shielded VM, Verschlüsselung)
    - Automatische Start-/Stop-Aktionen
    - DVD-Laufwerke, Controller, COM-Ports
    - VM Gruppen und Gast-Betriebssysteme (KVP)
    - Ressourcenmessung (Resource Metering)
    - Virtuelle Switches (Typ, Teaming, SR-IOV, Bandbreitenreservierung)
    - Switch Extensions
    - Host-Netzwerkadapter (VMQ, RSS, SR-IOV, Jumbo Frames, Teaming)
    - Netzwerk-QoS / DCB / RDMA
    - Live Migration Netzwerke
    - Storage-Konfiguration (Disks, Volumes, Storage Spaces, MPIO)
    - VHD-Detailanalyse
    - Storage QoS Policies
    - SMB 3.0 Storage (Shares, Multichannel, Verschlüsselung)
    - Failover Cluster (Nodes, Netzwerke, Quorum, CSV, Rollen)
    - Hyper-V Replica (Server-Konfiguration, VM-Replikationen, Autorisierung)
    - VSS Writer & Backup-Konfiguration
    - Hyper-V Administratoren, Credential Guard, HGS
    - Kerberos Delegation für Live Migration
    - Antivirus-Ausschlüsse (Best Practice für Hyper-V)
    - Firewall-Regeln und SMBv1-Status
    - Active Directory & FSMO-Rollen

.PARAMETER HyperVServers
    Array von Hyper-V Hostnamen, die dokumentiert werden sollen.
    Beispiel: @("HV01","HV02")

.PARAMETER OutputPath
    Pfad für die Ausgabe-Datei. Standard: C:\HyperVDoku

.PARAMETER CompanyName
    Name der Firma für die Dokumentation.

.PARAMETER OutputFormats
    Ausgabeformate: HTML, PDF, Markdown

.PARAMETER Sections
    Liste der zu erstellenden Sektions-Schlüssel. Leer = alle.

.PARAMETER Language
    Sprache für GUI, Konsolenausgaben und Report: DE (Deutsch) oder EN (Englisch).
    Standard: DE

.EXAMPLE
    .\HyperV_Documentation.ps1 -HyperVServers @("HV01","HV02") -CompanyName "Contoso GmbH"

.EXAMPLE
    .\HyperV_Documentation.ps1 -HyperVServers @("HV01") -CompanyName "Meine Firma" -OutputPath "D:\Doku"

.EXAMPLE
    .\HyperV_Documentation.ps1 -HyperVServers @("HV01") -Language EN -NoGui

.NOTES
    Autor:           Rocco Ammon
    Version:         1.1
    Erstellt:        2026-08-31
    Letzte Änderung: 2026-08-31
    Änderungen:      v1.1 - Zweisprachigkeit (DE/EN) über -Language inkl. GUI-Umschalter
                     v1.0 - Erstveröffentlichung (analog zur Exchange-Dokumentation)
    Voraussetzungen: - Hyper-V PowerShell-Modul (RSAT-Hyper-V-Tools)
                     - Optional: FailoverClusters-Modul für Cluster-Dokumentation
                     - Optional: Active Directory PowerShell-Modul
                     - Administratorrechte auf den Hyper-V-Hosts
                     - RPC/DCOM oder WinRM muss erreichbar sein
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Geben Sie die Hyper-V Hostnamen als Array an, z.B. @('HV01','HV02')")]
    [string[]]$HyperVServers,

    [Parameter(Mandatory = $false, HelpMessage = "Ausgabepfad für die Dokumentation")]
    [string]$OutputPath = "C:\HyperVDoku",

    [Parameter(Mandatory = $false, HelpMessage = "Firmenname für die Dokumentation")]
    [string]$CompanyName = "Meine Organisation",

    [Parameter(Mandatory = $false, HelpMessage = "Ausgabeformate: HTML, PDF, Markdown")]
    [ValidateSet("HTML", "PDF", "Markdown")]
    [string[]]$OutputFormats = @("HTML"),

    [Parameter(Mandatory = $false, HelpMessage = "Liste der zu erstellenden Sektions-Schlüssel. Leer = alle")]
    [string[]]$Sections,

    [Parameter(Mandatory = $false, HelpMessage = "Sprache für GUI, Konsole und Report: DE oder EN")]
    [ValidateSet("DE", "EN")]
    [string]$Language = "DE",

    [Parameter(Mandatory = $false, HelpMessage = "GUI zur Auswahl anzeigen (Standard, wenn keine Server angegeben)")]
    [switch]$ShowGui,

    [Parameter(Mandatory = $false, HelpMessage = "GUI unterdrücken und direkt mit Parametern starten")]
    [switch]$NoGui
)

#region ============================================================
# ADMINISTRATOR-PRÜFUNG & AUTO-RESTART
#endregion ============================================================

# Prüfe, ob das Skript als Administrator läuft
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)

if (-not $isAdmin) {
    Write-Host $(if ($Language -eq "EN") { "The script requires administrator privileges. Restarting elevated..." } else { "Das Skript erfordert Administrator-Rechte. Starte neu mit erhöhten Rechten..." }) -ForegroundColor Yellow

    # Sammle alle übergebenen Parameter
    $argumentList = @()

    if ($HyperVServers) {
        $argumentList += "-HyperVServers @($(($HyperVServers | ForEach-Object { "'{0}'" -f $_ }) -join ','))"
    }
    if ($OutputPath -ne "C:\HyperVDoku") {
        $argumentList += "-OutputPath '$OutputPath'"
    }
    if ($CompanyName -ne "Meine Organisation") {
        $argumentList += "-CompanyName '$CompanyName'"
    }
    if ($OutputFormats.Count -gt 0) {
        $argumentList += "-OutputFormats @($(($OutputFormats | ForEach-Object { "'{0}'" -f $_ }) -join ','))"
    }
    if ($Sections) {
        $argumentList += "-Sections @($(($Sections | ForEach-Object { "'{0}'" -f $_ }) -join ','))"
    }
    if ($Language) {
        $argumentList += "-Language '$Language'"
    }
    if ($ShowGui) {
        $argumentList += "-ShowGui"
    }
    if ($NoGui) {
        $argumentList += "-NoGui"
    }

    # Starte das Skript mit Administrator-Rechten
    $scriptPath = $MyInvocation.MyCommand.Path
    $command = "& '$scriptPath' $($argumentList -join ' ')"

    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"$command`"" -Verb RunAs -Wait
    exit
}

Write-Host $(if ($Language -eq "EN") { "Administrator privileges confirmed. Running script..." } else { "Administrator-Rechte bestätigt. Skript wird ausgeführt..." }) -ForegroundColor Green

#region ============================================================
# VARIABLEN-DEFINITION
#endregion ============================================================

# --- Pfade und Dateien ---
$script:LogPath                 = $OutputPath
$script:LogFile                 = Join-Path -Path $LogPath -ChildPath "HyperV-Dokumentation_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$script:HTMLOutputFile          = Join-Path -Path $LogPath -ChildPath "HyperV_Dokumentation_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$script:Timestamp               = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
$script:DateOnly                = Get-Date -Format "dd.MM.yyyy"

# --- Dokumentations-Variablen ---
$script:DocTitle                = "Hyper-V - Umgebungsdokumentation"
$script:DocSubTitle             = "$CompanyName"
$script:DocAuthor               = $env:USERNAME
$script:DocComputerName         = $env:COMPUTERNAME
$script:ScriptVersion           = "1.1"
$script:Language                = $Language

# --- Ausgabeverzeichnis erstellen ---
try {
    if (-not (Test-Path $script:LogPath)) {
        New-Item -Path $script:LogPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
}
catch {
    Write-Host "FEHLER: Ausgabeverzeichnis konnte nicht erstellt werden: $_" -ForegroundColor Red
}

# --- Sammelvariablen für HTML-Sektionen ---
$script:HTMLSections            = [System.Collections.ArrayList]::new()
$script:TOCEntries              = [System.Collections.ArrayList]::new()
$script:SectionCounter          = 0
$script:ErrorCount              = 0
$script:WarningCount            = 0

# --- Hyper-V Umgebungsvariablen (zur Laufzeit ermittelt) ---
$script:HyperVModuleLoaded      = $false
$script:ClusterModuleLoaded     = $false
$script:HyperVEdition           = "Unknown"    # z. B. "Windows Server 2022"
$script:ClusterName             = ""           # Name des Failover-Clusters (falls vorhanden)

# --- Grenzwerte für Warnungen ---
$script:WarningDiskSpaceGB      = 50      # Warnung bei weniger als 50 GB freiem Speicher
$script:WarningCheckpointDays   = 7       # Warnung bei Checkpoints älter als 7 Tage
$script:WarningVHDFragPercent   = 30      # Warnung bei VHDX-Fragmentierung über 30 %
$script:MaxVMsForDetailStats    = 500     # Maximale Anzahl VMs für Detailstatistiken
$script:MinFreeHostMemoryGB     = 4       # Reserve für das Host-Betriebssystem

#endregion

#region ============================================================
# LOKALISIERUNG (DE / EN)
#endregion ============================================================

# Wörterbuch Deutsch -> Englisch. Schlüssel sind exakt die im Skript verwendeten
# deutschen Texte (Überschriften, Spaltennamen, Statuswerte, Meldungen).
$script:Dict = @{
    # --- Kategorien & Sektions-Labels ---
    "Hardware & OS"                                = "Hardware & OS"
    "Hyper-V Host"                                 = "Hyper-V Host"
    "Virtuelle Maschinen"                          = "Virtual Machines"
    "Virtuelle Netzwerke"                          = "Virtual Networking"
    "Speicher"                                     = "Storage"
    "Failover Cluster"                             = "Failover Cluster"
    "Replikation & Backup"                         = "Replication & Backup"
    "Sicherheit"                                   = "Security"
    "Active Directory"                             = "Active Directory"

    "Hardware & Host-Details"                      = "Hardware & host details"
    "Virtualisierungs-Unterstützung (SLAT)"        = "Virtualization support (SLAT)"
    "NUMA-Topologie & Spanning"                    = "NUMA topology & spanning"
    "Windows Features & Rollen"                    = "Windows features & roles"
    "Installierte Software"                        = "Installed software"
    "Power Plan & Performance"                     = "Power plan & performance"
    "Ausstehende Neustarts"                        = "Pending reboots"
    "Windows Updates & Patch-Stand"                = "Windows updates & patch level"
    "Speicherplatz & VM-Pfade"                     = "Disk space & VM paths"
    "Event Logs (7 Tage)"                          = "Event logs (7 days)"
    "Zeitsynchronisierung (w32tm)"                 = "Time synchronization (w32tm)"
    "Lizenzierung & Aktivierung"                   = "Licensing & activation"
    "Hyper-V Host-Konfiguration"                   = "Hyper-V host configuration"
    "Standardpfade (VHD & VM)"                     = "Default paths (VHD & VM)"
    "Live Migration & Storage Migration"           = "Live migration & storage migration"
    "Enhanced Session & Metering"                  = "Enhanced session & metering"
    "Hyper-V Dienststatus"                         = "Hyper-V service status"
    "VM-Konfigurationsversionen"                   = "VM configuration versions"
    "Ressourcenauslastung & Überbuchung"           = "Resource usage & overcommitment"
    "VM-Übersicht"                                 = "VM overview"
    "VM Prozessor-Konfiguration"                   = "VM processor configuration"
    "VM Arbeitsspeicher & Dynamic Memory"          = "VM memory & dynamic memory"
    "VM Speicher (VHD/VHDX)"                       = "VM storage (VHD/VHDX)"
    "Integrationsdienste"                          = "Integration services"
    "Prüfpunkte (Checkpoints)"                     = "Checkpoints"
    "Firmware / BIOS & Secure Boot"                = "Firmware / BIOS & secure boot"
    "Automatische Start-/Stop-Aktionen"            = "Automatic start/stop actions"
    "Controller, DVD & COM-Ports"                  = "Controllers, DVD & COM ports"
    "VM-Gruppen"                                   = "VM groups"
    "Gast-Betriebssysteme (KVP)"                   = "Guest operating systems (KVP)"
    "Ressourcenmessung (Metering)"                 = "Resource metering"
    "Virtuelle Switches"                           = "Virtual switches"
    "Virtual Switch Extensions"                    = "Virtual switch extensions"
    "VM Netzwerkadapter"                           = "VM network adapters"
    "VLAN, Portsicherheit & ACLs"                  = "VLAN, port security & ACLs"
    "Host-NICs (VMQ/RSS/SR-IOV/RDMA)"              = "Host NICs (VMQ/RSS/SR-IOV/RDMA)"
    "Netzwerk-QoS / DCB"                           = "Network QoS / DCB"
    "Storage-Konfiguration"                        = "Storage configuration"
    "VHD-Analyse & verwaiste Dateien"              = "VHD analysis & orphaned files"
    "Storage QoS Policies"                         = "Storage QoS policies"
    "SMB 3.0 Storage"                              = "SMB 3.0 storage"
    "MPIO (Multipath I/O)"                         = "MPIO (multipath I/O)"
    "Cluster-Netzwerke"                            = "Cluster networks"
    "Cluster Shared Volumes (CSV)"                 = "Cluster shared volumes (CSV)"
    "Cluster-Quorum"                               = "Cluster quorum"
    "Hyper-V Replica"                              = "Hyper-V Replica"
    "Backup & VSS Writer"                          = "Backup & VSS writer"
    "VM Sicherheit (vTPM / Shielded)"              = "VM security (vTPM / shielded)"
    "Hyper-V Administratoren"                      = "Hyper-V administrators"
    "Credential Guard"                             = "Credential Guard"
    "Host Guardian Service"                        = "Host Guardian Service"
    "Kerberos-Delegierung"                         = "Kerberos delegation"
    "Antivirus-Ausschlüsse"                        = "Antivirus exclusions"
    "Firewall-Konfiguration"                       = "Firewall configuration"
    "SMBv1 Status (Sicherheit)"                    = "SMBv1 status (security)"
    "FSMO-Rollen"                                  = "FSMO roles"

    # --- Sektions-Titel im Report ---
    "Hardware-Informationen & Host-Details"        = "Hardware information & host details"
    "Virtualisierungs-Unterstützung (SLAT / VT-x / VBS)" = "Virtualization support (SLAT / VT-x / VBS)"
    "Windows Features &amp; Rollen"                = "Windows features &amp; roles"
    "Power Plan &amp; Performance"                 = "Power plan &amp; performance"
    "Windows Updates &amp; Patch-Stand"            = "Windows updates &amp; patch level"
    "Speicherplatz &amp; VM-Pfade"                 = "Disk space &amp; VM paths"
    "Event Logs (Fehler/Kritisch, 7 Tage)"         = "Event logs (error/critical, 7 days)"
    "Lizenzierung &amp; Aktivierung"               = "Licensing &amp; activation"
    "Standardpfade (VHD &amp; VM-Konfiguration)"   = "Default paths (VHD &amp; VM configuration)"
    "Live Migration &amp; Storage Migration"       = "Live migration &amp; storage migration"
    "Enhanced Session Mode &amp; Resource Metering" = "Enhanced session mode &amp; resource metering"
    "Host-Ressourcenauslastung &amp; Überbuchung"  = "Host resource usage &amp; overcommitment"
    "Virtuelle Maschinen - Übersicht"              = "Virtual machines - overview"
    "VM Arbeitsspeicher &amp; Dynamic Memory"      = "VM memory &amp; dynamic memory"
    "VM Netzwerk - VLAN, Sicherheit &amp; ACLs"    = "VM networking - VLAN, security &amp; ACLs"
    "VM Firmware / BIOS &amp; Secure Boot"         = "VM firmware / BIOS &amp; secure boot"
    "VM Sicherheit (vTPM / Shielded VM)"           = "VM security (vTPM / shielded VM)"
    "VM Controller, DVD &amp; COM-Ports"           = "VM controllers, DVD &amp; COM ports"
    "Host-Netzwerkadapter (VMQ / RSS / SR-IOV / RDMA)" = "Host network adapters (VMQ / RSS / SR-IOV / RDMA)"
    "VHD-Detailanalyse &amp; verwaiste Dateien"    = "VHD detail analysis &amp; orphaned files"
    "Backup &amp; VSS Writer"                      = "Backup &amp; VSS writer"
    "Host Guardian Service (Guarded Fabric)"       = "Host Guardian Service (guarded fabric)"
    "Kerberos-Delegierung (Live Migration)"        = "Kerberos delegation (live migration)"
    "Antivirus-Ausschlüsse (Hyper-V)"              = "Antivirus exclusions (Hyper-V)"

    # --- Zwischenüberschriften (h3 / h4) ---
    "Aktivierungsstatus"                           = "Activation status"
    "Alle installierten Features"                  = "All installed features"
    "Allgemeine Host-Einstellungen"                = "General host settings"
    "Autorisierungseinträge"                       = "Authorization entries"
    "Best Practice"                                = "Best practice"
    "Betriebssystem"                               = "Operating system"
    "Cluster-Knoten"                               = "Cluster nodes"
    "Cluster-Ressourcen"                           = "Cluster resources"
    "Cluster-Rollen"                               = "Cluster roles"
    "Cluster-Übersicht"                            = "Cluster overview"
    "COM-Ports (aktiv konfiguriert)"               = "COM ports (actively configured)"
    "CSV Block Cache"                              = "CSV block cache"
    "Definierte Migrationsnetzwerke"               = "Defined migration networks"
    "Diskettenlaufwerke"                           = "Floppy drives"
    "Dokumentations-Zusammenfassung"               = "Documentation summary"
    "Domäne &amp; Gesamtstruktur"                  = "Domain &amp; forest"
    "DVD-Laufwerke"                                = "DVD drives"
    "Edition &amp; VM-Rechte"                      = "Edition &amp; VM rights"
    "Empfohlene Dateityp-Ausschlüsse"              = "Recommended file type exclusions"
    "Empfohlene Hyper-V Ausschlüsse"               = "Recommended Hyper-V exclusions"
    "Empfohlene Prozess-Ausschlüsse"               = "Recommended process exclusions"
    "Enhanced Session Mode je VM"                  = "Enhanced session mode per VM"
    "Erweiterte Port ACLs"                         = "Extended port ACLs"
    "Firewall-Profile"                             = "Firewall profiles"
    "Generation 1 - BIOS"                          = "Generation 1 - BIOS"
    "Generation 2 - UEFI Firmware"                 = "Generation 2 - UEFI firmware"
    "Hinweis"                                      = "Note"
    "Hinweis zur Virtualisierungs-Lizenzierung"    = "Note on virtualization licensing"
    "Host-Einstellungen"                           = "Host settings"
    "Host-vNICs (Management OS)"                   = "Host vNICs (management OS)"
    "Hostcomputerkonten"                           = "Host computer accounts"
    "Hyper-V Dienste"                              = "Hyper-V services"
    "Hyper-V relevante Rollen &amp; Features"      = "Hyper-V relevant roles &amp; features"
    "Installierte Antivirenprodukte"               = "Installed antivirus products"
    "iSCSI Targets"                                = "iSCSI targets"
    "Kapazität &amp; Überbuchung"                  = "Capacity &amp; overcommitment"
    "Knoten-Stimmgewichtung"                       = "Node vote weight"
    "Live Migration Netzwerkpriorität"             = "Live migration network priority"
    "Logische Laufwerke"                           = "Logical drives"
    "Microsoft Defender Status"                    = "Microsoft Defender status"
    "Migrationseinstellungen"                      = "Migration settings"
    "MPIO-Einstellungen"                           = "MPIO settings"
    "MPIO-fähige Hardware"                         = "MPIO-capable hardware"
    "Netzwerkkonfiguration (IP)"                   = "Network configuration (IP)"
    "Netzwerkschnittstellen der Knoten"            = "Node network interfaces"
    "NIC-Teaming (LBFO)"                           = "NIC teaming (LBFO)"
    "NUMA Nodes"                                   = "NUMA nodes"
    "NUMA Spanning"                                = "NUMA spanning"
    "NUMA-Konfiguration der VMs"                   = "NUMA configuration of the VMs"
    "Pagefile"                                     = "Page file"
    "Pass-Through Disks"                           = "Pass-through disks"
    "Performance Counter (Momentaufnahme)"         = "Performance counters (snapshot)"
    "Pfad-Ausschlüsse"                             = "Path exclusions"
    "Physische Datenträger (Disks)"                = "Physical disks"
    "Physische Festplatten"                        = "Physical hard disks"
    "Physische Netzwerkadapter"                    = "Physical network adapters"
    "Physischer Arbeitsspeicher (DIMMs)"           = "Physical memory (DIMMs)"
    "Port ACLs"                                    = "Port ACLs"
    "Portsicherheit &amp; erweiterte Features"     = "Port security &amp; advanced features"
    "Priority Flow Control (PFC)"                  = "Priority flow control (PFC)"
    "Prozessor(en)"                                = "Processor(s)"
    "Prüfpunkt-Einstellungen"                      = "Checkpoint settings"
    "QoS-Richtlinien"                              = "QoS policies"
    "RDMA (SMB Direct)"                            = "RDMA (SMB Direct)"
    "Relevante aktive Firewall-Regeln"             = "Relevant active firewall rules"
    "Replikations-Serverkonfiguration"             = "Replication server configuration"
    "Replizierte virtuelle Maschinen"              = "Replicated virtual machines"
    "RSS (Receive Side Scaling)"                   = "RSS (receive side scaling)"
    "SCSI-Controller"                              = "SCSI controllers"
    "SMB Multichannel Verbindungen"                = "SMB multichannel connections"
    "SMB-Client Konfiguration"                     = "SMB client configuration"
    "SMB-Freigaben"                                = "SMB shares"
    "SMB-Server Konfiguration"                     = "SMB server configuration"
    "Speicherpfade der virtuellen Maschinen"       = "Storage paths of the virtual machines"
    "Standard-Lastverteilungsrichtlinie"           = "Default load balancing policy"
    "Standardpfade"                                = "Default paths"
    "Storage Pools"                                = "Storage pools"
    "Switch Embedded Teaming (SET)"                = "Switch embedded teaming (SET)"
    "System-Übersicht"                             = "System overview"
    "Traffic Classes (DCB)"                        = "Traffic classes (DCB)"
    "Unterstützte Konfigurationsversionen"         = "Supported configuration versions"
    "Versionen &amp; Status"                       = "Versions &amp; status"
    "Virtualization Based Security (VBS)"          = "Virtualization based security (VBS)"
    "Virtuelle Datenträger (Storage Spaces)"       = "Virtual disks (storage spaces)"
    "Virtuelle Festplatten"                        = "Virtual hard disks"
    "Virtuelle Fibre Channel Adapter"              = "Virtual fibre channel adapters"
    "VLAN-Konfiguration"                           = "VLAN configuration"
    "VMQ (Virtual Machine Queue)"                  = "VMQ (virtual machine queue)"
    "VMs mit abweichenden Pfaden"                  = "VMs with deviating paths"
    "Vollständige Cluster-Parameter"               = "Complete cluster parameters"
    "Vollständige Parameterliste (Get-VMHost)"     = "Complete parameter list (Get-VMHost)"
    "Volumes"                                      = "Volumes"
    "Vorhandene Prüfpunkte"                        = "Existing checkpoints"
    "VSS Writer"                                   = "VSS writer"
    "w32tm Status"                                 = "w32tm status"
    "Windows Server Backup"                        = "Windows Server Backup"
    "Zusammenfassung"                              = "Summary"

    # --- Statuswerte & Hinweistexte ---
    "Ja"                                           = "Yes"
    "Nein"                                         = "No"
    "Aktiviert"                                    = "Enabled"
    "Deaktiviert"                                  = "Disabled"
    "Unbekannt"                                    = "Unknown"
    "Kritisch"                                     = "Critical"
    "Fehler"                                       = "Error"
    "Warnung"                                      = "Warning"
    "Läuft"                                        = "Running"
    "Standard"                                     = "Default"
    "Nicht aktiviert"                              = "Not activated"
    "Nicht ermittelbar"                            = "Cannot be determined"
    "Nicht für Cluster verwenden"                  = "Do not use for cluster"
    "Nicht gesetzt (Standard)"                     = "Not set (default)"
    "Nicht installiert"                            = "Not installed"
    "Nicht konfiguriert"                           = "Not configured"
    "Nicht lizenziert"                             = "Not licensed"
    "Nicht verfügbar"                              = "Not available"
    "Nicht verfügbar."                             = "Not available."
    "Nicht verfügbar (WinRM erforderlich)."        = "Not available (WinRM required)."
    "Kein Cluster (Standalone)"                    = "No cluster (standalone)"
    "Kein LBFO-Team konfiguriert."                 = "No LBFO team configured."
    "Kein SET-Team konfiguriert."                  = "No SET team configured."
    "Kein Zeuge konfiguriert"                      = "No witness configured"
    "Keine (nur CredSSP möglich)"                  = "None (only CredSSP possible)"
    "Keine Adapter gefunden."                      = "No adapters found."
    "Keine Autorisierungseinträge konfiguriert."   = "No authorization entries configured."
    "Keine Cluster Shared Volumes vorhanden."      = "No cluster shared volumes present."
    "Keine COM-Ports konfiguriert."                = "No COM ports configured."
    "Keine Computerkonten gefunden."               = "No computer accounts found."
    "Keine Daten verfügbar."                       = "No data available."
    "Keine Datenträger gefunden."                  = "No disks found."
    "Keine Diskettenimages eingebunden."           = "No floppy images mounted."
    "Keine DVD-Laufwerke konfiguriert."            = "No DVD drives configured."
    "Keine erweiterten ACLs konfiguriert."         = "No extended ACLs configured."
    "Keine Features gefunden."                     = "No features found."
    "Keine Freigaben vorhanden."                   = "No shares present."
    "Keine Gastinformationen verfügbar."           = "No guest information available."
    "Keine Generation-1-VMs vorhanden."            = "No generation 1 VMs present."
    "Keine Generation-2-VMs vorhanden."            = "No generation 2 VMs present."
    "Keine Host-vNICs gefunden."                   = "No host vNICs found."
    "Keine Hyper-V Features gefunden."             = "No Hyper-V features found."
    "Keine Integrationsdienste gefunden."          = "No integration services found."
    "Keine iSCSI-Targets konfiguriert."            = "No iSCSI targets configured."
    "Keine Lizenzinformationen verfügbar."         = "No licensing information available."
    "Keine MPIO-Hardware erkannt."                 = "No MPIO hardware detected."
    "Keine Multichannel-Verbindungen aktiv."       = "No multichannel connections active."
    "Keine Netzwerkadapter gefunden."              = "No network adapters found."
    "Keine NUMA-Knoten gefunden."                  = "No NUMA nodes found."
    "Keine Pass-Through Disks konfiguriert."       = "No pass-through disks configured."
    "Keine Pfad-Ausschlüsse konfiguriert."         = "No path exclusions configured."
    "Keine Port-ACLs konfiguriert."                = "No port ACLs configured."
    "Keine Prüfpunkte vorhanden."                  = "No checkpoints present."
    "Keine QoS-Richtlinien konfiguriert."          = "No QoS policies configured."
    "Keine relevanten Regeln gefunden."            = "No relevant rules found."
    "Keine Replikationen konfiguriert."            = "No replications configured."
    "Keine SCSI-Controller gefunden."              = "No SCSI controllers found."
    "Keine Software gefunden."                     = "No software found."
    "Keine spezielle Reihenfolge konfiguriert."    = "No specific order configured."
    "Keine Storage Pools vorhanden."               = "No storage pools present."
    "Keine Storage QoS Policies konfiguriert."     = "No storage QoS policies configured."
    "Keine Switch-Erweiterungen gefunden."         = "No switch extensions found."
    "Keine Traffic Classes konfiguriert."          = "No traffic classes configured."
    "Keine Updates gefunden."                      = "No updates found."
    "Keine virtuellen Datenträger vorhanden."      = "No virtual disks present."
    "Keine virtuellen FC-Adapter konfiguriert."    = "No virtual FC adapters configured."
    "Keine virtuellen Festplatten gefunden."       = "No virtual hard disks found."
    "Keine virtuellen Switches gefunden."          = "No virtual switches found."
    "Keine VLAN-Konfiguration gefunden."           = "No VLAN configuration found."
    "Keine VM-Gruppen konfiguriert."               = "No VM groups configured."
    "Keine VMs auf diesem Host."                   = "No VMs on this host."
    "Keine VMs gefunden."                          = "No VMs found."
    "Keine Volumes gefunden."                      = "No volumes found."
    "Alle VMs verwenden den Standardpfad."         = "All VMs use the default path."
    "Keine dedizierten Migrationsnetzwerke definiert." = "No dedicated migration networks defined."
    "PFC nicht konfiguriert."                      = "PFC not configured."
    "Keine kritischen Ereignisse in den letzten 7 Tagen gefunden." = "No critical events found in the last 7 days."
    "Keine VBS-Informationen verfügbar."           = "No VBS information available."
    "Keine Verbindung möglich."                    = "No connection possible."
    "Keine Verbindung möglich (weder WsMan noch DCOM). Bitte Netzwerk/Firewall prüfen." = "No connection possible (neither WsMan nor DCOM). Please check network/firewall."
    "Keine VHD/VHDX-Dateien im Standardpfad gefunden oder Zugriff nicht möglich." = "No VHD/VHDX files found in the default path or access not possible."
    "Keine Produkte über SecurityCenter2 registriert (bei Server-Betriebssystemen normal)." = "No products registered via SecurityCenter2 (normal on server operating systems)."
    "Das ActiveDirectory-Modul ist nicht verfügbar." = "The ActiveDirectory module is not available."
    "Das ActiveDirectory-Modul ist nicht verfügbar - Delegierung konnte nicht geprüft werden." = "The ActiveDirectory module is not available - delegation could not be checked."
    "Das Modul FailoverClusters ist nicht verfügbar. Cluster-Dokumentation wurde übersprungen." = "The FailoverClusters module is not available. Cluster documentation was skipped."
    "Modul FailoverClusters nicht verfügbar."      = "FailoverClusters module not available."
    "Kein Failover-Cluster erkannt (Standalone-Host) oder keine Berechtigung." = "No failover cluster detected (standalone host) or no permission."
    "Kein Failover-Cluster erkannt oder keine CSVs vorhanden." = "No failover cluster detected or no CSVs present."
    "Kein Failover-Cluster erkannt."               = "No failover cluster detected."
    "Firewall-Informationen konnten nicht abgefragt werden (WinRM erforderlich)." = "Firewall information could not be queried (WinRM required)."
    "Netzwerkadapter-Details konnten nicht abgefragt werden (WinRM erforderlich)." = "Network adapter details could not be queried (WinRM required)."
    "QoS-Informationen konnten nicht abgefragt werden (WinRM / DCB-Feature erforderlich)." = "QoS information could not be queried (WinRM / DCB feature required)."
    "SMB-Informationen konnten nicht abgefragt werden (WinRM erforderlich)." = "SMB information could not be queried (WinRM required)."
    "Storage-Informationen konnten nicht abgefragt werden (WinRM erforderlich)." = "Storage information could not be queried (WinRM required)."
    "Windows Features konnten nicht abgefragt werden (WinRM erforderlich)." = "Windows features could not be queried (WinRM required)."
    "Zeitkonfiguration konnte nicht abgefragt werden (WinRM erforderlich)." = "Time configuration could not be queried (WinRM required)."
    "Host Guardian Service Client ist nicht installiert bzw. nicht konfiguriert (kein Guarded Fabric)." = "Host Guardian Service client is not installed or not configured (no guarded fabric)."
    "Microsoft Defender ist nicht installiert oder nicht abfragbar (Drittanbieter-AV möglich)." = "Microsoft Defender is not installed or cannot be queried (third-party AV possible)."
    "MPIO ist nicht installiert oder nicht abfragbar." = "MPIO is not installed or cannot be queried."
    "Windows Server Backup ist nicht installiert oder nicht abfragbar." = "Windows Server Backup is not installed or cannot be queried."
    "SMBv1 gilt als unsicher und sollte auf Hyper-V Hosts deaktiviert sein." = "SMBv1 is considered insecure and should be disabled on Hyper-V hosts."

    # --- Tabellen-Spaltennamen ---
    "Adaptername"                                  = "Adapter name"
    "Adresse"                                      = "Address"
    "Akt. Takt (MHz)"                              = "Current clock (MHz)"
    "Aktion"                                       = "Action"
    "Aktuelle Zeit"                                = "Current time"
    "Alle Server zugelassen"                       = "All servers allowed"
    "Alter (Tage)"                                 = "Age (days)"
    "Angehalten"                                   = "Paused"
    "Angeschl. Geräte"                             = "Connected devices"
    "Antivirus aktiviert"                          = "Antivirus enabled"
    "Anzahl Knoten"                                = "Number of nodes"
    "Anzahl vCPU"                                  = "Number of vCPUs"
    "Anzahl Versionen"                             = "Number of versions"
    "Anzahl VHDs"                                  = "Number of VHDs"
    "Anzahl VMs auf Host"                          = "Number of VMs on host"
    "Anzeigename"                                  = "Display name"
    "Architektur"                                  = "Architecture"
    "Ausgehend (Standard)"                         = "Outbound (default)"
    "Ausgeschaltet"                                = "Powered off"
    "Authentifizierung"                            = "Authentication"
    "Authentifizierungstyp"                        = "Authentication type"
    "Auto Metrik"                                  = "Auto metric"
    "Automatische Prüfpunkte"                      = "Automatic checkpoints"
    "Bandbreite (%)"                               = "Bandwidth (%)"
    "Bandbreitenmodus"                             = "Bandwidth mode"
    "Basis-Prozessor"                              = "Base processor"
    "Bedarf (Demand)"                              = "Demand"
    "Belegt auf Disk"                              = "Used on disk"
    "Belegt_%"                                     = "Used_%"
    "Belegter Speicher"                            = "Used memory"
    "Beliebiges Netzwerk verwenden"                = "Use any network"
    "Bereitstellung"                               = "Provisioning"
    "Beschreibung"                                 = "Description"
    "Besitzergruppe"                               = "Owner group"
    "Besitzerknoten"                               = "Owner node"
    "Betriebsstatus"                               = "Operational status"
    "Betriebszeit"                                 = "Uptime"
    "Bewertung"                                    = "Assessment"
    "Bewertung Authentifizierung"                  = "Assessment authentication"
    "Bewertung CPU"                                = "Assessment CPU"
    "Bewertung Netzwerk"                           = "Assessment network"
    "Bewertung RAM"                                = "Assessment RAM"
    "Bezeichnung"                                  = "Designation"
    "Bindung an Host"                              = "Binding to host"
    "BIOS Datum"                                   = "BIOS date"
    "BIOS/UEFI Version"                            = "BIOS/UEFI version"
    "Block Cache Größe (MB)"                       = "Block cache size (MB)"
    "Blockgröße"                                   = "Block size"
    "Bootreihenfolge"                              = "Boot order"
    "Bustyp"                                       = "Bus type"
    "Checkpointpfad"                               = "Checkpoint path"
    "Cluster-Funktionsebene"                       = "Cluster functional level"
    "Clustername"                                  = "Cluster name"
    "Controller-Nummer"                            = "Controller number"
    "CPU-Auslastung %"                             = "CPU usage %"
    "Credential Guard Status"                      = "Credential Guard status"
    "Datei"                                        = "File"
    "Dateisystem"                                  = "File system"
    "Dateityp"                                     = "File type"
    "Delegierung zu"                               = "Delegation to"
    "DEP (NX/XD) verfügbar"                        = "DEP (NX/XD) available"
    "DEP für Treiber aktiv"                        = "DEP for drivers active"
    "Dienstname"                                   = "Service name"
    "Disk Nr."                                     = "Disk no."
    "Domäne"                                       = "Domain"
    "Domäne (DNS)"                                 = "Domain (DNS)"
    "Domänencontroller"                            = "Domain controller"
    "Domänenmodus"                                 = "Domain mode"
    "Drain-Status"                                 = "Drain status"
    "Drosselung (bps)"                             = "Throttling (bps)"
    "Durchschn. CPU (MHz)"                         = "Avg. CPU (MHz)"
    "Durchschn. RAM (MB)"                          = "Avg. RAM (MB)"
    "Dyn. Gewicht"                                 = "Dyn. weight"
    "Dyn. Speicher"                                = "Dyn. memory"
    "Dynamische MAC"                               = "Dynamic MAC"
    "Dynamisches Quorum"                           = "Dynamic quorum"
    "Echtzeitschutz"                               = "Real-time protection"
    "Eingehend (MB)"                               = "Inbound (MB)"
    "Eingehend (Standard)"                         = "Inbound (default)"
    "Empfangsqueues"                               = "Receive queues"
    "Energiesparplan"                              = "Power plan"
    "Enhanced Session Mode (Host)"                 = "Enhanced session mode (host)"
    "Enhanced Session Mode Richtlinie (Host)"      = "Enhanced session mode policy (host)"
    "Erstellt"                                     = "Created"
    "Erstellt am"                                  = "Created on"
    "Erweiterung"                                  = "Extension"
    "Fallback konfiguriert"                        = "Fallback configured"
    "Fehleranzahl"                                 = "Error count"
    "Fehlerbehandlung"                             = "Error handling"
    "Fibre Channel WWPN (Max)"                     = "Fibre channel WWPN (max)"
    "Fibre Channel WWPN (Min)"                     = "Fibre channel WWPN (min)"
    "Flow Control aktiv"                           = "Flow control active"
    "Fragmentierung %"                             = "Fragmentation %"
    "Frei"                                         = "Free"
    "Frei %"                                       = "Free %"
    "Frei_%"                                       = "Free_%"
    "Frei_GB"                                      = "Free_GB"
    "Freigabe"                                     = "Share"
    "Freigabe (Shared)"                            = "Shared"
    "Gast-Hostname"                                = "Guest host name"
    "Geclustert"                                   = "Clustered"
    "Geräte"                                       = "Devices"
    "Gesamt_GB"                                    = "Total_GB"
    "Gesamtspeicher (GB)"                          = "Total memory (GB)"
    "Gesamtstatus"                                 = "Overall status"
    "Gesamtstruktur"                               = "Forest"
    "Gesamtstrukturmodus"                          = "Forest mode"
    "Geschwindigkeit"                              = "Speed"
    "Gespeichert"                                  = "Saved"
    "Gewichtung"                                   = "Weight"
    "Globale Kataloge"                             = "Global catalogs"
    "Größe"                                        = "Size"
    "Größe_GB"                                     = "Size_GB"
    "Gruppe"                                       = "Group"
    "Gruppen-Mitglieder"                           = "Group members"
    "Gruppenname"                                  = "Group name"
    "Guest Services aktiv"                         = "Guest services active"
    "Hersteller"                                   = "Manufacturer"
    "Host ist Guarded"                             = "Host is guarded"
    "Host-RAM gesamt"                              = "Host RAM total"
    "HW-Threads pro Kern"                          = "HW threads per core"
    "Hyperthreading"                               = "Hyper-threading"
    "Hypervisor aktiv"                             = "Hypervisor active"
    "InitGröße_MB"                                 = "InitSize_MB"
    "InstallDatum"                                 = "InstallDate"
    "InstalliertAm"                                = "InstalledOn"
    "InstalliertVon"                               = "InstalledBy"
    "Instanz-ID"                                   = "Instance ID"
    "Integration Services Ver."                    = "Integration services ver."
    "InterfaceTyp"                                 = "Interface type"
    "IOV (SR-IOV) Unterstützung"                   = "IOV (SR-IOV) support"
    "IOV Gewichtung"                               = "IOV weight"
    "IOV Queue Paare"                              = "IOV queue pairs"
    "IOV Support Grund"                            = "IOV support reason"
    "IP-Adressen"                                  = "IP addresses"
    "IPAdresse"                                    = "IPAddress"
    "IS Version"                                   = "IS version"
    "ISO-Pfad"                                     = "ISO path"
    "Kerberos Port (HTTP)"                         = "Kerberos port (HTTP)"
    "Kerne"                                        = "Cores"
    "KMS Server"                                   = "KMS server"
    "Knoten"                                       = "Node"
    "Knotengewicht"                                = "Node weight"
    "Kompatibilität (ältere OS)"                   = "Compatibility (older OS)"
    "Kompatibilität (Migration)"                   = "Compatibility (migration)"
    "Komprimierung"                                = "Compression"
    "Konfiguration"                                = "Configuration"
    "Konfigurationspfad"                           = "Configuration path"
    "Konfigurierte Services"                       = "Configured services"
    "Konsolenmodus"                                = "Console mode"
    "Konto"                                        = "Account"
    "Kritische Fehleraktion"                       = "Critical error action"
    "Lastverteilung"                               = "Load balancing"
    "Laufend"                                      = "Running"
    "Laufende Services"                            = "Running services"
    "Laufwerk"                                     = "Drive"
    "Legacy CPU Perf. Zähler"                      = "Legacy CPU perf. counter"
    "Letzte Anmeldung"                             = "Last logon"
    "Letzte Replikation"                           = "Last replication"
    "Letzte Sicherung"                             = "Last backup"
    "Letzter Fehler"                               = "Last error"
    "LetzterBoot"                                  = "LastBoot"
    "Letztes Ergebnis"                             = "Last result"
    "Limit (%)"                                    = "Limit (%)"
    "Live Migration aktiviert"                     = "Live migration enabled"
    "Lizenzhinweis"                                = "License note"
    "Lizenzstatus"                                 = "License status"
    "Log. Prozessoren"                             = "Log. processors"
    "Log. Prozessoren gesamt"                      = "Log. processors total"
    "Logische Proz."                               = "Logical proc."
    "Logische Prozessoren"                         = "Logical processors"
    "Logische Prozessoren (Host)"                  = "Logical processors (host)"
    "Logische Sektorgr."                           = "Logical sector size"
    "Lokale Adresse"                               = "Local address"
    "Lokale IP"                                    = "Local IP"
    "Lokaler Port"                                 = "Local port"
    "LsaCfgFlags (Richtlinie)"                     = "LsaCfgFlags (policy)"
    "MAC-Adressbereich (Max)"                      = "MAC address range (max)"
    "MAC-Adressbereich (Min)"                      = "MAC address range (min)"
    "MAC-Adresse"                                  = "MAC address"
    "Management OS teilt Adapter"                  = "Management OS shares adapter"
    "Max Takt (MHz)"                               = "Max clock (MHz)"
    "Max. Bandbreite"                              = "Max. bandwidth"
    "Max. Bandbreite (Abs)"                        = "Max. bandwidth (abs)"
    "Max. gleichzeitige Storage-Migrationen"       = "Max. concurrent storage migrations"
    "Max. gleichzeitige VM-Migrationen"            = "Max. concurrent VM migrations"
    "Max. Größe"                                   = "Max. size"
    "Max. IOPS (QoS)"                              = "Max. IOPS (QoS)"
    "Max. Nodes pro Socket"                        = "Max. nodes per socket"
    "Max. Nodes/Socket"                            = "Max. nodes/socket"
    "Max. Proz. pro NUMA-Node"                     = "Max. proc. per NUMA node"
    "Max. Prozessor"                               = "Max. processor"
    "Max. Prozessoren"                             = "Max. processors"
    "Max. Prozessoren/NUMA-Node"                   = "Max. processors/NUMA node"
    "Max. Queue Pairs"                             = "Max. queue pairs"
    "Max. RAM (MB)"                                = "Max. RAM (MB)"
    "MaxGröße_MB"                                  = "MaxSize_MB"
    "Medientyp"                                    = "Media type"
    "Meldung"                                      = "Message"
    "Messzeitraum"                                 = "Measurement period"
    "Metrik"                                       = "Metric"
    "Min. Bandbreite (Abs)"                        = "Min. bandwidth (abs)"
    "Min. Bandbreite (Gew.)"                       = "Min. bandwidth (weight)"
    "Min. IOPS (QoS)"                              = "Min. IOPS (QoS)"
    "Min. RAM (MB)"                                = "Min. RAM (MB)"
    "Mitglied"                                     = "Member"
    "Mitglieder"                                   = "Members"
    "Modell"                                       = "Model"
    "Modus"                                        = "Mode"
    "Monitoring Intervall"                         = "Monitoring interval"
    "Monitoring Startzeit"                         = "Monitoring start time"
    "Nächste Sicherung"                            = "Next backup"
    "NetBIOS-Name"                                 = "NetBIOS name"
    "Netzwerk"                                     = "Network"
    "Netzwerkadressen"                             = "Network addresses"
    "Notizen"                                      = "Notes"
    "Num Lock aktiviert"                           = "Num lock enabled"
    "NUMA Spanning aktiviert"                      = "NUMA spanning enabled"
    "OS Version"                                   = "OS version"
    "Partitionen"                                  = "Partitions"
    "Partitionsstil"                               = "Partition style"
    "Pause nach Boot-Fehler"                       = "Pause after boot failure"
    "Perfmon Zähler (PMU)"                         = "Perfmon counter (PMU)"
    "Performance-Option"                           = "Performance option"
    "Pfad"                                         = "Path"
    "Phys. Prozessoren"                            = "Phys. processors"
    "Physische Sektorgr."                          = "Physical sector size"
    "Physischer Adapter"                           = "Physical adapter"
    "Pipe-Pfad"                                    = "Pipe path"
    "Pool-Belegung"                                = "Pool usage"
    "Primärserver"                                 = "Primary server"
    "Priorität"                                    = "Priority"
    "Produkt"                                      = "Product"
    "Profil"                                       = "Profile"
    "Protokoll"                                    = "Protocol"
    "Protokolldatei"                               = "Log file"
    "Protokollübergang"                            = "Log transition"
    "Prozess"                                      = "Process"
    "Prozessor"                                    = "Processor"
    "Prozessor-Auslastung"                         = "Processor usage"
    "Prozessorarchitektur"                         = "Processor architecture"
    "Prozessoren (Anzahl)"                         = "Processors (count)"
    "Prüfpunkt"                                    = "Checkpoint"
    "Prüfpunkte"                                   = "Checkpoints"
    "Prüfpunkte aktiviert"                         = "Checkpoints enabled"
    "Prüfpunktpfad"                                = "Checkpoint path"
    "Puffer (%)"                                   = "Buffer (%)"
    "PXE Netzwerkadapter"                          = "PXE network adapter"
    "Quelle"                                       = "Source"
    "Quorum-Typ"                                   = "Quorum type"
    "RAM an VMs zugewiesen"                        = "RAM assigned to VMs"
    "RAM Startwert"                                = "RAM startup value"
    "RAM zugewiesen"                               = "RAM assigned"
    "RAM_Belegt_%"                                 = "RAM_Used_%"
    "RAM_Frei_GB"                                  = "RAM_Free_GB"
    "RAM_Gesamt_GB"                                = "RAM_Total_GB"
    "RAM-Auslastung durch VMs"                     = "RAM usage by VMs"
    "RDMA aktiviert"                               = "RDMA enabled"
    "Regel"                                        = "Rule"
    "Relative Gewichtung"                          = "Relative weight"
    "Replikat-Server aktiviert"                    = "Replica server enabled"
    "Replikation zugelassen von"                   = "Replication allowed from"
    "Replikationsfrequenz (s)"                     = "Replication frequency (s)"
    "Replikatserver"                               = "Replica server"
    "Reserve (%)"                                  = "Reserve (%)"
    "Resilienz"                                    = "Resiliency"
    "Resource Metering Intervall"                  = "Resource metering interval"
    "Resource Metering Speicherintervall"          = "Resource metering save interval"
    "Ressource"                                    = "Resource"
    "Richtlinie"                                   = "Policy"
    "Richtung"                                     = "Direction"
    "Rolle"                                        = "Role"
    "RSS aktiviert"                                = "RSS enabled"
    "S2D aktiviert"                                = "S2D enabled"
    "SAN-Name"                                     = "SAN name"
    "Schnittstelle"                                = "Interface"
    "Secure Boot Vorlage"                          = "Secure boot template"
    "Seriennummer"                                 = "Serial number"
    "SET (Embedded Teaming)"                       = "SET (embedded teaming)"
    "Shared Volumes Root"                          = "Shared volumes root"
    "Sicherungsrichtlinie vorhanden"               = "Backup policy present"
    "Signatur"                                     = "Signature"
    "Signatur aktualisiert"                        = "Signature updated"
    "SignaturDatum"                                = "SignatureDate"
    "Signaturversion"                              = "Signature version"
    "SLAT (EPT/NPT)"                               = "SLAT (EPT/NPT)"
    "Smart Paging aktiv"                           = "Smart paging active"
    "Smart Paging Pfad"                            = "Smart paging path"
    "SMB Server Protokoll"                         = "SMB server protocol"
    "Sockel"                                       = "Socket"
    "Software"                                     = "Software"
    "Spalten"                                      = "Columns"
    "Speicher belegt (%)"                          = "Memory used (%)"
    "Speicher gesamt (MB)"                         = "Memory total (MB)"
    "Speicher verfügbar (MB)"                      = "Memory available (MB)"
    "Speicherkapazität"                            = "Storage capacity"
    "Speicherort"                                  = "Location"
    "Speicherstatus"                               = "Memory status"
    "SR-IOV aktiv"                                 = "SR-IOV active"
    "SR-IOV aktiviert"                             = "SR-IOV enabled"
    "SR-IOV Grund"                                 = "SR-IOV reason"
    "Standard-Mindestbandbreite (abs)"             = "Default minimum bandwidth (abs)"
    "Standard-Mindestbandbreite (Gew.)"            = "Default minimum bandwidth (weight)"
    "Standard-Resilienz"                           = "Default resiliency"
    "Standard-Speicherort"                         = "Default location"
    "Standorte"                                    = "Sites"
    "Start-Aktion"                                 = "Start action"
    "Startspeicher"                                = "Startup memory"
    "Starttyp"                                     = "Startup type"
    "Startverzögerung (s)"                         = "Start delay (s)"
    "Status (hex)"                                 = "Status (hex)"
    "Steckplatz"                                   = "Slot"
    "Stimmgewicht"                                 = "Vote weight"
    "Stop-Aktion"                                  = "Stop action"
    "Subnetz"                                      = "Subnet"
    "Subnetzmaske"                                 = "Subnet mask"
    "Takt_MHz"                                     = "Clock_MHz"
    "Team-Mitglieder"                              = "Team members"
    "Teaming-Modus"                                = "Teaming mode"
    "Teil-Key"                                     = "Partial key"
    "Teilenummer"                                  = "Part number"
    "Timeout Fehleraktion"                         = "Timeout error action"
    "Treiberdatum"                                 = "Driver date"
    "Treiberversion"                               = "Driver version"
    "Trunk VLAN-Liste"                             = "Trunk VLAN list"
    "Typ"                                          = "Type"
    "Übergeordnet"                                 = "Parent"
    "Übergeordnet (Parent)"                        = "Parent"
    "Uneingeschränkte Delegierung"                 = "Unconstrained delegation"
    "Unterstützung"                                = "Support"
    "Upgrade-Version"                              = "Upgrade version"
    "Uptime_Tage"                                  = "Uptime_days"
    "VBS aktiviert (Registry)"                     = "VBS enabled (registry)"
    "VBS Status"                                   = "VBS status"
    "vCPU : pCPU Verhältnis"                       = "vCPU : pCPU ratio"
    "vCPUs zugewiesen (laufend)"                   = "vCPUs assigned (running)"
    "Verbindungsstatus"                            = "Connection status"
    "Verbleibend (Tage)"                           = "Remaining (days)"
    "Verbunden"                                    = "Connected"
    "Verbundene VMs"                               = "Connected VMs"
    "Verfügbare Sicherheits-EIG"                   = "Available security features"
    "Verfügbare VM-Queues"                         = "Available VM queues"
    "Verschlüsselt"                                = "Encrypted"
    "Vertrauensgruppe"                             = "Trust group"
    "Virtualisiert"                                = "Virtualized"
    "Virtualisierung"                              = "Virtualization"
    "Virtualisierung in Firmware"                  = "Virtualization in firmware"
    "Virtueller Switch"                            = "Virtual switch"
    "VLAN-Modus"                                   = "VLAN mode"
    "VM Migration aktiviert"                       = "VM migration enabled"
    "VM-Mitglieder"                                = "VM members"
    "VM-Pfad"                                      = "VM path"
    "VM-Version"                                   = "VM version"
    "VMMQ aktiviert"                               = "VMMQ enabled"
    "VMQ aktiviert"                                = "VMQ enabled"
    "VMQ Gewichtung"                               = "VMQ weight"
    "VMs gesamt"                                   = "VMs total"
    "VMs laufend"                                  = "VMs running"
    "vNIC (Host)"                                  = "vNIC (host)"
    "Vollduplex"                                   = "Full duplex"
    "Volumenname"                                  = "Volume name"
    "Vorlage"                                      = "Template"
    "VRSS aktiviert"                               = "VRSS enabled"
    "VSS-Snapshot Frequenz"                        = "VSS snapshot frequency"
    "vTPM aktiviert"                               = "vTPM enabled"
    "Wert"                                         = "Value"
    "Wiederherstellungspunkte"                     = "Recovery points"
    "Zähler"                                       = "Counter"
    "Zeit"                                         = "Time"
    "Zeitplan"                                     = "Schedule"
    "Zeitpunkt"                                    = "Timestamp"
    "Zeitquelle"                                   = "Time source"
    "Zeitzone"                                     = "Time zone"
    "Zertifikat Port (HTTPS)"                      = "Certificate port (HTTPS)"
    "Zertifikat Thumbprint"                        = "Certificate thumbprint"
    "Zeugen-Ressource"                             = "Witness resource"
    "Zugewiesen"                                   = "Assigned"
    "Zugewiesen (aktuell)"                         = "Assigned (current)"
    "Zuletzt geändert"                             = "Last modified"
    "Zuordnung"                                    = "Assignment"
    "Zuordnungseinheit"                            = "Allocation unit"
    "Zustand"                                      = "State"
    "Zustand &amp; Migration verschlüsselt"        = "State &amp; migration encrypted"

    # --- Dokumentrahmen ---
    "Hyper-V - Umgebungsdokumentation"             = "Hyper-V - Environment Documentation"
    "Inhaltsverzeichnis"                           = "Table of contents"
    "Dokumentierte Hosts"                          = "Documented hosts"
    "Host-Betriebssystem"                          = "Host operating system"
    "Erstellt von"                                 = "Created by"
    "Sektionen"                                    = "Sections"
    "Warnungen"                                    = "Warnings"
    "auf"                                          = "on"

    # --- GUI ---
    "Hyper-V Dokumentation"                        = "Hyper-V Documentation"
    "Host · VMs · Netzwerke · Speicher · Cluster · Replica · Multi-Format Export" = "Host · VMs · Networking · Storage · Cluster · Replica · Multi-format export"
    "Hyper-V Hosts"                                = "Hyper-V hosts"
    "Alle auswählen"                               = "Select all"
    "Organisation"                                 = "Organization"
    "Ausgabepfad"                                  = "Output path"
    "Ausgabeformate"                               = "Output formats"
    "Dokumentationsbereiche"                       = "Documentation sections"
    "Abbrechen"                                    = "Cancel"
    "  🚀  Dokumentation starten"                  = "  🚀  Start documentation"
    "Ausgabeverzeichnis wählen"                    = "Select output directory"
    "Fehler: Bitte mindestens einen Hyper-V Host auswählen." = "Error: Please select at least one Hyper-V host."
    "Fehler: Bitte einen Ausgabepfad angeben."     = "Error: Please specify an output path."
    "Fehler: Bitte mindestens einen Dokumentationsbereich auswählen." = "Error: Please select at least one documentation section."
    "Fehler: Bitte mindestens ein Ausgabeformat wählen." = "Error: Please select at least one output format."

    # --- Konsole & Log ---
    "Administrator-Rechte bestätigt. Skript wird ausgeführt..." = "Administrator privileges confirmed. Running script..."
    "Das Skript erfordert Administrator-Rechte. Starte neu mit erhöhten Rechten..." = "The script requires administrator privileges. Restarting elevated..."
    "Abgebrochen durch Benutzer (GUI)."            = "Cancelled by user (GUI)."
    "Keine Hyper-V Hosts angegeben. Bitte -HyperVServers verwenden oder die GUI nutzen." = "No Hyper-V hosts specified. Please use -HyperVServers or the GUI."
    "Ausgabeverzeichnis erstellt"                  = "Output directory created"
    "Hyper-V Dokumentation gestartet"              = "Hyper-V documentation started"
    "Zielhosts"                                    = "Target hosts"
    "Ausgabepfad:"                                 = "Output path:"
    "Ausgabeformate:"                              = "Output formats:"
    "Gewählte Sektionen"                           = "Selected sections"
    "Verbindungsmodus: CIM mit automatischem DCOM Fallback" = "Connection mode: CIM with automatic DCOM fallback"
    "Hyper-V PowerShell-Modul nicht verfügbar. Skript wird beendet." = "Hyper-V PowerShell module not available. Script will exit."
    "Hyper-V PowerShell-Modul nicht verfügbar. Bitte RSAT-Hyper-V-Tools installieren." = "Hyper-V PowerShell module not available. Please install RSAT-Hyper-V-Tools."
    "=== Prüfe Erreichbarkeit der Hosts ==="       = "=== Checking host reachability ==="
    "=== Starte Datensammlung"                     = "=== Starting data collection"
    "=== Generiere HTML-Dokument ==="              = "=== Generating HTML document ==="
    "=== Generiere PDF-Dokument ==="               = "=== Generating PDF document ==="
    "=== Generiere Markdown-Dokument ==="          = "=== Generating Markdown document ==="
    "Dokumentation erfolgreich erstellt!"          = "Documentation created successfully!"
    "Log-Datei"                                    = "Log file"
    "Warnungen:"                                   = "Warnings:"
    "Hyper-V Dokumentation abgeschlossen!"         = "Hyper-V documentation completed!"
    "Dokumentation jetzt oeffnen? (J/N)"           = "Open documentation now? (Y/N)"
    "Kritischer Fehler! Details:"                  = "Critical error! Details:"
    "Skript beendet um"                            = "Script finished at"
}

function Get-T {
    <#
    .SYNOPSIS
        Übersetzt einen deutschen Text in die aktive Sprache.
        Unbekannte Texte werden unverändert zurückgegeben.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)]
        [AllowEmptyString()]
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    if ($script:Language -eq "DE") { return $Text }
    if ($script:Dict.ContainsKey($Text)) { return $script:Dict[$Text] }
    return $Text
}

function ConvertTo-LocalizedHtml {
    <#
    .SYNOPSIS
        Übersetzt Überschriften und Hinweistexte innerhalb eines HTML-Fragments.
        Datenzellen bleiben unangetastet.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$Html
    )

    if ($script:Language -eq "DE" -or [string]::IsNullOrEmpty($Html)) { return $Html }

    $evaluator = { param($m) $m.Groups[1].Value + (Get-T $m.Groups[2].Value) + $m.Groups[3].Value }

    foreach ($pattern in @(
        '(<h3(?:\s[^>]*)?>)([^<]+)(</h3>)',
        '(<h4(?:\s[^>]*)?>)([^<]+)(</h4>)',
        "(<p class='no-data'>)([^<]+)(</p>)",
        "(<p class='info'>)([^<]+)(</p>)",
        '(<strong>)([^<]+)(</strong>)'
    )) {
        $Html = [regex]::Replace($Html, $pattern, $evaluator)
    }

    return $Html
}

#region ============================================================
# HILFSFUNKTIONEN
#endregion ============================================================

function Write-Log {
    <#
    .SYNOPSIS
        Schreibt eine Nachricht in die Log-Datei und auf die Konsole.
    .PARAMETER Message
        Die zu protokollierende Nachricht.
    .PARAMETER Level
        Log-Level: INFO, WARNING, ERROR
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )

    try {
        $logTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$logTimestamp] [$Level] $Message"

        # Sicherstellen, dass das Verzeichnis existiert
        $logDir = Split-Path $script:LogFile -Parent
        if (-not (Test-Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        # In Datei schreiben
        Add-Content -Path $script:LogFile -Value $logEntry -Encoding UTF8

        # Konsolenausgabe mit Farbe
        switch ($Level) {
            "INFO"    { Write-Host $logEntry -ForegroundColor Green }
            "WARNING" { Write-Host $logEntry -ForegroundColor Yellow; $script:WarningCount++ }
            "ERROR"   { Write-Host $logEntry -ForegroundColor Red; $script:ErrorCount++ }
        }
    }
    catch {
        Write-Host "FEHLER beim Schreiben der Log-Datei: $_" -ForegroundColor Red
    }
}

function New-HTMLSection {
    <#
    .SYNOPSIS
        Erstellt eine neue HTML-Sektion mit Überschrift und fügt sie dem Inhaltsverzeichnis hinzu.
    .PARAMETER Title
        Titel der Sektion.
    .PARAMETER Content
        HTML-Inhalt der Sektion.
    .PARAMETER Level
        Überschriften-Level (1-3). Standard: 2
    .PARAMETER Category
        Themenbereich für die Sortierung im Inhaltsverzeichnis (z. B. "Hardware & OS", "Hyper-V Host").
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Content,

        [Parameter(Mandatory = $false)]
        [int]$Level = 2,

        [Parameter(Mandatory = $false)]
        [string]$Category = ""
    )

    $script:SectionCounter++
    $anchorId = "section_$($script:SectionCounter)"

    $Title    = Get-T $Title
    $Category = Get-T $Category
    $Content  = ConvertTo-LocalizedHtml -Html $Content

    # Inhaltsverzeichnis-Eintrag
    [void]$script:TOCEntries.Add(@{
        Title    = $Title
        Anchor   = $anchorId
        Level    = $Level
        Category = $Category
    })

    # HTML-Sektion erstellen
    $sectionHTML = @"
    <div class="section">
        <h$Level id="$anchorId">$($script:SectionCounter). $Title</h$Level>
        $Content
    </div>
"@

    [void]$script:HTMLSections.Add($sectionHTML)
}

function ConvertTo-HTMLTable {
    <#
    .SYNOPSIS
        Konvertiert ein Array von Objekten in eine formatierte HTML-Tabelle.
    .PARAMETER Data
        Die zu konvertierenden Daten.
    .PARAMETER Properties
        Optionale Eigenschaftsauswahl.
    .PARAMETER NoDataMessage
        Nachricht, wenn keine Daten vorhanden sind.
    .PARAMETER HeaderColor
        Optionale Hintergrundfarbe für die Tabellenkopfzeile (z. B. "#005a9e").
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object[]]$Data,

        [Parameter(Mandatory = $false)]
        [string[]]$Properties,

        [Parameter(Mandatory = $false)]
        [string]$NoDataMessage = "Keine Daten verfügbar.",

        [Parameter(Mandatory = $false)]
        [string]$HeaderColor = ""
    )

    # Null-Einträge entfernen
    if ($Data) { $Data = @($Data | Where-Object { $null -ne $_ }) }

    if (-not $Data -or $Data.Count -eq 0) {
        return "<p class='no-data'>$(Get-T $NoDataMessage)</p>"
    }

    try {
        if ($Properties) {
            $Data = $Data | Select-Object -Property $Properties
        }

        $html = "<table>`n<thead>`n<tr>"

        # Header ermitteln
        $headers = $Data[0].PSObject.Properties.Name
        $headerStyle = if ($HeaderColor) { " style='background-color: $HeaderColor; color: white;'" } else { "" }
        foreach ($header in $headers) {
            $html += "<th$headerStyle>$(Get-T $header)</th>"
        }
        $html += "</tr>`n</thead>`n<tbody>`n"

        # Datenzeilen
        $rowIndex = 0
        foreach ($row in $Data) {
            $rowClass = if ($rowIndex % 2 -eq 0) { "even" } else { "odd" }
            $html += "<tr class='$rowClass'>"
            foreach ($header in $headers) {
                $value = $row.$header
                if ($null -eq $value) { $value = "-" }
                # "Link"-Spalten nicht encodieren (HTML-Tags erlauben)
                if ($header -eq "Link") {
                    $html += "<td>$value</td>"
                } else {
                    $html += "<td>$([System.Web.HttpUtility]::HtmlEncode((Get-T $value.ToString())))</td>"
                }
            }
            $html += "</tr>`n"
            $rowIndex++
        }

        $html += "</tbody>`n</table>"
        return $html
    }
    catch {
        Write-Log -Message "Fehler bei HTML-Tabellen-Konvertierung: $_" -Level "ERROR"
        return "<p class='error'>Fehler bei der Datenkonvertierung: $_</p>"
    }
}

function New-ServerCimSession {
    <#
    .SYNOPSIS
        Erstellt eine CIM-Session zu einem Server mit automatischem WsMan→DCOM Fallback.
    .PARAMETER ComputerName
        Name des Zielservers.
    .OUTPUTS
        CimSession-Objekt oder $null bei Fehler.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    # Methode 1: WsMan (Standard, nutzt WinRM)
    try {
        Write-Log -Message "CIM-Session über WsMan für $ComputerName..." -Level "INFO"
        $session = New-CimSession -ComputerName $ComputerName -ErrorAction Stop
        Write-Log -Message "CIM-Session (WsMan) zu $ComputerName erfolgreich." -Level "INFO"
        return $session
    }
    catch {
        Write-Log -Message "WsMan fehlgeschlagen für ${ComputerName}: $($_.Exception.Message)" -Level "WARNING"
    }

    # Methode 2: DCOM Fallback (kein WinRM nötig, nutzt RPC)
    try {
        Write-Log -Message "CIM-Session über DCOM für $ComputerName..." -Level "INFO"
        $dcomOption = New-CimSessionOption -Protocol Dcom
        $session = New-CimSession -ComputerName $ComputerName -SessionOption $dcomOption -ErrorAction Stop
        Write-Log -Message "CIM-Session (DCOM) zu $ComputerName erfolgreich." -Level "INFO"
        return $session
    }
    catch {
        Write-Log -Message "Auch DCOM fehlgeschlagen für ${ComputerName}: $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

function Get-RemoteRegistryValue {
    <#
    .SYNOPSIS
        Liest Registry-Werte remote über .NET-Methoden (kein WinRM nötig).
    .PARAMETER ComputerName
        Name des Zielcomputers.
    .PARAMETER RegistryPath
        Pfad innerhalb der Registry (HKLM).
    .PARAMETER ValueName
        Name des zu lesenden Wertes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $true)]
        [string]$RegistryPath,

        [Parameter(Mandatory = $true)]
        [string]$ValueName
    )

    try {
        $regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine, $ComputerName
        )
        $subKey = $regKey.OpenSubKey($RegistryPath)
        if ($subKey) {
            $value = $subKey.GetValue($ValueName)
            $subKey.Close()
            $regKey.Close()
            return $value
        }
        $regKey.Close()
        return $null
    }
    catch {
        Write-Log -Message "Remote Registry auf ${ComputerName} fehlgeschlagen ($ValueName): $_" -Level "WARNING"
        return $null
    }
}

function Get-RemoteRegistrySubKeyNames {
    <#
    .SYNOPSIS
        Listet SubKey-Namen eines Registry-Pfads remote über .NET-Methoden.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $true)]
        [string]$RegistryPath
    )

    try {
        $regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine, $ComputerName
        )
        $subKey = $regKey.OpenSubKey($RegistryPath)
        if ($subKey) {
            $names = $subKey.GetSubKeyNames()
            $subKey.Close()
            $regKey.Close()
            return $names
        }
        $regKey.Close()
        return @()
    }
    catch {
        Write-Log -Message "Remote Registry SubKey-Enumeration auf ${ComputerName} (${RegistryPath}) fehlgeschlagen: $_" -Level "WARNING"
        return @()
    }
}

function Test-LocalComputer {
    <#
    .SYNOPSIS
        Prüft, ob der angegebene Name den lokalen Computer bezeichnet.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    $shortName = ($ComputerName -split '\.')[0]
    return ($shortName -eq $env:COMPUTERNAME -or $ComputerName -eq "." -or $ComputerName -eq "localhost")
}

function Invoke-RemoteCommand {
    <#
    .SYNOPSIS
        Führt einen Scriptblock lokal oder remote (via WinRM) aus.
        Bei lokalem Ziel wird kein Remoting verwendet.
    .PARAMETER ComputerName
        Zielserver.
    .PARAMETER ScriptBlock
        Auszuführender Code.
    .PARAMETER ArgumentList
        Optionale Argumente für den Scriptblock.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory = $false)]
        [object[]]$ArgumentList = @()
    )

    try {
        if (Test-LocalComputer -ComputerName $ComputerName) {
            return (& $ScriptBlock @ArgumentList)
        }
        return (Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList -ErrorAction Stop)
    }
    catch {
        Write-Log -Message "Remote-Ausführung auf ${ComputerName} fehlgeschlagen: $($_.Exception.Message)" -Level "WARNING"
        return $null
    }
}

function Test-CommandAvailable {
    <#
    .SYNOPSIS
        Prüft, ob ein Cmdlet in der aktuellen Session verfügbar ist.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    return [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue)
}

function Format-ByteSize {
    <#
    .SYNOPSIS
        Formatiert eine Byte-Angabe in eine lesbare Größe (GB/MB).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$Bytes
    )

    if ($null -eq $Bytes) { return "-" }
    try {
        $b = [double]$Bytes
        if ($b -ge 1TB) { return ("{0:N2} TB" -f ($b / 1TB)) }
        if ($b -ge 1GB) { return ("{0:N2} GB" -f ($b / 1GB)) }
        if ($b -ge 1MB) { return ("{0:N2} MB" -f ($b / 1MB)) }
        if ($b -ge 1KB) { return ("{0:N2} KB" -f ($b / 1KB)) }
        return ("{0} Byte" -f $b)
    }
    catch {
        return "$Bytes"
    }
}

function ConvertTo-DisplayValue {
    <#
    .SYNOPSIS
        Wandelt beliebige Werte (Arrays, $null, Boolean) in einen lesbaren String um.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$Value,

        [Parameter(Mandatory = $false)]
        [string]$EmptyText = "-"
    )

    if ($null -eq $Value) { return $EmptyText }
    if ($Value -is [bool]) { return $(if ($Value) { Get-T "Ja" } else { Get-T "Nein" }) }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = @($Value | Where-Object { $null -ne $_ })
        if ($items.Count -eq 0) { return $EmptyText }
        return ($items -join ', ')
    }
    $text = "$Value"
    if ([string]::IsNullOrWhiteSpace($text)) { return $EmptyText }
    return $text
}

#endregion

#region ============================================================
# HYPER-V MODUL LADEN & UMGEBUNG ERKENNEN
#endregion ============================================================

function Initialize-HyperVEnvironment {
    <#
    .SYNOPSIS
        Lädt das Hyper-V PowerShell-Modul (und optional FailoverClusters / ActiveDirectory).
    #>
    Write-Log -Message "=== Hyper-V Management Tools werden geladen ===" -Level "INFO"

    try {
        if (Get-Module -Name Hyper-V -ErrorAction SilentlyContinue) {
            Write-Log -Message "Hyper-V Modul ist bereits geladen." -Level "INFO"
            $script:HyperVModuleLoaded = $true
        }
        elseif (Get-Module -ListAvailable -Name Hyper-V -ErrorAction SilentlyContinue) {
            Import-Module Hyper-V -ErrorAction Stop -DisableNameChecking
            Write-Log -Message "Hyper-V Modul erfolgreich geladen." -Level "INFO"
            $script:HyperVModuleLoaded = $true
        }
        else {
            Write-Log -Message "Hyper-V PowerShell-Modul nicht gefunden. Bitte RSAT-Hyper-V-Tools installieren." -Level "ERROR"
            $script:HyperVModuleLoaded = $false
        }
    }
    catch {
        Write-Log -Message "Fehler beim Laden des Hyper-V Moduls: $_" -Level "ERROR"
        $script:HyperVModuleLoaded = $false
    }

    # --- FailoverClusters-Modul (optional) ---
    try {
        if (Get-Module -ListAvailable -Name FailoverClusters -ErrorAction SilentlyContinue) {
            Import-Module FailoverClusters -ErrorAction Stop -DisableNameChecking
            $script:ClusterModuleLoaded = $true
            Write-Log -Message "FailoverClusters Modul geladen (Cluster-Dokumentation verfügbar)." -Level "INFO"
        }
        else {
            Write-Log -Message "FailoverClusters Modul nicht verfügbar - Cluster-Sektionen werden übersprungen." -Level "WARNING"
        }
    }
    catch {
        Write-Log -Message "FailoverClusters Modul konnte nicht geladen werden: $_" -Level "WARNING"
    }

    # --- ActiveDirectory-Modul (optional) ---
    try {
        if (Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue) {
            Import-Module ActiveDirectory -ErrorAction Stop -DisableNameChecking
            Write-Log -Message "ActiveDirectory Modul geladen." -Level "INFO"
        }
    }
    catch {
        Write-Log -Message "ActiveDirectory Modul konnte nicht geladen werden: $_" -Level "WARNING"
    }

    return $script:HyperVModuleLoaded
}

function Get-HyperVEdition {
    <#
    .SYNOPSIS
        Ermittelt die Windows Server Version des ersten Hyper-V Hosts sowie
        einen ggf. vorhandenen Failover-Cluster-Namen.
    #>
    Write-Log -Message "=== Ermittle Hyper-V Host-Version ===" -Level "INFO"

    $firstServer = @($HyperVServers)[0]
    if (-not $firstServer) { return }

    try {
        $cimSession = New-ServerCimSession -ComputerName $firstServer
        if ($cimSession) {
            $os = Get-CimInstance -CimSession $cimSession -ClassName Win32_OperatingSystem -ErrorAction Stop
            $script:HyperVEdition = "$($os.Caption) (Build $($os.BuildNumber))"
            Write-Log -Message "Host-Betriebssystem erkannt: $($script:HyperVEdition)" -Level "INFO"
            Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Log -Message "Host-Version konnte nicht ermittelt werden: $_" -Level "WARNING"
        $script:HyperVEdition = "Unbekannt"
    }

    # --- Cluster-Zugehörigkeit prüfen ---
    if ($script:ClusterModuleLoaded) {
        try {
            $cluster = Get-Cluster -Name $firstServer -ErrorAction Stop
            if ($cluster) {
                $script:ClusterName = $cluster.Name
                Write-Log -Message "Failover-Cluster erkannt: $($script:ClusterName)" -Level "INFO"
            }
        }
        catch {
            Write-Log -Message "Kein Failover-Cluster erkannt (Standalone-Host)." -Level "INFO"
        }
    }
}

#endregion

#region ============================================================
# DATENSAMMLUNGS-FUNKTIONEN
#endregion ============================================================

# ---------------------------------------------------------------
# 1. HARDWARE-INFORMATIONEN (CIM/DCOM FALLBACK)
# ---------------------------------------------------------------
function Get-HardwareInformation {
    <#
    .SYNOPSIS
        Sammelt Hardware-Informationen der Hyper-V Hosts über CIM-Sessions
        mit automatischem WsMan→DCOM Fallback.
    #>
    Write-Log -Message "=== Sammle Hardware-Informationen (CIM/DCOM Fallback) ===" -Level "INFO"

    $allHardwareHTML = ""

    foreach ($server in $HyperVServers) {
        try {
            Write-Log -Message "Hardware-Info für Host: $server" -Level "INFO"

            $cimSession = New-ServerCimSession -ComputerName $server
            if (-not $cimSession) {
                $allHardwareHTML += "<h3 class='server-break'>Host: $server</h3>"
                $allHardwareHTML += "<p class='error'>Keine Verbindung möglich (weder WsMan noch DCOM). Bitte Netzwerk/Firewall prüfen.</p>"
                continue
            }

            # --- Betriebssystem ---
            $osInfo = $null
            try {
                $osRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_OperatingSystem -ErrorAction Stop
                $osInfo = [PSCustomObject]@{
                    Computername    = $osRaw.CSName
                    Betriebssystem  = $osRaw.Caption
                    Version         = $osRaw.Version
                    BuildNumber     = $osRaw.BuildNumber
                    Architektur     = $osRaw.OSArchitecture
                    InstallDatum    = $osRaw.InstallDate.ToString("dd.MM.yyyy")
                    LetzterBoot     = $osRaw.LastBootUpTime.ToString("dd.MM.yyyy HH:mm")
                    "Uptime_Tage"   = [math]::Round(((Get-Date) - $osRaw.LastBootUpTime).TotalDays, 1)
                    "RAM_Gesamt_GB" = [math]::Round($osRaw.TotalVisibleMemorySize / 1MB, 2)
                    "RAM_Frei_GB"   = [math]::Round($osRaw.FreePhysicalMemory / 1MB, 2)
                    "RAM_Belegt_%"  = [math]::Round((1 - ($osRaw.FreePhysicalMemory / $osRaw.TotalVisibleMemorySize)) * 100, 1)
                }
            }
            catch {
                Write-Log -Message "OS-Abfrage fehlgeschlagen für ${server}: $_" -Level "ERROR"
            }

            # --- System / Hersteller ---
            $sysInfo = $null
            try {
                $csRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_ComputerSystem -ErrorAction Stop
                $biosRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_BIOS -ErrorAction SilentlyContinue
                $sysInfo = [PSCustomObject]@{
                    Hersteller          = $csRaw.Manufacturer
                    Modell              = $csRaw.Model
                    Seriennummer        = if ($biosRaw) { $biosRaw.SerialNumber } else { "-" }
                    "BIOS/UEFI Version" = if ($biosRaw) { $biosRaw.SMBIOSBIOSVersion } else { "-" }
                    "BIOS Datum"        = if ($biosRaw -and $biosRaw.ReleaseDate) { $biosRaw.ReleaseDate.ToString("dd.MM.yyyy") } else { "-" }
                    Domäne              = $csRaw.Domain
                    "Log. Prozessoren"  = $csRaw.NumberOfLogicalProcessors
                    "Phys. Prozessoren" = $csRaw.NumberOfProcessors
                    "RAM_Gesamt_GB"     = [math]::Round($csRaw.TotalPhysicalMemory / 1GB, 2)
                    Virtualisiert       = if ($csRaw.Model -match "Virtual|VMware|KVM|Xen|HVM") { "⚠️ Ja (Nested?)" } else { "Nein (physisch)" }
                    "Hypervisor aktiv"  = if ($csRaw.HypervisorPresent) { "✅ Ja" } else { "⚠️ Nein" }
                }
            }
            catch {
                Write-Log -Message "System-Abfrage fehlgeschlagen für ${server}: $_" -Level "ERROR"
            }

            # --- CPU ---
            $cpuInfo = $null
            try {
                $cpuRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_Processor -ErrorAction Stop
                $cpuInfo = foreach ($cpu in $cpuRaw) {
                    [PSCustomObject]@{
                        Prozessor          = $cpu.Name
                        Kerne              = $cpu.NumberOfCores
                        "Logische Proz."   = $cpu.NumberOfLogicalProcessors
                        "Max Takt (MHz)"   = $cpu.MaxClockSpeed
                        "Akt. Takt (MHz)"  = $cpu.CurrentClockSpeed
                        "L2 Cache (KB)"    = $cpu.L2CacheSize
                        "L3 Cache (KB)"    = $cpu.L3CacheSize
                        Sockel             = $cpu.SocketDesignation
                        Hyperthreading     = if ($cpu.NumberOfLogicalProcessors -gt $cpu.NumberOfCores) { "Ja" } else { "Nein" }
                        Virtualisierung    = if ($cpu.VirtualizationFirmwareEnabled) { "✅ Aktiviert" } else { "Unbekannt/Deaktiviert" }
                    }
                }
            }
            catch {
                Write-Log -Message "CPU-Abfrage fehlgeschlagen für ${server}: $_" -Level "ERROR"
            }

            # --- Physischer Speicher (DIMMs) ---
            $memInfo = $null
            try {
                $memRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_PhysicalMemory -ErrorAction Stop
                $memInfo = foreach ($m in $memRaw) {
                    [PSCustomObject]@{
                        Bank          = $m.BankLabel
                        Steckplatz    = $m.DeviceLocator
                        "Größe_GB"    = [math]::Round($m.Capacity / 1GB, 2)
                        "Takt_MHz"    = $m.Speed
                        Hersteller    = $m.Manufacturer
                        Teilenummer   = ($m.PartNumber -replace '\s+$', '')
                    }
                }
            }
            catch {
                Write-Log -Message "RAM-Modul-Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
            }

            # --- Logische Laufwerke ---
            $diskInfo = $null
            try {
                $diskRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop
                $diskInfo = foreach ($disk in $diskRaw) {
                    $belegtProzent = if ($disk.Size -gt 0) { [math]::Round((($disk.Size - $disk.FreeSpace) / $disk.Size) * 100, 1) } else { 0 }
                    $freiGB = [math]::Round($disk.FreeSpace / 1GB, 2)
                    [PSCustomObject]@{
                        Laufwerk    = $disk.DeviceID
                        Volumenname = $disk.VolumeName
                        "Gesamt_GB" = [math]::Round($disk.Size / 1GB, 2)
                        "Frei_GB"   = $freiGB
                        "Belegt_%"  = $belegtProzent
                        Dateisystem = $disk.FileSystem
                        Status      = if ($freiGB -lt $script:WarningDiskSpaceGB) { "⚠️ WENIG PLATZ!" } else { "✅ OK" }
                    }
                }
            }
            catch {
                Write-Log -Message "Festplatten-Abfrage fehlgeschlagen für ${server}: $_" -Level "ERROR"
            }

            # --- Physische Festplatten ---
            $physDiskInfo = $null
            try {
                $physDiskRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_DiskDrive -ErrorAction Stop
                $physDiskInfo = foreach ($pd in $physDiskRaw) {
                    [PSCustomObject]@{
                        Modell       = $pd.Model
                        "Größe_GB"   = [math]::Round($pd.Size / 1GB, 2)
                        InterfaceTyp = $pd.InterfaceType
                        MediaType    = $pd.MediaType
                        Partitionen  = $pd.Partitions
                        Seriennummer = ($pd.SerialNumber -replace '\s+', '')
                        Status       = $pd.Status
                    }
                }
            }
            catch {
                Write-Log -Message "Phys. Festplatten-Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
            }

            # --- Pagefile ---
            $pageFileInfo = $null
            try {
                $pfRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_PageFileSetting -ErrorAction SilentlyContinue
                if ($pfRaw) {
                    $pageFileInfo = foreach ($pf in $pfRaw) {
                        [PSCustomObject]@{
                            Pfad           = $pf.Name
                            "InitGröße_MB" = $pf.InitialSize
                            "MaxGröße_MB"  = $pf.MaximumSize
                        }
                    }
                }
                else {
                    $pageFileInfo = [PSCustomObject]@{
                        Pfad           = "Automatisch verwaltet"
                        "InitGröße_MB" = "Auto"
                        "MaxGröße_MB"  = "Auto"
                    }
                }
            }
            catch {
                Write-Log -Message "Pagefile-Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
            }

            # --- Netzwerk ---
            $nicInfo = $null
            try {
                $nicRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction Stop
                $nicInfo = foreach ($nic in $nicRaw) {
                    [PSCustomObject]@{
                        Adapter    = $nic.Description
                        IPAdresse  = ($nic.IPAddress -join ', ')
                        Subnetz    = ($nic.IPSubnet -join ', ')
                        Gateway    = ($nic.DefaultIPGateway -join ', ')
                        DNS_Server = ($nic.DNSServerSearchOrder -join ', ')
                        DHCP       = if ($nic.DHCPEnabled) { "Ja" } else { "Nein" }
                        MAC        = $nic.MACAddress
                    }
                }
            }
            catch {
                Write-Log -Message "Netzwerk-Abfrage fehlgeschlagen für ${server}: $_" -Level "ERROR"
            }

            # --- Hyper-V Dienste ---
            $hvServices = $null
            try {
                $servicesRaw = Get-CimInstance -CimSession $cimSession -ClassName Win32_Service -ErrorAction Stop |
                    Where-Object { $_.Name -match '^(vmms|vmcompute|nvspwmi|vmicvss|vmicheartbeat|vmicrdv|vmicguestinterface|vmickvpexchange|vmicshutdown|vmictimesync|hvhost|HvHost|storvsp)' } |
                    Sort-Object Name
                $hvServices = foreach ($svc in $servicesRaw) {
                    [PSCustomObject]@{
                        Dienstname  = $svc.Name
                        Anzeigename = $svc.DisplayName
                        Status      = $svc.State
                        Starttyp    = $svc.StartMode
                        Konto       = $svc.StartName
                        Warnung     = if ($svc.State -ne "Running" -and $svc.StartMode -eq "Auto") { "⚠️ GESTOPPT!" } else { "✅ OK" }
                    }
                }
            }
            catch {
                Write-Log -Message "Dienste-Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
            }

            if ($cimSession) {
                Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
            }

            # --- HTML zusammenbauen ---
            $serverHTML = "<h3 class='server-break'>Host: $server</h3>"
            $serverHTML += "<h4>System-Übersicht</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data @($sysInfo) -NoDataMessage "Nicht verfügbar")
            $serverHTML += "<h4>Betriebssystem</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data @($osInfo) -NoDataMessage "Nicht verfügbar")
            $serverHTML += "<h4>Prozessor(en)</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data $cpuInfo -NoDataMessage "Nicht verfügbar")
            $serverHTML += "<h4>Physischer Arbeitsspeicher (DIMMs)</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data $memInfo -NoDataMessage "Nicht verfügbar")
            $serverHTML += "<h4>Logische Laufwerke</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data $diskInfo -NoDataMessage "Nicht verfügbar")
            $serverHTML += "<h4>Physische Festplatten</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data $physDiskInfo -NoDataMessage "Nicht verfügbar")
            $serverHTML += "<h4>Pagefile</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data @($pageFileInfo) -NoDataMessage "Nicht verfügbar")
            $serverHTML += "<h4>Netzwerkkonfiguration (IP)</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data $nicInfo -NoDataMessage "Nicht verfügbar")
            $serverHTML += "<h4>Hyper-V Dienste</h4>"
            $serverHTML += (ConvertTo-HTMLTable -Data $hvServices -NoDataMessage "Nicht verfügbar")

            $allHardwareHTML += $serverHTML
        }
        catch {
            Write-Log -Message "Allgemeiner Fehler bei Hardware-Info für ${server}: $_" -Level "ERROR"
            $allHardwareHTML += "<h3 class='server-break'>Host: $server</h3><p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "Hardware-Informationen & Host-Details" -Category "Hardware & OS" -Content $allHardwareHTML
}

# ---------------------------------------------------------------
# 2. VIRTUALISIERUNGS-UNTERSTÜTZUNG (SLAT, VT-x/AMD-V, DEP)
# ---------------------------------------------------------------
function Get-CpuVirtualizationSupportInfo {
    <#
    .SYNOPSIS
        Prüft die Hardware-Voraussetzungen für Hyper-V (SLAT, VM Monitor Mode,
        Virtualisierung in der Firmware, DEP) sowie den Hypervisor-Status.
    #>
    Write-Log -Message "=== Sammle Virtualisierungs-Unterstützung ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $cimSession = New-ServerCimSession -ComputerName $server
            if (-not $cimSession) {
                $html += "<p class='error'>Keine Verbindung möglich.</p>"
                continue
            }

            $cpu = Get-CimInstance -CimSession $cimSession -ClassName Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
            $cs  = Get-CimInstance -CimSession $cimSession -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
            $os  = Get-CimInstance -CimSession $cimSession -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue

            $result = [PSCustomObject]@{
                "SLAT (EPT/NPT)"              = if ($cpu.SecondLevelAddressTranslationExtensions) { "✅ Unterstützt" } else { "❌ Nicht unterstützt/unbekannt" }
                "VM Monitor Mode Extensions"  = if ($cpu.VMMonitorModeExtensions) { "✅ Unterstützt" } else { "❌ Nicht unterstützt/unbekannt" }
                "Virtualisierung in Firmware" = if ($cpu.VirtualizationFirmwareEnabled) { "✅ Aktiviert" } else { "⚠️ Deaktiviert/unbekannt" }
                "DEP (NX/XD) verfügbar"       = if ($os -and $os.DataExecutionPrevention_Available) { "✅ Ja" } else { "⚠️ Nein/unbekannt" }
                "DEP für Treiber aktiv"       = if ($os -and $os.DataExecutionPrevention_Drivers) { "✅ Ja" } else { "⚠️ Nein/unbekannt" }
                "Hypervisor aktiv"            = if ($cs.HypervisorPresent) { "✅ Ja" } else { "⚠️ Nein" }
                "Prozessorarchitektur"        = switch ($cpu.Architecture) { 0 { "x86" } 5 { "ARM" } 9 { "x64" } 12 { "ARM64" } default { "$($cpu.Architecture)" } }
            }
            $html += (ConvertTo-HTMLTable -Data @($result))

            # --- Device Guard / VBS (kann Nested Virtualization beeinflussen) ---
            try {
                $dg = Get-CimInstance -CimSession $cimSession -Namespace "root\Microsoft\Windows\DeviceGuard" -ClassName Win32_DeviceGuard -ErrorAction Stop
                $dgInfo = [PSCustomObject]@{
                    "VBS Status"                = switch ($dg.VirtualizationBasedSecurityStatus) { 0 { "Nicht aktiviert" } 1 { "Aktiviert, nicht gestartet" } 2 { "Aktiviert und gestartet" } default { "Unbekannt" } }
                    "Konfigurierte Services"    = ConvertTo-DisplayValue -Value ($dg.SecurityServicesConfigured | ForEach-Object { switch ($_) { 1 { "Credential Guard" } 2 { "HVCI" } 3 { "System Guard" } default { "$_" } } })
                    "Laufende Services"         = ConvertTo-DisplayValue -Value ($dg.SecurityServicesRunning | ForEach-Object { switch ($_) { 1 { "Credential Guard" } 2 { "HVCI" } 3 { "System Guard" } default { "$_" } } })
                    "Verfügbare Sicherheits-EIG" = ConvertTo-DisplayValue -Value $dg.AvailableSecurityProperties
                }
                $html += "<h4>Virtualization Based Security (VBS)</h4>"
                $html += (ConvertTo-HTMLTable -Data @($dgInfo))
            }
            catch {
                $html += "<h4>Virtualization Based Security (VBS)</h4><p class='no-data'>Keine VBS-Informationen verfügbar.</p>"
            }

            Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
        }
        catch {
            Write-Log -Message "Virtualisierungs-Check fehlgeschlagen für ${server}: $_" -Level "WARNING"
            $html += "<p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "Virtualisierungs-Unterstützung (SLAT / VT-x / VBS)" -Category "Hardware & OS" -Content $html
}

# ---------------------------------------------------------------
# 3. NUMA-TOPOLOGIE
# ---------------------------------------------------------------
function Get-NumaTopologyInfo {
    <#
    .SYNOPSIS
        Dokumentiert die NUMA-Topologie des Hosts und die NUMA-Spanning-Einstellung.
    #>
    Write-Log -Message "=== Sammle NUMA-Topologie ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"

        # --- NUMA Nodes ---
        try {
            $numaNodes = Get-VMHostNumaNode -ComputerName $server -ErrorAction Stop
            $numaData = foreach ($n in $numaNodes) {
                [PSCustomObject]@{
                    "NUMA Node ID"          = $n.NodeId
                    "Prozessoren (Anzahl)"  = @($n.ProcessorsAvailability).Count
                    "Speicher gesamt (MB)"  = $n.MemoryTotal
                    "Speicher verfügbar (MB)" = $n.MemoryAvailable
                    "Speicher belegt (%)"   = if ($n.MemoryTotal -gt 0) { [math]::Round((1 - ($n.MemoryAvailable / $n.MemoryTotal)) * 100, 1) } else { 0 }
                    "Prozessor-Auslastung"  = ConvertTo-DisplayValue -Value $n.ProcessorsAvailability
                }
            }
            $html += "<h4>NUMA Nodes</h4>"
            $html += (ConvertTo-HTMLTable -Data $numaData -NoDataMessage "Keine NUMA-Knoten gefunden.")
        }
        catch {
            Write-Log -Message "NUMA-Node-Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
            $html += "<h4>NUMA Nodes</h4><p class='no-data'>Nicht verfügbar: $($_.Exception.Message)</p>"
        }

        # --- NUMA Spanning ---
        try {
            $vmHost = Get-VMHost -ComputerName $server -ErrorAction Stop
            $spanning = [PSCustomObject]@{
                "NUMA Spanning aktiviert" = if ($vmHost.NumaSpanningEnabled) { "Ja" } else { "Nein" }
                "Bewertung"               = if ($vmHost.NumaSpanningEnabled) { "ℹ️ Standard - erlaubt VMs über NUMA-Grenzen (bessere Dichte, evtl. geringere Performance)" } else { "✅ Deaktiviert - beste NUMA-Performance, VM-Größe auf Node-Größe begrenzt" }
                "Log. Prozessoren gesamt" = $vmHost.LogicalProcessorCount
                "Speicherkapazität"       = Format-ByteSize -Bytes $vmHost.MemoryCapacity
            }
            $html += "<h4>NUMA Spanning</h4>"
            $html += (ConvertTo-HTMLTable -Data @($spanning))
        }
        catch {
            $html += "<h4>NUMA Spanning</h4><p class='no-data'>Nicht verfügbar.</p>"
        }

        # --- VM NUMA Konfiguration ---
        try {
            $vms = Get-VM -ComputerName $server -ErrorAction Stop
            $vmNuma = foreach ($vm in $vms) {
                try {
                    $numa = Get-VMProcessor -VMName $vm.Name -ComputerName $server -ErrorAction Stop
                    [PSCustomObject]@{
                        VM                          = $vm.Name
                        "Max. Prozessoren/NUMA-Node" = $numa.MaximumCountPerNumaNode
                        "Max. Nodes/Socket"          = $numa.MaximumCountPerNumaSocket
                        "vCPUs"                      = $numa.Count
                    }
                }
                catch { $null }
            }
            $html += "<h4>NUMA-Konfiguration der VMs</h4>"
            $html += (ConvertTo-HTMLTable -Data $vmNuma -NoDataMessage "Keine VMs gefunden.")
        }
        catch {
            $html += "<h4>NUMA-Konfiguration der VMs</h4><p class='no-data'>Nicht verfügbar.</p>"
        }
    }

    New-HTMLSection -Title "NUMA-Topologie & Spanning" -Category "Hardware & OS" -Content $html
}

# ---------------------------------------------------------------
# 4. INSTALLIERTE SOFTWARE
# ---------------------------------------------------------------
function Get-InstalledSoftwareInfo {
    <#
    .SYNOPSIS
        Liest installierte Software über die Remote-Registry aus (32- und 64-Bit).
    #>
    Write-Log -Message "=== Sammle installierte Software ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $softwareList = [System.Collections.ArrayList]::new()
            $regPaths = @(
                "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            )

            foreach ($regPath in $regPaths) {
                $subKeys = Get-RemoteRegistrySubKeyNames -ComputerName $server -RegistryPath $regPath
                foreach ($key in $subKeys) {
                    $fullPath    = "$regPath\$key"
                    $displayName = Get-RemoteRegistryValue -ComputerName $server -RegistryPath $fullPath -ValueName "DisplayName"
                    if (-not $displayName) { continue }

                    $version     = Get-RemoteRegistryValue -ComputerName $server -RegistryPath $fullPath -ValueName "DisplayVersion"
                    $publisher   = Get-RemoteRegistryValue -ComputerName $server -RegistryPath $fullPath -ValueName "Publisher"
                    $installDate = Get-RemoteRegistryValue -ComputerName $server -RegistryPath $fullPath -ValueName "InstallDate"

                    $formattedDate = "-"
                    if ($installDate -and $installDate -match '^\d{8}$') {
                        try { $formattedDate = [datetime]::ParseExact($installDate, "yyyyMMdd", $null).ToString("dd.MM.yyyy") } catch { $formattedDate = $installDate }
                    }

                    [void]$softwareList.Add([PSCustomObject]@{
                        Software      = $displayName
                        Version       = ConvertTo-DisplayValue -Value $version
                        Hersteller    = ConvertTo-DisplayValue -Value $publisher
                        Installation  = $formattedDate
                        Architektur   = if ($regPath -match "WOW6432Node") { "32-Bit" } else { "64-Bit" }
                    })
                }
            }

            $sorted = $softwareList | Sort-Object Software -Unique
            $html += "<p><strong>Gefundene Programme:</strong> $($sorted.Count)</p>"
            $html += (ConvertTo-HTMLTable -Data $sorted -NoDataMessage "Keine Software gefunden.")
        }
        catch {
            Write-Log -Message "Software-Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
            $html += "<p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "Installierte Software" -Category "Hardware & OS" -Content $html
}

# ---------------------------------------------------------------
# 5. WINDOWS FEATURES & ROLLEN
# ---------------------------------------------------------------
function Get-WindowsFeaturesInfo {
    <#
    .SYNOPSIS
        Dokumentiert installierte Windows-Rollen und Features (Fokus Hyper-V).
    #>
    Write-Log -Message "=== Sammle Windows Features & Rollen ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"

        $features = Invoke-RemoteCommand -ComputerName $server -ScriptBlock {
            if (Get-Command Get-WindowsFeature -ErrorAction SilentlyContinue) {
                Get-WindowsFeature | Where-Object { $_.Installed } |
                    Select-Object Name, DisplayName, FeatureType, InstallState, @{N='Path';E={$_.Path}}
            }
        }

        if ($features) {
            $featureData = $features | Sort-Object Name | ForEach-Object {
                [PSCustomObject]@{
                    Name        = $_.Name
                    Anzeigename = $_.DisplayName
                    Typ         = $_.FeatureType
                    Status      = $_.InstallState
                }
            }
            $html += "<p><strong>Installierte Features/Rollen:</strong> $(@($featureData).Count)</p>"

            # Hyper-V relevante Features hervorheben
            $hvFeatures = $featureData | Where-Object { $_.Name -match 'Hyper-V|RSAT-Hyper|Failover|Multipath|Data-Center-Bridging|FS-SMBBW' }
            $html += "<h4>Hyper-V relevante Rollen &amp; Features</h4>"
            $html += (ConvertTo-HTMLTable -Data $hvFeatures -NoDataMessage "Keine Hyper-V Features gefunden.")

            $html += "<h4>Alle installierten Features</h4>"
            $html += (ConvertTo-HTMLTable -Data $featureData -NoDataMessage "Keine Features gefunden.")
        }
        else {
            $html += "<p class='no-data'>Windows Features konnten nicht abgefragt werden (WinRM erforderlich).</p>"
        }
    }

    New-HTMLSection -Title "Windows Features &amp; Rollen" -Category "Hardware & OS" -Content $html
}

# ---------------------------------------------------------------
# 6. POWER PLAN
# ---------------------------------------------------------------
function Get-PowerPlanInfo {
    <#
    .SYNOPSIS
        Prüft den aktiven Energiesparplan. Best Practice für Hyper-V: "Höchstleistung".
    #>
    Write-Log -Message "=== Sammle Power Plan Konfiguration ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Best Practice</h3><p>Hyper-V Hosts sollten den Energiesparplan <strong>Höchstleistung</strong> verwenden, "
    $html += "damit die CPU nicht heruntergetaktet wird und die Latenz für VMs niedrig bleibt.</p></div>"

    $data = foreach ($server in $HyperVServers) {
        try {
            $cimSession = New-ServerCimSession -ComputerName $server
            if (-not $cimSession) {
                [PSCustomObject]@{ Host = $server; Energiesparplan = "Nicht ermittelbar"; GUID = "-"; Bewertung = "❌ Keine Verbindung" }
                continue
            }

            $plan = Get-CimInstance -CimSession $cimSession -Namespace "root\cimv2\power" -ClassName Win32_PowerPlan -Filter "IsActive=True" -ErrorAction SilentlyContinue
            Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue

            if ($plan) {
                $isHigh = $plan.ElementName -match "Höchstleistung|High performance|Ultimate"
                [PSCustomObject]@{
                    Host            = $server
                    Energiesparplan = $plan.ElementName
                    GUID            = ($plan.InstanceID -replace '.*\\\{|\}.*', '')
                    Bewertung       = if ($isHigh) { "✅ OK" } else { "⚠️ Nicht Höchstleistung!" }
                }
            }
            else {
                [PSCustomObject]@{ Host = $server; Energiesparplan = "Unbekannt"; GUID = "-"; Bewertung = "⚠️ Nicht ermittelbar" }
            }
        }
        catch {
            Write-Log -Message "Power Plan Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
            [PSCustomObject]@{ Host = $server; Energiesparplan = "Fehler"; GUID = "-"; Bewertung = "❌ $($_.Exception.Message)" }
        }
    }

    $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine Daten verfügbar.")
    New-HTMLSection -Title "Power Plan &amp; Performance" -Category "Hardware & OS" -Content $html
}

# ---------------------------------------------------------------
# 7. AUSSTEHENDE NEUSTARTS
# ---------------------------------------------------------------
function Get-PendingRebootInfo {
    <#
    .SYNOPSIS
        Prüft mehrere Registry-Indikatoren auf ausstehende Neustarts.
    #>
    Write-Log -Message "=== Prüfe ausstehende Neustarts ===" -Level "INFO"

    $data = foreach ($server in $HyperVServers) {
        try {
            $cbs = Get-RemoteRegistrySubKeyNames -ComputerName $server -RegistryPath "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing" |
                Where-Object { $_ -eq "RebootPending" }
            $wu = Get-RemoteRegistrySubKeyNames -ComputerName $server -RegistryPath "SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update" |
                Where-Object { $_ -eq "RebootRequired" }
            $pfr = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SYSTEM\CurrentControlSet\Control\Session Manager" -ValueName "PendingFileRenameOperations"
            $srv = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SOFTWARE\Microsoft\ServerManager" -ValueName "CurrentRebootAttempts"
            $cn  = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SYSTEM\CurrentControlSet\Services\Netlogon" -ValueName "JoinDomain"

            $pending = @()
            if ($cbs) { $pending += "Component Based Servicing" }
            if ($wu)  { $pending += "Windows Update" }
            if ($pfr) { $pending += "Pending File Rename" }
            if ($srv) { $pending += "Server Manager" }
            if ($cn)  { $pending += "Domain Join" }

            [PSCustomObject]@{
                Host        = $server
                "CBS"       = if ($cbs) { "⚠️ Ja" } else { "✅ Nein" }
                "Windows Update" = if ($wu) { "⚠️ Ja" } else { "✅ Nein" }
                "File Rename"    = if ($pfr) { "⚠️ Ja" } else { "✅ Nein" }
                "Server Manager" = if ($srv) { "⚠️ Ja" } else { "✅ Nein" }
                Gesamtstatus     = if ($pending.Count -gt 0) { "⚠️ NEUSTART ERFORDERLICH ($($pending -join ', '))" } else { "✅ Kein Neustart erforderlich" }
            }
        }
        catch {
            [PSCustomObject]@{
                Host = $server; "CBS" = "?"; "Windows Update" = "?"; "File Rename" = "?"; "Server Manager" = "?"
                Gesamtstatus = "❌ Fehler: $($_.Exception.Message)"
            }
        }
    }

    $html = "<div class='summary-box'><h3>Hinweis</h3><p>Ein ausstehender Neustart auf einem Hyper-V Host sollte nur im Wartungsfenster "
    $html += "und nach Live-Migration bzw. sauberem Herunterfahren aller VMs durchgeführt werden.</p></div>"
    $html += (ConvertTo-HTMLTable -Data $data)

    New-HTMLSection -Title "Ausstehende Neustarts" -Category "Hardware & OS" -Content $html
}

# ---------------------------------------------------------------
# 8. PATCH-INFORMATIONEN
# ---------------------------------------------------------------
function Get-PatchInformation {
    <#
    .SYNOPSIS
        Listet die zuletzt installierten Windows-Updates je Host.
    #>
    Write-Log -Message "=== Sammle Patch-Informationen ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $cimSession = New-ServerCimSession -ComputerName $server
            if (-not $cimSession) {
                $html += "<p class='error'>Keine Verbindung möglich.</p>"
                continue
            }

            $hotfixes = Get-CimInstance -CimSession $cimSession -ClassName Win32_QuickFixEngineering -ErrorAction SilentlyContinue |
                Sort-Object InstalledOn -Descending | Select-Object -First 40

            $data = foreach ($hf in $hotfixes) {
                [PSCustomObject]@{
                    HotfixID       = $hf.HotFixID
                    Beschreibung   = $hf.Description
                    InstalliertAm  = if ($hf.InstalledOn) { $hf.InstalledOn.ToString("dd.MM.yyyy") } else { "Unbekannt" }
                    InstalliertVon = $hf.InstalledBy
                }
            }

            $lastPatch = ($hotfixes | Select-Object -First 1).InstalledOn
            if ($lastPatch) {
                $days = [math]::Round(((Get-Date) - $lastPatch).TotalDays, 0)
                $status = if ($days -gt 60) { "⚠️ Letztes Update vor $days Tagen!" } else { "✅ Letztes Update vor $days Tagen" }
                $html += "<p><strong>Patch-Status:</strong> $status</p>"
            }

            $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine Updates gefunden.")
            Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
        }
        catch {
            Write-Log -Message "Patch-Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
            $html += "<p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "Windows Updates &amp; Patch-Stand" -Category "Hardware & OS" -Content $html
}

# ---------------------------------------------------------------
# 9. SPEICHERPLATZ (VM-RELEVANTE PFADE)
# ---------------------------------------------------------------
function Get-DiskSpaceInfo {
    <#
    .SYNOPSIS
        Analysiert den freien Speicherplatz mit Fokus auf VM- und VHD-Pfade.
    #>
    Write-Log -Message "=== Sammle Speicherplatz-Informationen ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $cimSession = New-ServerCimSession -ComputerName $server
            if ($cimSession) {
                $volumes = Get-CimInstance -CimSession $cimSession -ClassName Win32_Volume -Filter "DriveType=3" -ErrorAction SilentlyContinue
                $volData = foreach ($v in $volumes) {
                    $freiGB = [math]::Round($v.FreeSpace / 1GB, 2)
                    $gesamtGB = [math]::Round($v.Capacity / 1GB, 2)
                    [PSCustomObject]@{
                        Laufwerk      = ConvertTo-DisplayValue -Value $v.DriveLetter
                        Bezeichnung   = ConvertTo-DisplayValue -Value $v.Label
                        Pfad          = $v.Name
                        Dateisystem   = $v.FileSystem
                        "Blockgröße"  = ConvertTo-DisplayValue -Value $v.BlockSize
                        "Gesamt_GB"   = $gesamtGB
                        "Frei_GB"     = $freiGB
                        "Frei_%"      = if ($gesamtGB -gt 0) { [math]::Round(($freiGB / $gesamtGB) * 100, 1) } else { 0 }
                        Status        = if ($freiGB -lt $script:WarningDiskSpaceGB) { "⚠️ WENIG PLATZ" } else { "✅ OK" }
                    }
                }
                $html += "<h4>Volumes</h4>"
                $html += (ConvertTo-HTMLTable -Data $volData -NoDataMessage "Keine Volumes gefunden.")
                Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
            }

            # --- Speicherbedarf der VMs pro Pfad ---
            try {
                $vms = Get-VM -ComputerName $server -ErrorAction Stop
                $pathData = foreach ($vm in $vms) {
                    $vhds = @()
                    try { $vhds = @(Get-VMHardDiskDrive -VMName $vm.Name -ComputerName $server -ErrorAction SilentlyContinue) } catch { $vhds = @() }
                    $totalSize = 0
                    foreach ($vhd in $vhds) {
                        try {
                            $v = Get-VHD -Path $vhd.Path -ComputerName $server -ErrorAction SilentlyContinue
                            if ($v) { $totalSize += $v.FileSize }
                        }
                        catch { }
                    }
                    [PSCustomObject]@{
                        VM                    = $vm.Name
                        Konfigurationspfad    = $vm.ConfigurationLocation
                        Checkpointpfad        = $vm.SnapshotFileLocation
                        "Smart Paging Pfad"   = $vm.SmartPagingFilePath
                        "Anzahl VHDs"         = $vhds.Count
                        "Belegter Speicher"   = Format-ByteSize -Bytes $totalSize
                    }
                }
                $html += "<h4>Speicherpfade der virtuellen Maschinen</h4>"
                $html += (ConvertTo-HTMLTable -Data $pathData -NoDataMessage "Keine VMs gefunden.")
            }
            catch {
                $html += "<h4>Speicherpfade der virtuellen Maschinen</h4><p class='no-data'>Nicht verfügbar.</p>"
            }
        }
        catch {
            Write-Log -Message "Speicherplatz-Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
            $html += "<p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "Speicherplatz &amp; VM-Pfade" -Category "Hardware & OS" -Content $html
}

# ---------------------------------------------------------------
# 10. EVENT LOGS (HYPER-V)
# ---------------------------------------------------------------
function Get-EventLogInfo {
    <#
    .SYNOPSIS
        Liest kritische Ereignisse und Fehler der letzten 7 Tage aus System-
        und Hyper-V-spezifischen Ereignisprotokollen.
    #>
    Write-Log -Message "=== Sammle Event Logs (7 Tage) ===" -Level "INFO"

    $html = ""
    $startTime = (Get-Date).AddDays(-7)

    $logNames = @(
        "System",
        "Microsoft-Windows-Hyper-V-VMMS-Admin",
        "Microsoft-Windows-Hyper-V-Worker-Admin",
        "Microsoft-Windows-Hyper-V-Compute-Admin",
        "Microsoft-Windows-Hyper-V-VmSwitch-Operational",
        "Microsoft-Windows-Hyper-V-StorageVSP-Admin",
        "Microsoft-Windows-Hyper-V-High-Availability-Admin",
        "Microsoft-Windows-Hyper-V-Config-Admin",
        "Microsoft-Windows-Hyper-V-Hypervisor-Admin",
        "Microsoft-Windows-FailoverClustering/Operational"
    )

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"

        foreach ($logName in $logNames) {
            try {
                $events = Get-WinEvent -ComputerName $server -FilterHashtable @{
                    LogName   = $logName
                    Level     = 1, 2
                    StartTime = $startTime
                } -MaxEvents 40 -ErrorAction Stop

                if ($events) {
                    $data = foreach ($e in $events) {
                        [PSCustomObject]@{
                            Zeitpunkt = $e.TimeCreated.ToString("dd.MM.yyyy HH:mm:ss")
                            Level     = switch ($e.Level) { 1 { "Kritisch" } 2 { "Fehler" } default { "$($e.Level)" } }
                            EventID   = $e.Id
                            Quelle    = $e.ProviderName
                            Meldung   = if ($e.Message) { ($e.Message -split "`n")[0].Substring(0, [Math]::Min(200, ($e.Message -split "`n")[0].Length)) } else { "-" }
                        }
                    }
                    $html += "<h4>$logName ($(@($data).Count) Ereignisse)</h4>"
                    $html += (ConvertTo-HTMLTable -Data $data)
                }
            }
            catch {
                # Log existiert nicht oder keine Ereignisse - stillschweigend überspringen
                continue
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($html)) {
        $html = "<p class='no-data'>Keine kritischen Ereignisse in den letzten 7 Tagen gefunden.</p>"
    }

    New-HTMLSection -Title "Event Logs (Fehler/Kritisch, 7 Tage)" -Category "Hardware & OS" -Content $html
}

# ---------------------------------------------------------------
# 11. ZEITSYNCHRONISIERUNG
# ---------------------------------------------------------------
function Get-TimeSyncInfo {
    <#
    .SYNOPSIS
        Dokumentiert die Zeitsynchronisierung des Hosts (w32tm) - wichtig, weil
        VMs die Hostzeit über den Integrationsdienst "Zeitsynchronisierung" beziehen.
    #>
    Write-Log -Message "=== Sammle Zeitsynchronisierungs-Konfiguration ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Best Practice</h3><p>Der Hyper-V Host sollte seine Zeit von einem zuverlässigen NTP-Server (bzw. dem PDC-Emulator) beziehen. "
    $html += "In Domänen-VMs sollte die Zeitsynchronisierung über den Integrationsdienst mit Vorsicht verwendet werden (Domain Controller als VM: Integrationsdienst deaktivieren).</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"

        $timeInfo = Invoke-RemoteCommand -ComputerName $server -ScriptBlock {
            $status = & w32tm /query /status 2>&1
            $config = & w32tm /query /configuration 2>&1
            $source = & w32tm /query /source 2>&1
            [PSCustomObject]@{
                Status = ($status | Out-String)
                Config = ($config | Out-String)
                Source = ($source | Out-String).Trim()
                Zeit   = (Get-Date).ToString("dd.MM.yyyy HH:mm:ss")
                Zone   = (Get-TimeZone).DisplayName
            }
        }

        if ($timeInfo) {
            $summary = [PSCustomObject]@{
                "Aktuelle Zeit"  = $timeInfo.Zeit
                "Zeitzone"       = $timeInfo.Zone
                "Zeitquelle"     = $timeInfo.Source
            }
            $html += (ConvertTo-HTMLTable -Data @($summary))

            $statusLines = $timeInfo.Status -split "`r?`n" | Where-Object { $_ -match ':' }
            $statusData = foreach ($line in $statusLines) {
                $parts = $line -split ':', 2
                [PSCustomObject]@{
                    Parameter = $parts[0].Trim()
                    Wert      = if ($parts.Count -gt 1) { $parts[1].Trim() } else { "-" }
                }
            }
            $html += "<h4>w32tm Status</h4>"
            $html += (ConvertTo-HTMLTable -Data $statusData -NoDataMessage "Nicht verfügbar.")
        }
        else {
            $html += "<p class='no-data'>Zeitkonfiguration konnte nicht abgefragt werden (WinRM erforderlich).</p>"
        }
    }

    New-HTMLSection -Title "Zeitsynchronisierung (w32tm)" -Category "Hardware & OS" -Content $html
}

# ---------------------------------------------------------------
# 12. LIZENZIERUNG
# ---------------------------------------------------------------
function Get-LicensingInfo {
    <#
    .SYNOPSIS
        Dokumentiert die Windows-Lizenzierung des Hosts (relevant für VM-Lizenzrechte:
        Standard = 2 VMs, Datacenter = unbegrenzt).
    #>
    Write-Log -Message "=== Sammle Lizenzierungs-Informationen ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Hinweis zur Virtualisierungs-Lizenzierung</h3>"
    $html += "<p><strong>Windows Server Standard:</strong> Rechte für 2 virtualisierte OS-Umgebungen je vollständig lizenziertem Host.<br>"
    $html += "<strong>Windows Server Datacenter:</strong> Unbegrenzte Anzahl virtualisierter OS-Umgebungen.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $cimSession = New-ServerCimSession -ComputerName $server
            if (-not $cimSession) {
                $html += "<p class='error'>Keine Verbindung möglich.</p>"
                continue
            }

            $os = Get-CimInstance -CimSession $cimSession -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
            $lic = Get-CimInstance -CimSession $cimSession -ClassName SoftwareLicensingProduct -Filter "PartialProductKey IS NOT NULL AND Name LIKE 'Windows%'" -ErrorAction SilentlyContinue

            $vmCount = 0
            try { $vmCount = @(Get-VM -ComputerName $server -ErrorAction SilentlyContinue).Count } catch { $vmCount = 0 }

            $data = foreach ($l in $lic) {
                [PSCustomObject]@{
                    Produkt          = $l.Name
                    Beschreibung     = $l.Description
                    "Teil-Key"       = $l.PartialProductKey
                    Lizenzstatus     = switch ($l.LicenseStatus) { 0 { "Nicht lizenziert" } 1 { "✅ Lizenziert" } 2 { "OOB Grace" } 3 { "OOT Grace" } 4 { "Non-Genuine Grace" } 5 { "⚠️ Benachrichtigung" } 6 { "Extended Grace" } default { "Unbekannt" } }
                    "KMS Server"     = ConvertTo-DisplayValue -Value $l.KeyManagementServiceMachine
                    "Verbleibend (Tage)" = if ($l.GracePeriodRemaining) { [math]::Round($l.GracePeriodRemaining / 1440, 0) } else { "-" }
                }
            }

            $edition = [PSCustomObject]@{
                "Betriebssystem"        = if ($os) { $os.Caption } else { "-" }
                "Edition"               = if ($os) { $os.OperatingSystemSKU } else { "-" }
                "Anzahl VMs auf Host"   = $vmCount
                "Lizenzhinweis"         = if ($os -and $os.Caption -match "Datacenter") { "✅ Datacenter - unbegrenzte VM-Rechte" } elseif ($os -and $os.Caption -match "Standard") { if ($vmCount -gt 2) { "⚠️ Standard-Edition mit $vmCount VMs - zusätzliche Lizenzen prüfen!" } else { "✅ Standard-Edition, $vmCount VM(s)" } } else { "Manuell prüfen" }
            }

            $html += "<h4>Edition &amp; VM-Rechte</h4>"
            $html += (ConvertTo-HTMLTable -Data @($edition))
            $html += "<h4>Aktivierungsstatus</h4>"
            $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine Lizenzinformationen verfügbar.")

            Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
        }
        catch {
            Write-Log -Message "Lizenz-Abfrage fehlgeschlagen für ${server}: $_" -Level "WARNING"
            $html += "<p class='error'>Fehler: $_</p>"
        }
    }

    New-HTMLSection -Title "Lizenzierung &amp; Aktivierung" -Category "Hardware & OS" -Content $html
}

# ---------------------------------------------------------------
# 13. HYPER-V HOST-KONFIGURATION (Get-VMHost)
# ---------------------------------------------------------------
function Get-VMHostConfigurationInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle Konfigurationsparameter des Hyper-V Hosts (Get-VMHost).
    #>
    Write-Log -Message "=== Sammle Hyper-V Host-Konfiguration ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $vmHost = Get-VMHost -ComputerName $server -ErrorAction Stop

            $general = [PSCustomObject]@{
                "Hostname"                        = $vmHost.Name
                "Betriebssystem"                  = $script:HyperVEdition
                "Logische Prozessoren"            = $vmHost.LogicalProcessorCount
                "Speicherkapazität"               = Format-ByteSize -Bytes $vmHost.MemoryCapacity
                "NUMA Spanning aktiviert"         = ConvertTo-DisplayValue -Value $vmHost.NumaSpanningEnabled
                "IOV (SR-IOV) Unterstützung"      = ConvertTo-DisplayValue -Value $vmHost.IovSupport
                "IOV Support Grund"               = ConvertTo-DisplayValue -Value $vmHost.IovSupportReasons
                "Enhanced Session Mode (Host)"    = ConvertTo-DisplayValue -Value $vmHost.EnableEnhancedSessionMode
                "VM Migration aktiviert"          = ConvertTo-DisplayValue -Value $vmHost.VirtualMachineMigrationEnabled
                "Resource Metering Intervall"     = ConvertTo-DisplayValue -Value $vmHost.ResourceMeteringSaveInterval
                "Fibre Channel WWNN"              = ConvertTo-DisplayValue -Value $vmHost.FibreChannelWwnn
                "Fibre Channel WWPN (Min)"        = ConvertTo-DisplayValue -Value $vmHost.FibreChannelWwpnMinimum
                "Fibre Channel WWPN (Max)"        = ConvertTo-DisplayValue -Value $vmHost.FibreChannelWwpnMaximum
                "MAC-Adressbereich (Min)"         = ConvertTo-DisplayValue -Value $vmHost.MacAddressMinimum
                "MAC-Adressbereich (Max)"         = ConvertTo-DisplayValue -Value $vmHost.MacAddressMaximum
            }

            $html += "<h4>Allgemeine Host-Einstellungen</h4>"
            $html += (ConvertTo-HTMLTable -Data @($general))

            # --- Alle Eigenschaften als Rohdaten ---
            $allProps = foreach ($p in ($vmHost.PSObject.Properties | Sort-Object Name)) {
                [PSCustomObject]@{
                    Parameter = $p.Name
                    Wert      = ConvertTo-DisplayValue -Value $p.Value
                }
            }
            $html += "<h4>Vollständige Parameterliste (Get-VMHost)</h4>"
            $html += (ConvertTo-HTMLTable -Data $allProps)
        }
        catch {
            Write-Log -Message "Get-VMHost fehlgeschlagen für ${server}: $_" -Level "ERROR"
            $html += "<p class='error'>Hyper-V Host-Konfiguration konnte nicht gelesen werden: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "Hyper-V Host-Konfiguration" -Category "Hyper-V Host" -Content $html
}

# ---------------------------------------------------------------
# 14. STANDARDPFADE
# ---------------------------------------------------------------
function Get-VMHostPathsInfo {
    <#
    .SYNOPSIS
        Dokumentiert die Standardpfade für virtuelle Festplatten und VM-Konfigurationen
        inkl. Prüfung des verfügbaren Speicherplatzes.
    #>
    Write-Log -Message "=== Sammle Standardpfade des Hosts ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Best Practice</h3><p>VHD- und VM-Konfigurationspfade sollten nicht auf dem Systemlaufwerk (C:) liegen. "
    $html += "Für Cluster sollten CSV-Pfade (C:\ClusterStorage\...) verwendet werden.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $vmHost = Get-VMHost -ComputerName $server -ErrorAction Stop

            $paths = @(
                [PSCustomObject]@{ Typ = "Virtuelle Festplatten (VHD/VHDX)"; Pfad = $vmHost.VirtualHardDiskPath }
                [PSCustomObject]@{ Typ = "VM-Konfigurationsdateien";        Pfad = $vmHost.VirtualMachinePath }
            )

            $pathData = foreach ($p in $paths) {
                $freeGB = "-"
                $warn   = "-"
                try {
                    $drive = ($p.Pfad -split ':')[0]
                    if ($drive -and $p.Pfad -notmatch '^\\\\') {
                        $cimSession = New-ServerCimSession -ComputerName $server
                        if ($cimSession) {
                            $ld = Get-CimInstance -CimSession $cimSession -ClassName Win32_LogicalDisk -Filter "DeviceID='${drive}:'" -ErrorAction SilentlyContinue
                            if ($ld) { $freeGB = [math]::Round($ld.FreeSpace / 1GB, 2) }
                            Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
                        }
                    }
                    $warn = if ($p.Pfad -match '^C:\\') { "⚠️ Liegt auf dem Systemlaufwerk" } else { "✅ OK" }
                }
                catch { }

                [PSCustomObject]@{
                    Typ         = $p.Typ
                    Pfad        = $p.Pfad
                    "Frei_GB"   = $freeGB
                    Bewertung   = $warn
                }
            }

            $html += "<h4>Standardpfade</h4>"
            $html += (ConvertTo-HTMLTable -Data $pathData)

            # --- Abweichende Pfade der VMs ---
            try {
                $vms = Get-VM -ComputerName $server -ErrorAction Stop
                $deviating = foreach ($vm in $vms) {
                    if ($vm.Path -and $vmHost.VirtualMachinePath -and ($vm.Path -notlike "$($vmHost.VirtualMachinePath)*")) {
                        [PSCustomObject]@{
                            VM                  = $vm.Name
                            "VM-Pfad"           = $vm.Path
                            "Konfiguration"     = $vm.ConfigurationLocation
                            "Checkpoints"       = $vm.SnapshotFileLocation
                            "Smart Paging"      = $vm.SmartPagingFilePath
                        }
                    }
                }
                $html += "<h4>VMs mit abweichenden Pfaden</h4>"
                $html += (ConvertTo-HTMLTable -Data $deviating -NoDataMessage "Alle VMs verwenden den Standardpfad.")
            }
            catch { }
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "Standardpfade (VHD &amp; VM-Konfiguration)" -Category "Hyper-V Host" -Content $html
}

# ---------------------------------------------------------------
# 15. LIVE MIGRATION
# ---------------------------------------------------------------
function Get-LiveMigrationInfo {
    <#
    .SYNOPSIS
        Dokumentiert die Live-Migration- und Storage-Migration-Konfiguration.
    #>
    Write-Log -Message "=== Sammle Live Migration Konfiguration ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Best Practice</h3><p>Für Live Migration außerhalb eines Clusters wird <strong>Kerberos</strong> "
    $html += "mit eingeschränkter Delegierung empfohlen. Die Option &quot;Beliebiges Netzwerk verwenden&quot; sollte deaktiviert und "
    $html += "stattdessen ein dediziertes Migrationsnetzwerk konfiguriert sein. Als Performance-Option ist <strong>SMB</strong> (mit RDMA) "
    $html += "oder <strong>Compression</strong> sinnvoll.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $vmHost = Get-VMHost -ComputerName $server -ErrorAction Stop

            $mig = [PSCustomObject]@{
                "Live Migration aktiviert"          = ConvertTo-DisplayValue -Value $vmHost.VirtualMachineMigrationEnabled
                "Authentifizierungstyp"             = ConvertTo-DisplayValue -Value $vmHost.VirtualMachineMigrationAuthenticationType
                "Bewertung Authentifizierung"       = if ("$($vmHost.VirtualMachineMigrationAuthenticationType)" -eq "Kerberos") { "✅ Kerberos (empfohlen)" } else { "⚠️ CredSSP - erfordert Anmeldung am Quellhost" }
                "Performance-Option"                = ConvertTo-DisplayValue -Value $vmHost.VirtualMachineMigrationPerformanceOption
                "Beliebiges Netzwerk verwenden"     = ConvertTo-DisplayValue -Value $vmHost.UseAnyNetworkForMigration
                "Bewertung Netzwerk"                = if ($vmHost.UseAnyNetworkForMigration) { "⚠️ Alle Netzwerke erlaubt - dediziertes Migrationsnetz empfohlen" } else { "✅ Nur definierte Netzwerke" }
                "Max. gleichzeitige VM-Migrationen" = ConvertTo-DisplayValue -Value $vmHost.MaximumVirtualMachineMigrations
                "Max. gleichzeitige Storage-Migrationen" = ConvertTo-DisplayValue -Value $vmHost.MaximumStorageMigrations
            }

            $html += "<h4>Migrationseinstellungen</h4>"
            $html += (ConvertTo-HTMLTable -Data @($mig))

            # --- Migrationsnetzwerke ---
            try {
                $migNets = Get-VMMigrationNetwork -ComputerName $server -ErrorAction Stop
                $netData = foreach ($n in $migNets) {
                    [PSCustomObject]@{
                        "Subnetz"     = $n.Subnet
                        "Priorität"   = $n.Priority
                        "ComputerName" = $n.ComputerName
                    }
                }
                $html += "<h4>Definierte Migrationsnetzwerke</h4>"
                $html += (ConvertTo-HTMLTable -Data $netData -NoDataMessage "Keine dedizierten Migrationsnetzwerke definiert.")
            }
            catch {
                $html += "<h4>Definierte Migrationsnetzwerke</h4><p class='no-data'>Nicht verfügbar.</p>"
            }
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "Live Migration &amp; Storage Migration" -Category "Hyper-V Host" -Content $html
}

# ---------------------------------------------------------------
# 16. ENHANCED SESSION MODE & RESOURCE METERING
# ---------------------------------------------------------------
function Get-EnhancedSessionModeInfo {
    <#
    .SYNOPSIS
        Dokumentiert Enhanced Session Mode (Host + VMs) und Resource Metering.
    #>
    Write-Log -Message "=== Sammle Enhanced Session Mode / Resource Metering ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $vmHost = Get-VMHost -ComputerName $server -ErrorAction Stop

            $esm = [PSCustomObject]@{
                "Enhanced Session Mode Richtlinie (Host)" = ConvertTo-DisplayValue -Value $vmHost.EnableEnhancedSessionMode
                "Resource Metering Speicherintervall"     = ConvertTo-DisplayValue -Value $vmHost.ResourceMeteringSaveInterval
            }
            $html += "<h4>Host-Einstellungen</h4>"
            $html += (ConvertTo-HTMLTable -Data @($esm))

            # --- Enhanced Session Mode je VM ---
            try {
                $vms = Get-VM -ComputerName $server -ErrorAction Stop
                $vmEsm = foreach ($vm in $vms) {
                    [PSCustomObject]@{
                        VM                            = $vm.Name
                        Generation                    = $vm.Generation
                        "Enhanced Session Transport"  = ConvertTo-DisplayValue -Value $vm.EnhancedSessionTransportType
                        "Guest Services aktiv"        = try { ConvertTo-DisplayValue -Value ((Get-VMIntegrationService -VMName $vm.Name -ComputerName $server -Name "Guest Service Interface" -ErrorAction SilentlyContinue).Enabled) } catch { "-" }
                    }
                }
                $html += "<h4>Enhanced Session Mode je VM</h4>"
                $html += (ConvertTo-HTMLTable -Data $vmEsm -NoDataMessage "Keine VMs gefunden.")
            }
            catch { }
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "Enhanced Session Mode &amp; Resource Metering" -Category "Hyper-V Host" -Content $html
}

# ---------------------------------------------------------------
# 17. HYPER-V DIENSTE
# ---------------------------------------------------------------
function Get-HyperVServiceStatusInfo {
    <#
    .SYNOPSIS
        Prüft die Hyper-V relevanten Windows-Dienste auf Status und Starttyp.
    #>
    Write-Log -Message "=== Sammle Hyper-V Dienststatus ===" -Level "INFO"

    $criticalServices = @(
        @{ Name = "vmms";        Desc = "Hyper-V Virtual Machine Management" }
        @{ Name = "vmcompute";   Desc = "Hyper-V Host Compute Service" }
        @{ Name = "nvspwmi";     Desc = "Hyper-V Networking Management" }
        @{ Name = "hvhost";      Desc = "Hyper-V Host Service" }
        @{ Name = "vmickvpexchange"; Desc = "Hyper-V Data Exchange (KVP)" }
        @{ Name = "vmicheartbeat";   Desc = "Hyper-V Heartbeat" }
        @{ Name = "vmicshutdown";    Desc = "Hyper-V Guest Shutdown" }
        @{ Name = "vmictimesync";    Desc = "Hyper-V Time Synchronization" }
        @{ Name = "vmicvss";         Desc = "Hyper-V Volume Shadow Copy Requestor" }
        @{ Name = "vmicrdv";         Desc = "Hyper-V Remote Desktop Virtualization" }
        @{ Name = "vmicguestinterface"; Desc = "Hyper-V Guest Service Interface" }
        @{ Name = "ClusSvc";         Desc = "Cluster Service (Failover Cluster)" }
        @{ Name = "VSS";             Desc = "Volume Shadow Copy" }
        @{ Name = "WinRM";           Desc = "Windows Remote Management" }
    )

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $cimSession = New-ServerCimSession -ComputerName $server
            if (-not $cimSession) {
                $html += "<p class='error'>Keine Verbindung möglich.</p>"
                continue
            }

            $allServices = Get-CimInstance -CimSession $cimSession -ClassName Win32_Service -ErrorAction SilentlyContinue
            Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue

            $data = foreach ($cs in $criticalServices) {
                $svc = $allServices | Where-Object { $_.Name -eq $cs.Name }
                if ($svc) {
                    [PSCustomObject]@{
                        Dienstname  = $svc.Name
                        Beschreibung = $cs.Desc
                        Anzeigename = $svc.DisplayName
                        Status      = $svc.State
                        Starttyp    = $svc.StartMode
                        Konto       = $svc.StartName
                        Bewertung   = if ($svc.State -eq "Running") { "✅ Läuft" } elseif ($svc.StartMode -eq "Disabled") { "ℹ️ Deaktiviert" } else { "⚠️ Gestoppt" }
                    }
                }
                else {
                    [PSCustomObject]@{
                        Dienstname  = $cs.Name
                        Beschreibung = $cs.Desc
                        Anzeigename = "-"
                        Status      = "Nicht installiert"
                        Starttyp    = "-"
                        Konto       = "-"
                        Bewertung   = "ℹ️ Nicht vorhanden"
                    }
                }
            }

            $html += (ConvertTo-HTMLTable -Data $data)
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "Hyper-V Dienststatus" -Category "Hyper-V Host" -Content $html
}

# ---------------------------------------------------------------
# 18. UNTERSTÜTZTE VM-VERSIONEN
# ---------------------------------------------------------------
function Get-HyperVVersionInfo {
    <#
    .SYNOPSIS
        Dokumentiert die vom Host unterstützten VM-Konfigurationsversionen sowie
        die tatsächlich verwendeten Versionen der VMs.
    #>
    Write-Log -Message "=== Sammle unterstützte VM-Konfigurationsversionen ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Hinweis</h3><p>VMs mit veralteter Konfigurationsversion können neue Hyper-V Funktionen nicht nutzen. "
    $html += "Ein Upgrade (Update-VMVersion) ist erst sinnvoll, wenn alle Cluster-Knoten die neue Version unterstützen und keine Rückmigration nötig ist.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"

        try {
            $supported = Get-VMHostSupportedVersion -ComputerName $server -ErrorAction Stop
            $supData = foreach ($s in $supported) {
                [PSCustomObject]@{
                    Version       = $s.Version
                    Name          = $s.Name
                    "Standard"    = ConvertTo-DisplayValue -Value $s.IsDefault
                }
            }
            $html += "<h4>Unterstützte Konfigurationsversionen</h4>"
            $html += (ConvertTo-HTMLTable -Data $supData -NoDataMessage "Nicht verfügbar.")
        }
        catch {
            $html += "<h4>Unterstützte Konfigurationsversionen</h4><p class='no-data'>Nicht verfügbar: $($_.Exception.Message)</p>"
        }

        try {
            $vms = Get-VM -ComputerName $server -ErrorAction Stop
            $maxVersion = 0.0
            try { $maxVersion = [double](($vms | ForEach-Object { [double]$_.Version } | Measure-Object -Maximum).Maximum) } catch { }
            $verData = foreach ($vm in $vms) {
                $v = 0.0
                try { $v = [double]$vm.Version } catch { }
                [PSCustomObject]@{
                    VM              = $vm.Name
                    "VM-Version"    = $vm.Version
                    Generation      = $vm.Generation
                    Status          = $vm.State
                    Bewertung       = if ($maxVersion -gt 0 -and $v -lt $maxVersion) { "⚠️ Veraltet (aktuell max. $maxVersion)" } else { "✅ Aktuell" }
                }
            }
            $html += "<h4>VM-Konfigurationsversionen</h4>"
            $html += (ConvertTo-HTMLTable -Data $verData -NoDataMessage "Keine VMs gefunden.")
        }
        catch {
            $html += "<h4>VM-Konfigurationsversionen</h4><p class='no-data'>Nicht verfügbar.</p>"
        }
    }

    New-HTMLSection -Title "VM-Konfigurationsversionen" -Category "Hyper-V Host" -Content $html
}

# ---------------------------------------------------------------
# 19. HOST-RESSOURCENAUSLASTUNG
# ---------------------------------------------------------------
function Get-HostResourceUsageInfo {
    <#
    .SYNOPSIS
        Ermittelt die Auslastung des Hosts (Hypervisor Logical Processor, Speicher)
        und die Überbuchungs-Kennzahlen (vCPU:pCPU, zugewiesener RAM).
    #>
    Write-Log -Message "=== Sammle Host-Ressourcenauslastung ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $vmHost = Get-VMHost -ComputerName $server -ErrorAction Stop
            $vms    = @(Get-VM -ComputerName $server -ErrorAction SilentlyContinue)
            $runningVMs = @($vms | Where-Object { $_.State -eq 'Running' })

            $totalVCPU = 0
            foreach ($vm in $runningVMs) {
                try { $totalVCPU += (Get-VMProcessor -VMName $vm.Name -ComputerName $server -ErrorAction SilentlyContinue).Count } catch { }
            }

            $assignedMemory = 0
            foreach ($vm in $runningVMs) { $assignedMemory += [double]$vm.MemoryAssigned }

            $ratio = if ($vmHost.LogicalProcessorCount -gt 0) { [math]::Round($totalVCPU / $vmHost.LogicalProcessorCount, 2) } else { 0 }
            $memPercent = if ($vmHost.MemoryCapacity -gt 0) { [math]::Round(($assignedMemory / $vmHost.MemoryCapacity) * 100, 1) } else { 0 }

            $usage = [PSCustomObject]@{
                "VMs gesamt"                 = $vms.Count
                "VMs laufend"                = $runningVMs.Count
                "Logische Prozessoren (Host)" = $vmHost.LogicalProcessorCount
                "vCPUs zugewiesen (laufend)" = $totalVCPU
                "vCPU : pCPU Verhältnis"     = "$ratio : 1"
                "Bewertung CPU"              = if ($ratio -gt 8) { "⚠️ Sehr hohe Überbuchung" } elseif ($ratio -gt 4) { "ℹ️ Erhöhte Überbuchung" } else { "✅ Unkritisch" }
                "Host-RAM gesamt"            = Format-ByteSize -Bytes $vmHost.MemoryCapacity
                "RAM an VMs zugewiesen"      = Format-ByteSize -Bytes $assignedMemory
                "RAM-Auslastung durch VMs"   = "$memPercent %"
                "Bewertung RAM"              = if ($memPercent -gt 90) { "⚠️ Kaum Reserve für den Host" } elseif ($memPercent -gt 80) { "ℹ️ Hohe Auslastung" } else { "✅ OK" }
            }

            $html += "<h4>Kapazität &amp; Überbuchung</h4>"
            $html += (ConvertTo-HTMLTable -Data @($usage))

            # --- Performance Counter ---
            $counters = Invoke-RemoteCommand -ComputerName $server -ScriptBlock {
                try {
                    $c = Get-Counter -Counter @(
                        "\Hyper-V Hypervisor Logical Processor(_Total)\% Total Run Time",
                        "\Hyper-V Hypervisor Root Virtual Processor(_Total)\% Total Run Time",
                        "\Hyper-V Dynamic Memory Balancer(*)\Available Memory",
                        "\Hyper-V Virtual Switch(*)\Bytes/sec"
                    ) -ErrorAction SilentlyContinue
                    $c.CounterSamples | Select-Object Path, CookedValue
                }
                catch { $null }
            }

            if ($counters) {
                $cData = foreach ($c in $counters) {
                    [PSCustomObject]@{
                        Zähler = $c.Path
                        Wert   = [math]::Round([double]$c.CookedValue, 2)
                    }
                }
                $html += "<h4>Performance Counter (Momentaufnahme)</h4>"
                $html += (ConvertTo-HTMLTable -Data $cData)
            }
            else {
                $html += "<h4>Performance Counter (Momentaufnahme)</h4><p class='no-data'>Nicht verfügbar (WinRM erforderlich).</p>"
            }
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "Host-Ressourcenauslastung &amp; Überbuchung" -Category "Hyper-V Host" -Content $html
}

# ---------------------------------------------------------------
# 20. VM-ÜBERSICHT
# ---------------------------------------------------------------
function Get-VMOverviewInfo {
    <#
    .SYNOPSIS
        Erstellt eine Übersicht aller virtuellen Maschinen je Host.
    #>
    Write-Log -Message "=== Sammle VM-Übersicht ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $vms = Get-VM -ComputerName $server -ErrorAction Stop | Sort-Object Name

            $summary = [PSCustomObject]@{
                "VMs gesamt"     = @($vms).Count
                "Laufend"        = @($vms | Where-Object { $_.State -eq 'Running' }).Count
                "Ausgeschaltet"  = @($vms | Where-Object { $_.State -eq 'Off' }).Count
                "Gespeichert"    = @($vms | Where-Object { $_.State -eq 'Saved' }).Count
                "Angehalten"     = @($vms | Where-Object { $_.State -eq 'Paused' }).Count
                "Generation 1"   = @($vms | Where-Object { $_.Generation -eq 1 }).Count
                "Generation 2"   = @($vms | Where-Object { $_.Generation -eq 2 }).Count
                "Geclustert"     = @($vms | Where-Object { $_.IsClustered }).Count
            }
            $html += "<h4>Zusammenfassung</h4>"
            $html += (ConvertTo-HTMLTable -Data @($summary))

            $data = foreach ($vm in $vms) {
                [PSCustomObject]@{
                    VM                = $vm.Name
                    Status            = $vm.State
                    Generation        = $vm.Generation
                    "Version"         = $vm.Version
                    "vCPU"            = $vm.ProcessorCount
                    "RAM zugewiesen"  = Format-ByteSize -Bytes $vm.MemoryAssigned
                    "RAM Startwert"   = Format-ByteSize -Bytes $vm.MemoryStartup
                    "Dyn. Speicher"   = ConvertTo-DisplayValue -Value $vm.DynamicMemoryEnabled
                    "CPU-Auslastung %" = $vm.CPUUsage
                    Betriebszeit      = if ($vm.Uptime) { $vm.Uptime.ToString() } else { "-" }
                    "Zustand"         = $vm.Status
                    Geclustert        = ConvertTo-DisplayValue -Value $vm.IsClustered
                    "Erstellt am"     = if ($vm.CreationTime) { $vm.CreationTime.ToString("dd.MM.yyyy HH:mm") } else { "-" }
                    Notizen           = ConvertTo-DisplayValue -Value $vm.Notes
                }
            }
            $html += "<h4>Virtuelle Maschinen</h4>"
            $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine VMs auf diesem Host.")
        }
        catch {
            Write-Log -Message "VM-Übersicht fehlgeschlagen für ${server}: $_" -Level "ERROR"
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "Virtuelle Maschinen - Übersicht" -Category "Virtuelle Maschinen" -Content $html
}

# ---------------------------------------------------------------
# 21. VM PROZESSOR-KONFIGURATION
# ---------------------------------------------------------------
function Get-VMProcessorInfo {
    <#
    .SYNOPSIS
        Dokumentiert die Prozessor-Konfiguration aller VMs (Get-VMProcessor).
    #>
    Write-Log -Message "=== Sammle VM Prozessor-Konfiguration ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Hinweis</h3><p><strong>Reserve</strong> garantiert Rechenzeit, <strong>Limit</strong> begrenzt sie, "
    $html += "<strong>Relative Gewichtung</strong> steuert die Priorität bei Ressourcenknappheit (Standard 100). "
    $html += "<strong>ExposeVirtualizationExtensions</strong> aktiviert Nested Virtualization.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $vms = Get-VM -ComputerName $server -ErrorAction Stop | Sort-Object Name
            $data = foreach ($vm in $vms) {
                try {
                    $p = Get-VMProcessor -VMName $vm.Name -ComputerName $server -ErrorAction Stop
                    [PSCustomObject]@{
                        VM                              = $vm.Name
                        "Anzahl vCPU"                   = $p.Count
                        "Reserve (%)"                   = $p.Reserve
                        "Limit (%)"                     = $p.Maximum
                        "Relative Gewichtung"           = $p.RelativeWeight
                        "Max. Proz. pro NUMA-Node"      = $p.MaximumCountPerNumaNode
                        "Max. Nodes pro Socket"         = $p.MaximumCountPerNumaSocket
                        "Nested Virtualization"         = ConvertTo-DisplayValue -Value $p.ExposeVirtualizationExtensions
                        "Kompatibilität (Migration)"    = ConvertTo-DisplayValue -Value $p.CompatibilityForMigrationEnabled
                        "Kompatibilität (ältere OS)"    = ConvertTo-DisplayValue -Value $p.CompatibilityForOlderOperatingSystemsEnabled
                        "HW-Threads pro Kern"           = ConvertTo-DisplayValue -Value $p.HwThreadCountPerCore
                        "Host Resource Protection"      = ConvertTo-DisplayValue -Value $p.EnableHostResourceProtection
                        "Legacy CPU Perf. Zähler"       = ConvertTo-DisplayValue -Value $p.EnableLegacyApicMode
                        "Perfmon Zähler (PMU)"          = ConvertTo-DisplayValue -Value $p.EnablePerfmonPmu
                    }
                }
                catch {
                    [PSCustomObject]@{ VM = $vm.Name; "Anzahl vCPU" = "Fehler"; "Reserve (%)" = "-"; "Limit (%)" = "-"; "Relative Gewichtung" = "-"; "Max. Proz. pro NUMA-Node" = "-"; "Max. Nodes pro Socket" = "-"; "Nested Virtualization" = "-"; "Kompatibilität (Migration)" = "-"; "Kompatibilität (ältere OS)" = "-"; "HW-Threads pro Kern" = "-"; "Host Resource Protection" = "-"; "Legacy CPU Perf. Zähler" = "-"; "Perfmon Zähler (PMU)" = "-" }
                }
            }
            $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine VMs gefunden.")
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "VM Prozessor-Konfiguration" -Category "Virtuelle Maschinen" -Content $html
}

# ---------------------------------------------------------------
# 22. VM ARBEITSSPEICHER
# ---------------------------------------------------------------
function Get-VMMemoryInfo {
    <#
    .SYNOPSIS
        Dokumentiert die Speicherkonfiguration aller VMs inkl. Dynamic Memory.
    #>
    Write-Log -Message "=== Sammle VM Speicher-Konfiguration ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Hinweis</h3><p>Dynamic Memory ist für Server-Workloads wie SQL Server oder Exchange nicht empfohlen. "
    $html += "Der <strong>Speicherpuffer</strong> definiert die zusätzliche Reserve in Prozent, die <strong>Speichergewichtung</strong> die Priorität bei Knappheit.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $vms = Get-VM -ComputerName $server -ErrorAction Stop | Sort-Object Name
            $data = foreach ($vm in $vms) {
                try {
                    $m = Get-VMMemory -VMName $vm.Name -ComputerName $server -ErrorAction Stop
                    [PSCustomObject]@{
                        VM                       = $vm.Name
                        "Dynamic Memory"         = ConvertTo-DisplayValue -Value $m.DynamicMemoryEnabled
                        "Startspeicher"          = Format-ByteSize -Bytes $m.Startup
                        "Minimum"                = Format-ByteSize -Bytes $m.Minimum
                        "Maximum"                = Format-ByteSize -Bytes $m.Maximum
                        "Zugewiesen (aktuell)"   = Format-ByteSize -Bytes $vm.MemoryAssigned
                        "Bedarf (Demand)"        = Format-ByteSize -Bytes $vm.MemoryDemand
                        "Speicherstatus"         = ConvertTo-DisplayValue -Value $vm.MemoryStatus
                        "Puffer (%)"             = $m.Buffer
                        "Gewichtung"             = $m.Priority
                        "Smart Paging Pfad"      = $vm.SmartPagingFilePath
                        "Smart Paging aktiv"     = ConvertTo-DisplayValue -Value $vm.SmartPagingFileInUse
                        "Bewertung"              = if ($m.DynamicMemoryEnabled -and $vm.MemoryDemand -gt $m.Maximum) { "⚠️ Bedarf über Maximum" } else { "✅ OK" }
                    }
                }
                catch {
                    [PSCustomObject]@{ VM = $vm.Name; "Dynamic Memory" = "Fehler"; "Startspeicher" = "-"; "Minimum" = "-"; "Maximum" = "-"; "Zugewiesen (aktuell)" = "-"; "Bedarf (Demand)" = "-"; "Speicherstatus" = "-"; "Puffer (%)" = "-"; "Gewichtung" = "-"; "Smart Paging Pfad" = "-"; "Smart Paging aktiv" = "-"; "Bewertung" = "-" }
                }
            }
            $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine VMs gefunden.")
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "VM Arbeitsspeicher &amp; Dynamic Memory" -Category "Virtuelle Maschinen" -Content $html
}

# ---------------------------------------------------------------
# 23. VM SPEICHER (VHD/VHDX)
# ---------------------------------------------------------------
function Get-VMStorageInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle virtuellen Festplatten der VMs inkl. Typ, Größe,
        Fragmentierung, Differencing-Kette und Storage QoS.
    #>
    Write-Log -Message "=== Sammle VM Speicher-Konfiguration ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $vms = Get-VM -ComputerName $server -ErrorAction Stop | Sort-Object Name
            $data = foreach ($vm in $vms) {
                $drives = @()
                try { $drives = @(Get-VMHardDiskDrive -VMName $vm.Name -ComputerName $server -ErrorAction Stop) } catch { $drives = @() }

                foreach ($d in $drives) {
                    $vhd = $null
                    try { $vhd = Get-VHD -Path $d.Path -ComputerName $server -ErrorAction SilentlyContinue } catch { }

                    [PSCustomObject]@{
                        VM                   = $vm.Name
                        Controller           = "$($d.ControllerType) $($d.ControllerNumber):$($d.ControllerLocation)"
                        Pfad                 = $d.Path
                        Format               = if ($vhd) { $vhd.VhdFormat } else { "-" }
                        Typ                  = if ($vhd) { $vhd.VhdType } else { "-" }
                        "Max. Größe"         = if ($vhd) { Format-ByteSize -Bytes $vhd.Size } else { "-" }
                        "Belegt auf Disk"    = if ($vhd) { Format-ByteSize -Bytes $vhd.FileSize } else { "-" }
                        "Fragmentierung %"   = if ($vhd -and $null -ne $vhd.FragmentationPercentage) { $vhd.FragmentationPercentage } else { "-" }
                        "Logische Sektorgr." = if ($vhd) { $vhd.LogicalSectorSize } else { "-" }
                        "Physische Sektorgr." = if ($vhd) { $vhd.PhysicalSectorSize } else { "-" }
                        "Übergeordnet (Parent)" = if ($vhd -and $vhd.ParentPath) { $vhd.ParentPath } else { "-" }
                        "Min. IOPS (QoS)"    = ConvertTo-DisplayValue -Value $d.MinimumIOPS
                        "Max. IOPS (QoS)"    = ConvertTo-DisplayValue -Value $d.MaximumIOPS
                        "QoS Policy"         = ConvertTo-DisplayValue -Value $d.QoSPolicyID
                        "Freigabe (Shared)"  = ConvertTo-DisplayValue -Value $d.SupportPersistentReservations
                        Bewertung            = if ($vhd -and $vhd.VhdType -eq 'Differencing') { "⚠️ Differencing Disk" } elseif ($vhd -and $vhd.FragmentationPercentage -gt $script:WarningVHDFragPercent) { "⚠️ Hohe Fragmentierung" } else { "✅ OK" }
                    }
                }
            }
            $html += "<h4>Virtuelle Festplatten</h4>"
            $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine virtuellen Festplatten gefunden.")

            # --- Pass-Through Disks ---
            $passThrough = foreach ($vm in $vms) {
                try {
                    $ptd = Get-VMHardDiskDrive -VMName $vm.Name -ComputerName $server -ErrorAction SilentlyContinue |
                        Where-Object { -not $_.Path }
                    foreach ($p in $ptd) {
                        [PSCustomObject]@{
                            VM          = $vm.Name
                            Controller  = "$($p.ControllerType) $($p.ControllerNumber):$($p.ControllerLocation)"
                            "Disk Nr."  = ConvertTo-DisplayValue -Value $p.DiskNumber
                            Hinweis     = "⚠️ Pass-Through Disk (keine Checkpoints/Live-Migration mit Storage möglich)"
                        }
                    }
                }
                catch { }
            }
            $html += "<h4>Pass-Through Disks</h4>"
            $html += (ConvertTo-HTMLTable -Data $passThrough -NoDataMessage "Keine Pass-Through Disks konfiguriert.")
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "VM Speicher (VHD/VHDX)" -Category "Virtuelle Maschinen" -Content $html
}

# ---------------------------------------------------------------
# 24. VM NETZWERKADAPTER
# ---------------------------------------------------------------
function Get-VMNetworkAdapterInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle virtuellen Netzwerkadapter der VMs.
    #>
    Write-Log -Message "=== Sammle VM Netzwerkadapter ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $adapters = Get-VMNetworkAdapter -ComputerName $server -All -ErrorAction Stop
            $data = foreach ($a in $adapters) {
                [PSCustomObject]@{
                    VM                   = ConvertTo-DisplayValue -Value $a.VMName -EmptyText "(Management OS)"
                    Adaptername          = $a.Name
                    "Virtueller Switch"  = ConvertTo-DisplayValue -Value $a.SwitchName
                    "MAC-Adresse"        = ConvertTo-DisplayValue -Value $a.MacAddress
                    "Dynamische MAC"     = ConvertTo-DisplayValue -Value $a.DynamicMacAddressEnabled
                    "IP-Adressen"        = ConvertTo-DisplayValue -Value $a.IPAddresses
                    Status               = ConvertTo-DisplayValue -Value $a.Status
                    "Verbindungsstatus"  = ConvertTo-DisplayValue -Value $a.Connected
                    "Legacy Adapter"     = ConvertTo-DisplayValue -Value $a.IsLegacy
                    "Device Naming"      = ConvertTo-DisplayValue -Value $a.DeviceNaming
                    "VMQ Gewichtung"     = ConvertTo-DisplayValue -Value $a.VmqWeight
                    "IOV Gewichtung"     = ConvertTo-DisplayValue -Value $a.IovWeight
                    "IOV Queue Paare"    = ConvertTo-DisplayValue -Value $a.IovQueuePairsRequested
                    "Min. Bandbreite (Abs)" = ConvertTo-DisplayValue -Value $a.BandwidthSetting.MinimumBandwidthAbsolute
                    "Max. Bandbreite (Abs)" = ConvertTo-DisplayValue -Value $a.BandwidthSetting.MaximumBandwidth
                    "Min. Bandbreite (Gew.)" = ConvertTo-DisplayValue -Value $a.BandwidthSetting.MinimumBandwidthWeight
                }
            }
            $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine Netzwerkadapter gefunden.")
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "VM Netzwerkadapter" -Category "Virtuelle Netzwerke" -Content $html
}

# ---------------------------------------------------------------
# 25. VM NETZWERK - VLAN, SICHERHEIT & ERWEITERTE FEATURES
# ---------------------------------------------------------------
function Get-VMNetworkAdvancedInfo {
    <#
    .SYNOPSIS
        Dokumentiert VLAN-Konfiguration, Portsicherheit (DHCP Guard, Router Guard,
        MAC Spoofing, Port Mirroring) sowie ACLs der VM-Netzwerkadapter.
    #>
    Write-Log -Message "=== Sammle erweiterte VM-Netzwerkeinstellungen ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Best Practice</h3><p><strong>MAC Address Spoofing</strong> sollte nur aktiviert werden, wenn benötigt "
    $html += "(z. B. NLB oder Nested Virtualization). <strong>DHCP Guard</strong> und <strong>Router Guard</strong> schützen vor "
    $html += "unerwünschten DHCP-/Router-Ankündigungen aus VMs.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $adapters = Get-VMNetworkAdapter -ComputerName $server -All -ErrorAction Stop

            # --- VLAN ---
            $vlanData = foreach ($a in $adapters) {
                try {
                    $v = Get-VMNetworkAdapterVlan -VMNetworkAdapter $a -ErrorAction Stop
                    [PSCustomObject]@{
                        VM                = ConvertTo-DisplayValue -Value $a.VMName -EmptyText "(Management OS)"
                        Adapter           = $a.Name
                        "VLAN-Modus"      = ConvertTo-DisplayValue -Value $v.OperationMode
                        "Access VLAN ID"  = ConvertTo-DisplayValue -Value $v.AccessVlanId
                        "Native VLAN ID"  = ConvertTo-DisplayValue -Value $v.NativeVlanId
                        "Trunk VLAN-Liste" = ConvertTo-DisplayValue -Value $v.AllowedVlanIdList
                        "Primary VLAN"    = ConvertTo-DisplayValue -Value $v.PrimaryVlanId
                        "Secondary VLAN"  = ConvertTo-DisplayValue -Value $v.SecondaryVlanId
                    }
                }
                catch { $null }
            }
            $html += "<h4>VLAN-Konfiguration</h4>"
            $html += (ConvertTo-HTMLTable -Data $vlanData -NoDataMessage "Keine VLAN-Konfiguration gefunden.")

            # --- Port-Sicherheit ---
            $secData = foreach ($a in $adapters) {
                [PSCustomObject]@{
                    VM                    = ConvertTo-DisplayValue -Value $a.VMName -EmptyText "(Management OS)"
                    Adapter               = $a.Name
                    "MAC Spoofing"        = ConvertTo-DisplayValue -Value $a.MacAddressSpoofing
                    "DHCP Guard"          = ConvertTo-DisplayValue -Value $a.DhcpGuard
                    "Router Guard"        = ConvertTo-DisplayValue -Value $a.RouterGuard
                    "Port Mirroring"      = ConvertTo-DisplayValue -Value $a.PortMirroringMode
                    "Protected Network"   = ConvertTo-DisplayValue -Value $a.ClusterMonitored
                    "NIC Teaming (VM)"    = ConvertTo-DisplayValue -Value $a.AllowTeaming
                    "Storm Limit"         = ConvertTo-DisplayValue -Value $a.StormLimit
                    "IPsec Offload (SA)"  = ConvertTo-DisplayValue -Value $a.IPsecOffloadMaxSA
                    "VRSS aktiviert"      = ConvertTo-DisplayValue -Value $a.VrssEnabled
                    "VMMQ aktiviert"      = ConvertTo-DisplayValue -Value $a.VmmqEnabled
                    Bewertung             = if ($a.MacAddressSpoofing -eq 'On') { "⚠️ MAC Spoofing aktiv" } else { "✅ OK" }
                }
            }
            $html += "<h4>Portsicherheit &amp; erweiterte Features</h4>"
            $html += (ConvertTo-HTMLTable -Data $secData -NoDataMessage "Keine Adapter gefunden.")

            # --- Port ACLs ---
            $aclData = foreach ($a in $adapters) {
                try {
                    $acls = Get-VMNetworkAdapterAcl -VMNetworkAdapter $a -ErrorAction SilentlyContinue
                    foreach ($acl in $acls) {
                        [PSCustomObject]@{
                            VM         = ConvertTo-DisplayValue -Value $a.VMName -EmptyText "(Management OS)"
                            Adapter    = $a.Name
                            Richtung   = ConvertTo-DisplayValue -Value $acl.Direction
                            Aktion     = ConvertTo-DisplayValue -Value $acl.Action
                            "Lokale Adresse"  = ConvertTo-DisplayValue -Value $acl.LocalAddress
                            "Remote Adresse"  = ConvertTo-DisplayValue -Value $acl.RemoteAddress
                        }
                    }
                }
                catch { }
            }
            $html += "<h4>Port ACLs</h4>"
            $html += (ConvertTo-HTMLTable -Data $aclData -NoDataMessage "Keine Port-ACLs konfiguriert.")

            # --- Erweiterte Adapter-Features (Get-VMNetworkAdapterExtendedAcl) ---
            $extAclData = foreach ($a in $adapters) {
                try {
                    $eacls = Get-VMNetworkAdapterExtendedAcl -VMNetworkAdapter $a -ErrorAction SilentlyContinue
                    foreach ($e in $eacls) {
                        [PSCustomObject]@{
                            VM          = ConvertTo-DisplayValue -Value $a.VMName -EmptyText "(Management OS)"
                            Adapter     = $a.Name
                            Richtung    = ConvertTo-DisplayValue -Value $e.Direction
                            Aktion      = ConvertTo-DisplayValue -Value $e.Action
                            "Lokale IP" = ConvertTo-DisplayValue -Value $e.LocalIPAddress
                            "Remote IP" = ConvertTo-DisplayValue -Value $e.RemoteIPAddress
                            "Lokaler Port" = ConvertTo-DisplayValue -Value $e.LocalPort
                            "Remote Port"  = ConvertTo-DisplayValue -Value $e.RemotePort
                            Protokoll   = ConvertTo-DisplayValue -Value $e.Protocol
                            "Gewichtung" = ConvertTo-DisplayValue -Value $e.Weight
                            Stateful    = ConvertTo-DisplayValue -Value $e.Stateful
                        }
                    }
                }
                catch { }
            }
            $html += "<h4>Erweiterte Port ACLs</h4>"
            $html += (ConvertTo-HTMLTable -Data $extAclData -NoDataMessage "Keine erweiterten ACLs konfiguriert.")
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "VM Netzwerk - VLAN, Sicherheit &amp; ACLs" -Category "Virtuelle Netzwerke" -Content $html
}

# ---------------------------------------------------------------
# 26. INTEGRATIONSDIENSTE
# ---------------------------------------------------------------
function Get-VMIntegrationServicesInfo {
    <#
    .SYNOPSIS
        Dokumentiert die Integrationsdienste je VM.
    #>
    Write-Log -Message "=== Sammle Integrationsdienste ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Hinweis</h3><p>Bei virtualisierten Domänencontrollern sollte der Integrationsdienst "
    $html += "<strong>Zeitsynchronisierung</strong> deaktiviert werden, damit die AD-Zeithierarchie nicht gestört wird.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $vms = Get-VM -ComputerName $server -ErrorAction Stop | Sort-Object Name
            $data = foreach ($vm in $vms) {
                try {
                    $is = Get-VMIntegrationService -VMName $vm.Name -ComputerName $server -ErrorAction Stop
                    $row = [ordered]@{ VM = $vm.Name }
                    foreach ($s in $is) {
                        $state = if ($s.Enabled) { "✅ Aktiviert" } else { "❌ Deaktiviert" }
                        if ($s.PrimaryStatusDescription -and $s.PrimaryStatusDescription -ne "OK") {
                            $state += " ($($s.PrimaryStatusDescription))"
                        }
                        $row[$s.Name] = $state
                    }
                    [PSCustomObject]$row
                }
                catch { $null }
            }
            $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine Integrationsdienste gefunden.")

            # --- Integration Services Version (KVP) ---
            $verData = foreach ($vm in $vms) {
                [PSCustomObject]@{
                    VM                          = $vm.Name
                    "Integration Services Ver." = ConvertTo-DisplayValue -Value $vm.IntegrationServicesVersion
                    "Status"                    = ConvertTo-DisplayValue -Value $vm.IntegrationServicesState
                    "Heartbeat"                 = ConvertTo-DisplayValue -Value $vm.Heartbeat
                }
            }
            $html += "<h4>Versionen &amp; Status</h4>"
            $html += (ConvertTo-HTMLTable -Data $verData -NoDataMessage "Keine Daten verfügbar.")
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "Integrationsdienste" -Category "Virtuelle Maschinen" -Content $html
}

# ---------------------------------------------------------------
# 27. PRÜFPUNKTE / CHECKPOINTS
# ---------------------------------------------------------------
function Get-VMCheckpointInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle Prüfpunkte (Checkpoints/Snapshots) inkl. Alter-Bewertung.
    #>
    Write-Log -Message "=== Sammle Checkpoints ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Best Practice</h3><p>Checkpoints sind <strong>kein Backup</strong>. Sie sollten nicht länger als "
    $html += "$($script:WarningCheckpointDays) Tage bestehen, da die Differencing-Disks (AVHDX) wachsen und die Performance beeinträchtigen.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $vms = Get-VM -ComputerName $server -ErrorAction Stop | Sort-Object Name

            $data = foreach ($vm in $vms) {
                try {
                    $snaps = Get-VMSnapshot -VMName $vm.Name -ComputerName $server -ErrorAction SilentlyContinue
                    foreach ($s in $snaps) {
                        $age = [math]::Round(((Get-Date) - $s.CreationTime).TotalDays, 1)
                        [PSCustomObject]@{
                            VM                = $vm.Name
                            Prüfpunkt         = $s.Name
                            Typ               = ConvertTo-DisplayValue -Value $s.SnapshotType
                            "Erstellt am"     = $s.CreationTime.ToString("dd.MM.yyyy HH:mm")
                            "Alter (Tage)"    = $age
                            "Übergeordnet"    = ConvertTo-DisplayValue -Value $s.ParentSnapshotName
                            Bewertung         = if ($age -gt $script:WarningCheckpointDays) { "⚠️ Älter als $($script:WarningCheckpointDays) Tage!" } else { "✅ OK" }
                        }
                    }
                }
                catch { }
            }
            $html += "<h4>Vorhandene Prüfpunkte</h4>"
            $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine Prüfpunkte vorhanden.")

            # --- Checkpoint-Einstellungen je VM ---
            $cfgData = foreach ($vm in $vms) {
                [PSCustomObject]@{
                    VM                          = $vm.Name
                    "Prüfpunkte aktiviert"      = ConvertTo-DisplayValue -Value $vm.CheckpointType
                    "Automatische Prüfpunkte"   = ConvertTo-DisplayValue -Value $vm.AutomaticCheckpointsEnabled
                    "Prüfpunktpfad"             = $vm.SnapshotFileLocation
                    Bewertung                   = if ($vm.AutomaticCheckpointsEnabled) { "⚠️ Automatische Prüfpunkte aktiv (in Produktion meist unerwünscht)" } else { "✅ OK" }
                }
            }
            $html += "<h4>Prüfpunkt-Einstellungen</h4>"
            $html += (ConvertTo-HTMLTable -Data $cfgData -NoDataMessage "Keine VMs gefunden.")
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "Prüfpunkte (Checkpoints)" -Category "Virtuelle Maschinen" -Content $html
}

# ---------------------------------------------------------------
# 28. VM FIRMWARE / BIOS
# ---------------------------------------------------------------
function Get-VMFirmwareInfo {
    <#
    .SYNOPSIS
        Dokumentiert Bootreihenfolge und Secure Boot (Gen 2) bzw. BIOS-Einstellungen (Gen 1).
    #>
    Write-Log -Message "=== Sammle VM Firmware/BIOS-Konfiguration ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $vms = Get-VM -ComputerName $server -ErrorAction Stop | Sort-Object Name

            # --- Generation 2: Firmware ---
            $fwData = foreach ($vm in ($vms | Where-Object { $_.Generation -eq 2 })) {
                try {
                    $fw = Get-VMFirmware -VMName $vm.Name -ComputerName $server -ErrorAction Stop
                    $bootOrder = ($fw.BootOrder | ForEach-Object {
                        if ($_.Device) { "$($_.BootType): $($_.Device.Name)" } else { "$($_.BootType)" }
                    }) -join " → "
                    [PSCustomObject]@{
                        VM                    = $vm.Name
                        "Secure Boot"         = ConvertTo-DisplayValue -Value $fw.SecureBoot
                        "Secure Boot Vorlage" = ConvertTo-DisplayValue -Value $fw.SecureBootTemplate
                        "Bootreihenfolge"     = ConvertTo-DisplayValue -Value $bootOrder
                        "PXE Netzwerkadapter" = ConvertTo-DisplayValue -Value ($fw.BootOrder | Where-Object { $_.BootType -eq 'Network' } | ForEach-Object { $_.Device.Name })
                        "Konsolenmodus"       = ConvertTo-DisplayValue -Value $fw.ConsoleMode
                        "Pause nach Boot-Fehler" = ConvertTo-DisplayValue -Value $fw.PauseAfterBootFailure
                        Bewertung             = if ($fw.SecureBoot -eq 'On') { "✅ Secure Boot aktiv" } else { "⚠️ Secure Boot deaktiviert" }
                    }
                }
                catch { $null }
            }
            $html += "<h4>Generation 2 - UEFI Firmware</h4>"
            $html += (ConvertTo-HTMLTable -Data $fwData -NoDataMessage "Keine Generation-2-VMs vorhanden.")

            # --- Generation 1: BIOS ---
            $biosData = foreach ($vm in ($vms | Where-Object { $_.Generation -eq 1 })) {
                try {
                    $bios = Get-VMBios -VMName $vm.Name -ComputerName $server -ErrorAction Stop
                    [PSCustomObject]@{
                        VM                    = $vm.Name
                        "Bootreihenfolge"     = ConvertTo-DisplayValue -Value $bios.StartupOrder
                        "Num Lock aktiviert"  = ConvertTo-DisplayValue -Value $bios.NumLockEnabled
                    }
                }
                catch { $null }
            }
            $html += "<h4>Generation 1 - BIOS</h4>"
            $html += (ConvertTo-HTMLTable -Data $biosData -NoDataMessage "Keine Generation-1-VMs vorhanden.")
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "VM Firmware / BIOS &amp; Secure Boot" -Category "Virtuelle Maschinen" -Content $html
}

# ---------------------------------------------------------------
# 29. VM SICHERHEIT (vTPM, SHIELDED VM)
# ---------------------------------------------------------------
function Get-VMSecurityInfo {
    <#
    .SYNOPSIS
        Dokumentiert vTPM, Shielded VM Status und Verschlüsselung des VM-Zustands
        sowie des Migrationsverkehrs.
    #>
    Write-Log -Message "=== Sammle VM Sicherheitskonfiguration ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Hinweis</h3><p><strong>vTPM</strong> ist Voraussetzung für BitLocker in der VM und für "
    $html += "Windows 11 Gäste. <strong>Shielded VMs</strong> erfordern zusätzlich einen Host Guardian Service (HGS).</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $vms = Get-VM -ComputerName $server -ErrorAction Stop | Sort-Object Name
            $data = foreach ($vm in $vms) {
                try {
                    $sec = Get-VMSecurity -VMName $vm.Name -ComputerName $server -ErrorAction Stop
                    [PSCustomObject]@{
                        VM                                = $vm.Name
                        Generation                        = $vm.Generation
                        "vTPM aktiviert"                  = ConvertTo-DisplayValue -Value $sec.TpmEnabled
                        "Shielded VM"                     = ConvertTo-DisplayValue -Value $sec.Shielded
                        "Zustand &amp; Migration verschlüsselt" = ConvertTo-DisplayValue -Value $sec.EncryptStateAndVmMigrationTraffic
                        "Virtualization Based Security"   = ConvertTo-DisplayValue -Value $sec.VirtualizationBasedSecurityOptOut
                        "Bindung an Host"                 = ConvertTo-DisplayValue -Value $sec.BindToHostTpm
                        "KSD (Key Storage Drive)"         = ConvertTo-DisplayValue -Value $sec.KsdEnabled
                        Bewertung                         = if ($sec.Shielded) { "🛡️ Shielded" } elseif ($sec.TpmEnabled) { "✅ vTPM aktiv" } else { "ℹ️ Standard" }
                    }
                }
                catch {
                    [PSCustomObject]@{
                        VM = $vm.Name; Generation = $vm.Generation; "vTPM aktiviert" = "-"; "Shielded VM" = "-"
                        "Zustand &amp; Migration verschlüsselt" = "-"; "Virtualization Based Security" = "-"
                        "Bindung an Host" = "-"; "KSD (Key Storage Drive)" = "-"; Bewertung = "ℹ️ Nicht unterstützt (Gen 1)"
                    }
                }
            }
            $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine VMs gefunden.")
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "VM Sicherheit (vTPM / Shielded VM)" -Category "Sicherheit" -Content $html
}

# ---------------------------------------------------------------
# 30. AUTOMATISCHE START-/STOP-AKTIONEN
# ---------------------------------------------------------------
function Get-VMAutomaticActionsInfo {
    <#
    .SYNOPSIS
        Dokumentiert das Verhalten der VMs beim Start und Herunterfahren des Hosts.
    #>
    Write-Log -Message "=== Sammle automatische Start-/Stop-Aktionen ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Best Practice</h3><p>Als Stop-Aktion sollte <strong>Shutdown</strong> (sauberes Herunterfahren) "
    $html += "statt <strong>Save</strong> gewählt werden. Für Domänencontroller und Datenbank-VMs empfiehlt sich ein gestaffelter "
    $html += "Startverzögerungswert. In Failover-Clustern werden diese Einstellungen von der Cluster-Rolle überschrieben.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $vms = Get-VM -ComputerName $server -ErrorAction Stop | Sort-Object Name
            $data = foreach ($vm in $vms) {
                [PSCustomObject]@{
                    VM                      = $vm.Name
                    "Start-Aktion"          = ConvertTo-DisplayValue -Value $vm.AutomaticStartAction
                    "Startverzögerung (s)"  = ConvertTo-DisplayValue -Value $vm.AutomaticStartDelay
                    "Stop-Aktion"           = ConvertTo-DisplayValue -Value $vm.AutomaticStopAction
                    "Kritische Fehleraktion" = ConvertTo-DisplayValue -Value $vm.AutomaticCriticalErrorAction
                    "Timeout Fehleraktion"  = ConvertTo-DisplayValue -Value $vm.AutomaticCriticalErrorActionTimeout
                    Geclustert              = ConvertTo-DisplayValue -Value $vm.IsClustered
                    Bewertung               = if ("$($vm.AutomaticStopAction)" -eq "Save") { "⚠️ Save - besser Shutdown" } elseif ("$($vm.AutomaticStopAction)" -eq "TurnOff") { "⚠️ TurnOff - Datenverlust möglich" } else { "✅ OK" }
                }
            }
            $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine VMs gefunden.")
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "Automatische Start-/Stop-Aktionen" -Category "Virtuelle Maschinen" -Content $html
}

# ---------------------------------------------------------------
# 31. DVD-LAUFWERKE & CONTROLLER
# ---------------------------------------------------------------
function Get-VMControllerInfo {
    <#
    .SYNOPSIS
        Dokumentiert DVD-Laufwerke, SCSI-/IDE-Controller, COM-Ports und Diskettenlaufwerke.
    #>
    Write-Log -Message "=== Sammle VM Controller und Laufwerke ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $vms = Get-VM -ComputerName $server -ErrorAction Stop | Sort-Object Name

            # --- DVD-Laufwerke ---
            $dvdData = foreach ($vm in $vms) {
                try {
                    $dvds = Get-VMDvdDrive -VMName $vm.Name -ComputerName $server -ErrorAction SilentlyContinue
                    foreach ($d in $dvds) {
                        [PSCustomObject]@{
                            VM             = $vm.Name
                            Controller     = "$($d.ControllerType) $($d.ControllerNumber):$($d.ControllerLocation)"
                            "ISO-Pfad"     = ConvertTo-DisplayValue -Value $d.Path -EmptyText "(leer)"
                            Bewertung      = if ($d.Path) { "⚠️ ISO eingebunden (blockiert ggf. Live Migration)" } else { "✅ Leer" }
                        }
                    }
                }
                catch { }
            }
            $html += "<h4>DVD-Laufwerke</h4>"
            $html += (ConvertTo-HTMLTable -Data $dvdData -NoDataMessage "Keine DVD-Laufwerke konfiguriert.")

            # --- SCSI-Controller ---
            $scsiData = foreach ($vm in $vms) {
                try {
                    $ctrls = Get-VMScsiController -VMName $vm.Name -ComputerName $server -ErrorAction SilentlyContinue
                    foreach ($c in $ctrls) {
                        [PSCustomObject]@{
                            VM                  = $vm.Name
                            "Controller-Nummer" = $c.ControllerNumber
                            "Angeschl. Geräte"  = @($c.Drives).Count
                            "Geräte"            = ConvertTo-DisplayValue -Value (@($c.Drives | ForEach-Object { $_.Name }) -join ', ')
                        }
                    }
                }
                catch { }
            }
            $html += "<h4>SCSI-Controller</h4>"
            $html += (ConvertTo-HTMLTable -Data $scsiData -NoDataMessage "Keine SCSI-Controller gefunden.")

            # --- COM-Ports ---
            $comData = foreach ($vm in $vms) {
                try {
                    $coms = Get-VMComPort -VMName $vm.Name -ComputerName $server -ErrorAction SilentlyContinue
                    foreach ($c in $coms) {
                        if ($c.Path) {
                            [PSCustomObject]@{
                                VM       = $vm.Name
                                Port     = $c.Name
                                "Pipe-Pfad" = $c.Path
                                "Debugger" = ConvertTo-DisplayValue -Value $c.DebuggerMode
                            }
                        }
                    }
                }
                catch { }
            }
            $html += "<h4>COM-Ports (aktiv konfiguriert)</h4>"
            $html += (ConvertTo-HTMLTable -Data $comData -NoDataMessage "Keine COM-Ports konfiguriert.")

            # --- Diskettenlaufwerke (nur Gen 1) ---
            $fdData = foreach ($vm in ($vms | Where-Object { $_.Generation -eq 1 })) {
                try {
                    $fd = Get-VMFloppyDiskDrive -VMName $vm.Name -ComputerName $server -ErrorAction SilentlyContinue
                    if ($fd -and $fd.Path) {
                        [PSCustomObject]@{ VM = $vm.Name; "VFD-Pfad" = $fd.Path }
                    }
                }
                catch { }
            }
            $html += "<h4>Diskettenlaufwerke</h4>"
            $html += (ConvertTo-HTMLTable -Data $fdData -NoDataMessage "Keine Diskettenimages eingebunden.")

            # --- Fibre Channel Adapter ---
            $fcData = foreach ($vm in $vms) {
                try {
                    $fcs = Get-VMFibreChannelHba -VMName $vm.Name -ComputerName $server -ErrorAction SilentlyContinue
                    foreach ($f in $fcs) {
                        [PSCustomObject]@{
                            VM              = $vm.Name
                            "SAN-Name"      = $f.SanName
                            "WWNN A"        = $f.WorldWideNodeNameSetA
                            "WWPN A"        = $f.WorldWidePortNameSetA
                            "WWNN B"        = $f.WorldWideNodeNameSetB
                            "WWPN B"        = $f.WorldWidePortNameSetB
                        }
                    }
                }
                catch { }
            }
            $html += "<h4>Virtuelle Fibre Channel Adapter</h4>"
            $html += (ConvertTo-HTMLTable -Data $fcData -NoDataMessage "Keine virtuellen FC-Adapter konfiguriert.")
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "VM Controller, DVD &amp; COM-Ports" -Category "Virtuelle Maschinen" -Content $html
}

# ---------------------------------------------------------------
# 32. VM-GRUPPEN
# ---------------------------------------------------------------
function Get-VMGroupInfo {
    <#
    .SYNOPSIS
        Dokumentiert konfigurierte VM-Gruppen (VMCollectionType / ManagementCollectionType).
    #>
    Write-Log -Message "=== Sammle VM-Gruppen ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $groups = Get-VMGroup -ComputerName $server -ErrorAction Stop
            $data = foreach ($g in $groups) {
                [PSCustomObject]@{
                    Gruppenname  = $g.Name
                    Typ          = ConvertTo-DisplayValue -Value $g.GroupType
                    "Instanz-ID" = ConvertTo-DisplayValue -Value $g.InstanceId
                    "VM-Mitglieder" = ConvertTo-DisplayValue -Value (@($g.VMMembers | ForEach-Object { $_.Name }) -join ', ')
                    "Gruppen-Mitglieder" = ConvertTo-DisplayValue -Value (@($g.VMGroupMembers | ForEach-Object { $_.Name }) -join ', ')
                }
            }
            $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine VM-Gruppen konfiguriert.")
        }
        catch {
            $html += "<p class='no-data'>VM-Gruppen konnten nicht abgefragt werden: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "VM-Gruppen" -Category "Virtuelle Maschinen" -Content $html
}

# ---------------------------------------------------------------
# 33. GAST-BETRIEBSSYSTEME (KVP)
# ---------------------------------------------------------------
function Get-VMGuestOperatingSystemInfo {
    <#
    .SYNOPSIS
        Liest die Gast-Betriebssysteminformationen über den KVP-Austausch
        (Msvm_KvpExchangeComponent) aus.
    #>
    Write-Log -Message "=== Sammle Gast-Betriebssysteminformationen (KVP) ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Hinweis</h3><p>Die Werte stammen aus dem Integrationsdienst &quot;Datenaustausch&quot; (KVP). "
    $html += "Ist dieser deaktiviert oder das Gastsystem nicht unterstützt, bleiben die Felder leer.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $cimSession = New-ServerCimSession -ComputerName $server
            if (-not $cimSession) {
                $html += "<p class='error'>Keine Verbindung möglich.</p>"
                continue
            }

            $vmsCim = Get-CimInstance -CimSession $cimSession -Namespace "root\virtualization\v2" -ClassName Msvm_ComputerSystem -Filter "Caption='Virtual Machine'" -ErrorAction SilentlyContinue

            $data = foreach ($vmCim in $vmsCim) {
                $kvpValues = @{}
                try {
                    $kvp = Get-CimAssociatedInstance -CimSession $cimSession -InputObject $vmCim -ResultClassName Msvm_KvpExchangeComponent -ErrorAction SilentlyContinue
                    foreach ($item in $kvp.GuestIntrinsicExchangeItems) {
                        $x = [xml]$item
                        $name  = ($x.INSTANCE.PROPERTY | Where-Object { $_.Name -eq 'Name' }).VALUE
                        $value = ($x.INSTANCE.PROPERTY | Where-Object { $_.Name -eq 'Data' }).VALUE
                        if ($name) { $kvpValues[$name] = $value }
                    }
                }
                catch { }

                [PSCustomObject]@{
                    VM                  = $vmCim.ElementName
                    "Gast-Hostname"     = ConvertTo-DisplayValue -Value $kvpValues['FullyQualifiedDomainName']
                    "Betriebssystem"    = ConvertTo-DisplayValue -Value $kvpValues['OSName']
                    "OS Version"        = ConvertTo-DisplayValue -Value $kvpValues['OSVersion']
                    "OS Build"          = ConvertTo-DisplayValue -Value $kvpValues['OSBuildNumber']
                    "Architektur"       = ConvertTo-DisplayValue -Value $kvpValues['ProcessorArchitecture']
                    "Service Pack"      = ConvertTo-DisplayValue -Value $kvpValues['ServicePackMajor']
                    "IS Version"        = ConvertTo-DisplayValue -Value $kvpValues['IntegrationServicesVersion']
                    "Netzwerkadressen"  = ConvertTo-DisplayValue -Value $kvpValues['NetworkAddressIPv4']
                }
            }

            Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
            $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine Gastinformationen verfügbar.")
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "Gast-Betriebssysteme (KVP)" -Category "Virtuelle Maschinen" -Content $html
}

# ---------------------------------------------------------------
# 34. RESOURCE METERING
# ---------------------------------------------------------------
function Get-VMResourceMeteringInfo {
    <#
    .SYNOPSIS
        Liest die Ressourcenmessung (Measure-VM) für VMs mit aktiviertem
        Resource Metering aus.
    #>
    Write-Log -Message "=== Sammle Resource Metering Daten ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Hinweis</h3><p>Resource Metering muss je VM mit <code>Enable-VMResourceMetering</code> aktiviert werden. "
    $html += "Die Werte sind kumulativ seit der letzten Zurücksetzung.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $vms = Get-VM -ComputerName $server -ErrorAction Stop | Sort-Object Name

            $data = foreach ($vm in $vms) {
                try {
                    $m = Measure-VM -VMName $vm.Name -ComputerName $server -ErrorAction Stop
                    [PSCustomObject]@{
                        VM                        = $vm.Name
                        "Durchschn. CPU (MHz)"    = $m.AverageProcessorUsage
                        "Durchschn. RAM (MB)"     = $m.AverageMemoryUsage
                        "Max. RAM (MB)"           = $m.MaximumMemoryUsage
                        "Min. RAM (MB)"           = $m.MinimumMemoryUsage
                        "Gesamtspeicher (GB)"     = $m.TotalDiskAllocation
                        "Eingehend (MB)"          = ConvertTo-DisplayValue -Value ($m.NetworkMeteredTrafficReport | Measure-Object -Property TotalTraffic -Sum).Sum
                        "Messzeitraum"            = ConvertTo-DisplayValue -Value $m.MeteringDuration
                    }
                }
                catch {
                    [PSCustomObject]@{
                        VM = $vm.Name; "Durchschn. CPU (MHz)" = "Metering inaktiv"; "Durchschn. RAM (MB)" = "-"
                        "Max. RAM (MB)" = "-"; "Min. RAM (MB)" = "-"; "Gesamtspeicher (GB)" = "-"
                        "Eingehend (MB)" = "-"; "Messzeitraum" = "-"
                    }
                }
            }
            $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine VMs gefunden.")
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "Ressourcenmessung (Resource Metering)" -Category "Virtuelle Maschinen" -Content $html
}

# ---------------------------------------------------------------
# 35. VIRTUELLE SWITCHES
# ---------------------------------------------------------------
function Get-VMSwitchInfo {
    <#
    .SYNOPSIS
        Dokumentiert alle virtuellen Switches inkl. Typ, Teaming, SR-IOV und
        Bandbreitenreservierung.
    #>
    Write-Log -Message "=== Sammle virtuelle Switches ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Hinweis</h3><p>In einem Cluster müssen virtuelle Switches auf allen Knoten <strong>identisch benannt</strong> sein, "
    $html += "damit Live Migration funktioniert. <strong>SET</strong> (Switch Embedded Teaming) ist die empfohlene Teaming-Variante ab Windows Server 2016.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $switches = Get-VMSwitch -ComputerName $server -ErrorAction Stop | Sort-Object Name

            $data = foreach ($s in $switches) {
                [PSCustomObject]@{
                    Switch                        = $s.Name
                    Typ                           = ConvertTo-DisplayValue -Value $s.SwitchType
                    "Physischer Adapter"          = ConvertTo-DisplayValue -Value $s.NetAdapterInterfaceDescription
                    "Management OS teilt Adapter" = ConvertTo-DisplayValue -Value $s.AllowManagementOS
                    "SET (Embedded Teaming)"      = ConvertTo-DisplayValue -Value $s.EmbeddedTeamingEnabled
                    "SR-IOV aktiviert"            = ConvertTo-DisplayValue -Value $s.IovEnabled
                    "SR-IOV Support"              = ConvertTo-DisplayValue -Value $s.IovSupport
                    "SR-IOV Grund"                = ConvertTo-DisplayValue -Value $s.IovSupportReasons
                    "Bandbreitenmodus"            = ConvertTo-DisplayValue -Value $s.BandwidthReservationMode
                    "Standard-Mindestbandbreite (abs)" = ConvertTo-DisplayValue -Value $s.DefaultFlowMinimumBandwidthAbsolute
                    "Standard-Mindestbandbreite (Gew.)" = ConvertTo-DisplayValue -Value $s.DefaultFlowMinimumBandwidthWeight
                    "Packet Direct"               = ConvertTo-DisplayValue -Value $s.PacketDirectEnabled
                    "Verfügbare VM-Queues"        = ConvertTo-DisplayValue -Value $s.AvailableVMQueues
                    "Verbundene VMs"              = ConvertTo-DisplayValue -Value (@(Get-VMNetworkAdapter -ComputerName $server -All -ErrorAction SilentlyContinue | Where-Object { $_.SwitchName -eq $s.Name } | ForEach-Object { $_.VMName } | Where-Object { $_ } | Sort-Object -Unique) -join ', ')
                }
            }
            $html += "<h4>Virtuelle Switches</h4>"
            $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine virtuellen Switches gefunden.")

            # --- SET-Team Details ---
            $teamData = foreach ($s in ($switches | Where-Object { $_.EmbeddedTeamingEnabled })) {
                try {
                    $team = Get-VMSwitchTeam -Name $s.Name -ComputerName $server -ErrorAction Stop
                    [PSCustomObject]@{
                        Switch              = $s.Name
                        "Team-Mitglieder"   = ConvertTo-DisplayValue -Value $team.NetAdapterInterfaceDescription
                        "Teaming-Modus"     = ConvertTo-DisplayValue -Value $team.TeamingMode
                        "Lastverteilung"    = ConvertTo-DisplayValue -Value $team.LoadBalancingAlgorithm
                    }
                }
                catch { $null }
            }
            $html += "<h4>Switch Embedded Teaming (SET)</h4>"
            $html += (ConvertTo-HTMLTable -Data $teamData -NoDataMessage "Kein SET-Team konfiguriert.")

            # --- Management-OS Adapter auf Switches ---
            $mgmtData = foreach ($s in $switches) {
                try {
                    $mgmt = Get-VMNetworkAdapter -ComputerName $server -ManagementOS -SwitchName $s.Name -ErrorAction SilentlyContinue
                    foreach ($m in $mgmt) {
                        [PSCustomObject]@{
                            Switch          = $s.Name
                            "vNIC (Host)"   = $m.Name
                            "MAC-Adresse"   = ConvertTo-DisplayValue -Value $m.MacAddress
                            "IP-Adressen"   = ConvertTo-DisplayValue -Value $m.IPAddresses
                            Status          = ConvertTo-DisplayValue -Value $m.Status
                        }
                    }
                }
                catch { }
            }
            $html += "<h4>Host-vNICs (Management OS)</h4>"
            $html += (ConvertTo-HTMLTable -Data $mgmtData -NoDataMessage "Keine Host-vNICs gefunden.")
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "Virtuelle Switches" -Category "Virtuelle Netzwerke" -Content $html
}

# ---------------------------------------------------------------
# 36. SWITCH EXTENSIONS
# ---------------------------------------------------------------
function Get-VMSwitchExtensionInfo {
    <#
    .SYNOPSIS
        Dokumentiert installierte Erweiterungen der virtuellen Switches.
    #>
    Write-Log -Message "=== Sammle Switch Extensions ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $switches = Get-VMSwitch -ComputerName $server -ErrorAction Stop
            $data = foreach ($s in $switches) {
                try {
                    $exts = Get-VMSwitchExtension -VMSwitchName $s.Name -ComputerName $server -ErrorAction SilentlyContinue
                    foreach ($e in $exts) {
                        [PSCustomObject]@{
                            Switch        = $s.Name
                            Erweiterung   = $e.Name
                            Hersteller    = ConvertTo-DisplayValue -Value $e.Vendor
                            Version       = ConvertTo-DisplayValue -Value $e.Version
                            Typ           = ConvertTo-DisplayValue -Value $e.ExtensionType
                            Aktiviert     = ConvertTo-DisplayValue -Value $e.Enabled
                            "Läuft"       = ConvertTo-DisplayValue -Value $e.Running
                        }
                    }
                }
                catch { }
            }
            $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine Switch-Erweiterungen gefunden.")
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "Virtual Switch Extensions" -Category "Virtuelle Netzwerke" -Content $html
}

# ---------------------------------------------------------------
# 37. HOST-NETZWERKADAPTER (VMQ, RSS, SR-IOV, JUMBO)
# ---------------------------------------------------------------
function Get-HostNetworkAdapterInfo {
    <#
    .SYNOPSIS
        Dokumentiert die physischen Netzwerkadapter des Hosts inkl. Offload-Features,
        VMQ, RSS, SR-IOV, Jumbo Frames und NIC-Teaming.
    #>
    Write-Log -Message "=== Sammle Host-Netzwerkadapter ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Best Practice</h3><p><strong>VMQ</strong> sollte auf 10 GbE+ Adaptern aktiviert sein, auf 1 GbE Adaptern "
    $html += "hingegen deaktiviert. <strong>RSS</strong> verteilt die Last auf mehrere CPU-Kerne. Für SMB Direct wird <strong>RDMA</strong> benötigt.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"

        $netInfo = Invoke-RemoteCommand -ComputerName $server -ScriptBlock {
            $result = @{}
            try { $result.Adapters = Get-NetAdapter -ErrorAction SilentlyContinue | Select-Object Name, InterfaceDescription, Status, LinkSpeed, MacAddress, MtuSize, DriverVersion, DriverDate, FullDuplex, MediaType, ifIndex } catch { }
            try { $result.Vmq      = Get-NetAdapterVmq -ErrorAction SilentlyContinue | Select-Object Name, Enabled, NumberOfReceiveQueues, BaseProcessorNumber, MaxProcessors } catch { }
            try { $result.Rss      = Get-NetAdapterRss -ErrorAction SilentlyContinue | Select-Object Name, Enabled, NumberOfReceiveQueues, BaseProcessorNumber, MaxProcessorNumber, Profile } catch { }
            try { $result.Sriov    = Get-NetAdapterSriov -ErrorAction SilentlyContinue | Select-Object Name, Enabled, NumVFs, SriovSupport } catch { }
            try { $result.Rdma     = Get-NetAdapterRdma -ErrorAction SilentlyContinue | Select-Object Name, Enabled, MaxQueuePairCount } catch { }
            try { $result.Teams    = Get-NetLbfoTeam -ErrorAction SilentlyContinue | Select-Object Name, Members, TeamingMode, LoadBalancingAlgorithm, Status } catch { }
            try { $result.Binding  = Get-NetAdapterBinding -ErrorAction SilentlyContinue | Where-Object { $_.Enabled } | Select-Object Name, DisplayName, ComponentID } catch { }
            [PSCustomObject]$result
        }

        if (-not $netInfo) {
            $html += "<p class='no-data'>Netzwerkadapter-Details konnten nicht abgefragt werden (WinRM erforderlich).</p>"
            continue
        }

        if ($netInfo.Adapters) {
            $adapterData = foreach ($a in $netInfo.Adapters) {
                [PSCustomObject]@{
                    Adapter          = $a.Name
                    Beschreibung     = $a.InterfaceDescription
                    Status           = $a.Status
                    Geschwindigkeit  = $a.LinkSpeed
                    "MAC-Adresse"    = $a.MacAddress
                    "MTU"            = $a.MtuSize
                    "Jumbo Frames"   = if ($a.MtuSize -gt 1500) { "✅ Aktiv ($($a.MtuSize))" } else { "Standard (1500)" }
                    Vollduplex       = ConvertTo-DisplayValue -Value $a.FullDuplex
                    Medientyp        = $a.MediaType
                    Treiberversion   = $a.DriverVersion
                    Treiberdatum     = if ($a.DriverDate) { ([datetime]$a.DriverDate).ToString("dd.MM.yyyy") } else { "-" }
                }
            }
            $html += "<h4>Physische Netzwerkadapter</h4>"
            $html += (ConvertTo-HTMLTable -Data $adapterData)
        }

        if ($netInfo.Vmq) {
            $vmqData = foreach ($v in $netInfo.Vmq) {
                [PSCustomObject]@{
                    Adapter            = $v.Name
                    "VMQ aktiviert"    = ConvertTo-DisplayValue -Value $v.Enabled
                    "Empfangsqueues"   = ConvertTo-DisplayValue -Value $v.NumberOfReceiveQueues
                    "Basis-Prozessor"  = ConvertTo-DisplayValue -Value $v.BaseProcessorNumber
                    "Max. Prozessoren" = ConvertTo-DisplayValue -Value $v.MaxProcessors
                }
            }
            $html += "<h4>VMQ (Virtual Machine Queue)</h4>"
            $html += (ConvertTo-HTMLTable -Data $vmqData)
        }

        if ($netInfo.Rss) {
            $rssData = foreach ($r in $netInfo.Rss) {
                [PSCustomObject]@{
                    Adapter            = $r.Name
                    "RSS aktiviert"    = ConvertTo-DisplayValue -Value $r.Enabled
                    "Empfangsqueues"   = ConvertTo-DisplayValue -Value $r.NumberOfReceiveQueues
                    "Basis-Prozessor"  = ConvertTo-DisplayValue -Value $r.BaseProcessorNumber
                    "Max. Prozessor"   = ConvertTo-DisplayValue -Value $r.MaxProcessorNumber
                    Profil             = ConvertTo-DisplayValue -Value $r.Profile
                }
            }
            $html += "<h4>RSS (Receive Side Scaling)</h4>"
            $html += (ConvertTo-HTMLTable -Data $rssData)
        }

        if ($netInfo.Sriov) {
            $sriovData = foreach ($s in $netInfo.Sriov) {
                [PSCustomObject]@{
                    Adapter          = $s.Name
                    "SR-IOV aktiv"   = ConvertTo-DisplayValue -Value $s.Enabled
                    "Virtual Functions" = ConvertTo-DisplayValue -Value $s.NumVFs
                    "Unterstützung"  = ConvertTo-DisplayValue -Value $s.SriovSupport
                }
            }
            $html += "<h4>SR-IOV</h4>"
            $html += (ConvertTo-HTMLTable -Data $sriovData)
        }

        if ($netInfo.Rdma) {
            $rdmaData = foreach ($r in $netInfo.Rdma) {
                [PSCustomObject]@{
                    Adapter           = $r.Name
                    "RDMA aktiviert"  = ConvertTo-DisplayValue -Value $r.Enabled
                    "Max. Queue Pairs" = ConvertTo-DisplayValue -Value $r.MaxQueuePairCount
                }
            }
            $html += "<h4>RDMA (SMB Direct)</h4>"
            $html += (ConvertTo-HTMLTable -Data $rdmaData)
        }

        if ($netInfo.Teams) {
            $teamData = foreach ($t in $netInfo.Teams) {
                [PSCustomObject]@{
                    Team              = $t.Name
                    Mitglieder        = ConvertTo-DisplayValue -Value $t.Members
                    "Teaming-Modus"   = ConvertTo-DisplayValue -Value $t.TeamingMode
                    Lastverteilung    = ConvertTo-DisplayValue -Value $t.LoadBalancingAlgorithm
                    Status            = ConvertTo-DisplayValue -Value $t.Status
                }
            }
            $html += "<h4>NIC-Teaming (LBFO)</h4>"
            $html += (ConvertTo-HTMLTable -Data $teamData -NoDataMessage "Kein LBFO-Team konfiguriert.")
        }
    }

    New-HTMLSection -Title "Host-Netzwerkadapter (VMQ / RSS / SR-IOV / RDMA)" -Category "Virtuelle Netzwerke" -Content $html
}

# ---------------------------------------------------------------
# 38. NETZWERK-QOS / DCB
# ---------------------------------------------------------------
function Get-NetworkQoSInfo {
    <#
    .SYNOPSIS
        Dokumentiert QoS-Richtlinien, Traffic Classes und DCB-Konfiguration
        (relevant für konvergente Netzwerke und SMB Direct).
    #>
    Write-Log -Message "=== Sammle Netzwerk-QoS Konfiguration ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"

        $qos = Invoke-RemoteCommand -ComputerName $server -ScriptBlock {
            $r = @{}
            try { $r.Policies = Get-NetQosPolicy -ErrorAction SilentlyContinue | Select-Object Name, PriorityValue8021Action, NetDirectPortMatchCondition, IPProtocolMatchCondition, TemplateMatchCondition, ThrottleRateActionBitsPerSecond } catch { }
            try { $r.TrafficClasses = Get-NetQosTrafficClass -ErrorAction SilentlyContinue | Select-Object Name, Algorithm, Bandwidth, Priority } catch { }
            try { $r.Flow = Get-NetQosFlowControl -ErrorAction SilentlyContinue | Select-Object Priority, Enabled } catch { }
            try { $r.Dcbx = Get-NetQosDcbxSetting -ErrorAction SilentlyContinue | Select-Object InterfaceAlias, Willing } catch { }
            try { $r.AdapterQos = Get-NetAdapterQos -ErrorAction SilentlyContinue | Select-Object Name, Enabled, OperationalTrafficClasses } catch { }
            [PSCustomObject]$r
        }

        if (-not $qos) {
            $html += "<p class='no-data'>QoS-Informationen konnten nicht abgefragt werden (WinRM / DCB-Feature erforderlich).</p>"
            continue
        }

        $polData = foreach ($p in $qos.Policies) {
            [PSCustomObject]@{
                Richtlinie          = $p.Name
                "802.1p Priorität"  = ConvertTo-DisplayValue -Value $p.PriorityValue8021Action
                "NetDirect Port"    = ConvertTo-DisplayValue -Value $p.NetDirectPortMatchCondition
                "IP Protokoll"      = ConvertTo-DisplayValue -Value $p.IPProtocolMatchCondition
                "Vorlage"           = ConvertTo-DisplayValue -Value $p.TemplateMatchCondition
                "Drosselung (bps)"  = ConvertTo-DisplayValue -Value $p.ThrottleRateActionBitsPerSecond
            }
        }
        $html += "<h4>QoS-Richtlinien</h4>"
        $html += (ConvertTo-HTMLTable -Data $polData -NoDataMessage "Keine QoS-Richtlinien konfiguriert.")

        $tcData = foreach ($t in $qos.TrafficClasses) {
            [PSCustomObject]@{
                "Traffic Class"  = $t.Name
                Algorithmus      = ConvertTo-DisplayValue -Value $t.Algorithm
                "Bandbreite (%)" = ConvertTo-DisplayValue -Value $t.Bandwidth
                Priorität        = ConvertTo-DisplayValue -Value $t.Priority
            }
        }
        $html += "<h4>Traffic Classes (DCB)</h4>"
        $html += (ConvertTo-HTMLTable -Data $tcData -NoDataMessage "Keine Traffic Classes konfiguriert.")

        $flowData = foreach ($f in $qos.Flow) {
            [PSCustomObject]@{
                "Priorität"           = $f.Priority
                "Flow Control aktiv"  = ConvertTo-DisplayValue -Value $f.Enabled
            }
        }
        $html += "<h4>Priority Flow Control (PFC)</h4>"
        $html += (ConvertTo-HTMLTable -Data $flowData -NoDataMessage "PFC nicht konfiguriert.")
    }

    New-HTMLSection -Title "Netzwerk-QoS / DCB" -Category "Virtuelle Netzwerke" -Content $html
}

# ---------------------------------------------------------------
# 39. STORAGE-KONFIGURATION
# ---------------------------------------------------------------
function Get-StorageConfigurationInfo {
    <#
    .SYNOPSIS
        Dokumentiert Disks, Volumes, Partitionen, Storage Spaces und iSCSI.
    #>
    Write-Log -Message "=== Sammle Storage-Konfiguration ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Best Practice</h3><p>Volumes für VHDX-Dateien sollten mit einer <strong>Zuordnungseinheit von 64 KB</strong> "
    $html += "formatiert sein. Für CSVFS wird ReFS (VM-Workloads) oder NTFS verwendet.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"

        $storage = Invoke-RemoteCommand -ComputerName $server -ScriptBlock {
            $r = @{}
            try { $r.Disks = Get-Disk -ErrorAction SilentlyContinue | Select-Object Number, FriendlyName, SerialNumber, Size, PartitionStyle, OperationalStatus, HealthStatus, BusType, IsClustered, IsBoot, IsSystem } catch { }
            try { $r.Volumes = Get-Volume -ErrorAction SilentlyContinue | Select-Object DriveLetter, FileSystemLabel, FileSystem, FileSystemType, Size, SizeRemaining, HealthStatus, AllocationUnitSize, Path } catch { }
            try { $r.PhysicalDisks = Get-PhysicalDisk -ErrorAction SilentlyContinue | Select-Object FriendlyName, MediaType, Size, HealthStatus, OperationalStatus, Usage, CanPool, BusType } catch { }
            try { $r.Pools = Get-StoragePool -ErrorAction SilentlyContinue | Where-Object { -not $_.IsPrimordial } | Select-Object FriendlyName, HealthStatus, OperationalStatus, Size, AllocatedSize, IsReadOnly, ResiliencySettingNameDefault } catch { }
            try { $r.VirtualDisks = Get-VirtualDisk -ErrorAction SilentlyContinue | Select-Object FriendlyName, ResiliencySettingName, Size, FootprintOnPool, HealthStatus, OperationalStatus, NumberOfColumns, ProvisioningType } catch { }
            try { $r.iSCSI = Get-IscsiTarget -ErrorAction SilentlyContinue | Select-Object NodeAddress, IsConnected } catch { }
            [PSCustomObject]$r
        }

        if (-not $storage) {
            $html += "<p class='no-data'>Storage-Informationen konnten nicht abgefragt werden (WinRM erforderlich).</p>"
            continue
        }

        $diskData = foreach ($d in $storage.Disks) {
            [PSCustomObject]@{
                "Disk Nr."      = $d.Number
                Bezeichnung     = $d.FriendlyName
                Seriennummer    = ConvertTo-DisplayValue -Value $d.SerialNumber
                "Größe"         = Format-ByteSize -Bytes $d.Size
                Partitionsstil  = ConvertTo-DisplayValue -Value $d.PartitionStyle
                Bustyp          = ConvertTo-DisplayValue -Value $d.BusType
                Betriebsstatus  = ConvertTo-DisplayValue -Value $d.OperationalStatus
                Zustand         = ConvertTo-DisplayValue -Value $d.HealthStatus
                Geclustert      = ConvertTo-DisplayValue -Value $d.IsClustered
                "Boot/System"   = if ($d.IsBoot -or $d.IsSystem) { "Ja" } else { "Nein" }
            }
        }
        $html += "<h4>Physische Datenträger (Disks)</h4>"
        $html += (ConvertTo-HTMLTable -Data $diskData -NoDataMessage "Keine Datenträger gefunden.")

        $volData = foreach ($v in $storage.Volumes) {
            [PSCustomObject]@{
                Laufwerk           = ConvertTo-DisplayValue -Value $v.DriveLetter
                Bezeichnung        = ConvertTo-DisplayValue -Value $v.FileSystemLabel
                Dateisystem        = ConvertTo-DisplayValue -Value $v.FileSystem
                "Größe"            = Format-ByteSize -Bytes $v.Size
                "Frei"             = Format-ByteSize -Bytes $v.SizeRemaining
                "Frei %"           = if ($v.Size -gt 0) { [math]::Round(($v.SizeRemaining / $v.Size) * 100, 1) } else { 0 }
                "Zuordnungseinheit" = ConvertTo-DisplayValue -Value $v.AllocationUnitSize
                Zustand            = ConvertTo-DisplayValue -Value $v.HealthStatus
                Bewertung          = if ($v.AllocationUnitSize -and $v.AllocationUnitSize -lt 65536 -and $v.FileSystem -eq 'NTFS') { "ℹ️ 64K empfohlen für VHDX" } else { "✅ OK" }
            }
        }
        $html += "<h4>Volumes</h4>"
        $html += (ConvertTo-HTMLTable -Data $volData -NoDataMessage "Keine Volumes gefunden.")

        if ($storage.Pools) {
            $poolData = foreach ($p in $storage.Pools) {
                [PSCustomObject]@{
                    "Storage Pool"   = $p.FriendlyName
                    "Größe"          = Format-ByteSize -Bytes $p.Size
                    "Zugewiesen"     = Format-ByteSize -Bytes $p.AllocatedSize
                    Zustand          = ConvertTo-DisplayValue -Value $p.HealthStatus
                    Betriebsstatus   = ConvertTo-DisplayValue -Value $p.OperationalStatus
                    "Standard-Resilienz" = ConvertTo-DisplayValue -Value $p.ResiliencySettingNameDefault
                }
            }
            $html += "<h4>Storage Pools</h4>"
            $html += (ConvertTo-HTMLTable -Data $poolData -NoDataMessage "Keine Storage Pools vorhanden.")
        }

        if ($storage.VirtualDisks) {
            $vdData = foreach ($v in $storage.VirtualDisks) {
                [PSCustomObject]@{
                    "Virtual Disk"     = $v.FriendlyName
                    Resilienz          = ConvertTo-DisplayValue -Value $v.ResiliencySettingName
                    "Größe"            = Format-ByteSize -Bytes $v.Size
                    "Pool-Belegung"    = Format-ByteSize -Bytes $v.FootprintOnPool
                    Spalten            = ConvertTo-DisplayValue -Value $v.NumberOfColumns
                    Bereitstellung     = ConvertTo-DisplayValue -Value $v.ProvisioningType
                    Zustand            = ConvertTo-DisplayValue -Value $v.HealthStatus
                }
            }
            $html += "<h4>Virtuelle Datenträger (Storage Spaces)</h4>"
            $html += (ConvertTo-HTMLTable -Data $vdData -NoDataMessage "Keine virtuellen Datenträger vorhanden.")
        }

        if ($storage.iSCSI) {
            $iscsiData = foreach ($i in $storage.iSCSI) {
                [PSCustomObject]@{
                    "iSCSI Target" = $i.NodeAddress
                    Verbunden      = ConvertTo-DisplayValue -Value $i.IsConnected
                }
            }
            $html += "<h4>iSCSI Targets</h4>"
            $html += (ConvertTo-HTMLTable -Data $iscsiData -NoDataMessage "Keine iSCSI-Targets konfiguriert.")
        }
    }

    New-HTMLSection -Title "Storage-Konfiguration" -Category "Speicher" -Content $html
}

# ---------------------------------------------------------------
# 40. VHD-DETAILANALYSE
# ---------------------------------------------------------------
function Get-VHDDetailsInfo {
    <#
    .SYNOPSIS
        Analysiert alle VHD/VHDX-Dateien im Standardpfad und identifiziert
        verwaiste (keiner VM zugeordnete) Dateien.
    #>
    Write-Log -Message "=== Sammle VHD-Detailanalyse ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Hinweis</h3><p>Verwaiste VHD/VHDX-Dateien belegen unnötig Speicherplatz. "
    $html += "Vor dem Löschen unbedingt prüfen, ob die Datei zu einer Cluster-Rolle oder einem Backup gehört.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"
        try {
            $vmHost = Get-VMHost -ComputerName $server -ErrorAction Stop
            $searchPath = $vmHost.VirtualHardDiskPath

            # --- Zugeordnete VHDs ermitteln ---
            $attached = @()
            try {
                $vms = Get-VM -ComputerName $server -ErrorAction SilentlyContinue
                foreach ($vm in $vms) {
                    $drives = Get-VMHardDiskDrive -VMName $vm.Name -ComputerName $server -ErrorAction SilentlyContinue
                    foreach ($d in $drives) { if ($d.Path) { $attached += $d.Path } }
                }
            }
            catch { }

            $files = Invoke-RemoteCommand -ComputerName $server -ScriptBlock {
                param($path)
                if ($path -and (Test-Path $path)) {
                    Get-ChildItem -Path $path -Include *.vhd, *.vhdx, *.avhd, *.avhdx -Recurse -File -ErrorAction SilentlyContinue |
                        Select-Object FullName, Length, LastWriteTime, CreationTime
                }
            } -ArgumentList @($searchPath)

            $data = foreach ($f in $files) {
                $isAttached = $attached -contains $f.FullName
                [PSCustomObject]@{
                    Datei             = Split-Path $f.FullName -Leaf
                    Pfad              = $f.FullName
                    "Größe"           = Format-ByteSize -Bytes $f.Length
                    "Zuletzt geändert" = ([datetime]$f.LastWriteTime).ToString("dd.MM.yyyy HH:mm")
                    "Erstellt"        = ([datetime]$f.CreationTime).ToString("dd.MM.yyyy HH:mm")
                    Zuordnung         = if ($isAttached) { "✅ VM zugeordnet" } else { "⚠️ Verwaist / nicht zugeordnet" }
                }
            }

            $html += "<p><strong>Durchsuchter Pfad:</strong> $searchPath</p>"
            $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine VHD/VHDX-Dateien im Standardpfad gefunden oder Zugriff nicht möglich.")
        }
        catch {
            $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
        }
    }

    New-HTMLSection -Title "VHD-Detailanalyse &amp; verwaiste Dateien" -Category "Speicher" -Content $html
}

# ---------------------------------------------------------------
# 41. STORAGE QOS
# ---------------------------------------------------------------
function Get-StorageQoSInfo {
    <#
    .SYNOPSIS
        Dokumentiert Storage QoS Policies (Scale-Out File Server / S2D).
    #>
    Write-Log -Message "=== Sammle Storage QoS Policies ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"

        $qos = Invoke-RemoteCommand -ComputerName $server -ScriptBlock {
            try {
                Get-StorageQosPolicy -ErrorAction SilentlyContinue |
                    Select-Object Name, PolicyId, PolicyType, MinimumIops, MaximumIops, MaximumIoBandwidthFlow, Status
            }
            catch { $null }
        }

        $data = foreach ($q in $qos) {
            [PSCustomObject]@{
                Richtlinie      = $q.Name
                Typ             = ConvertTo-DisplayValue -Value $q.PolicyType
                "Min. IOPS"     = ConvertTo-DisplayValue -Value $q.MinimumIops
                "Max. IOPS"     = ConvertTo-DisplayValue -Value $q.MaximumIops
                "Max. Bandbreite" = ConvertTo-DisplayValue -Value $q.MaximumIoBandwidthFlow
                Status          = ConvertTo-DisplayValue -Value $q.Status
                "Policy-ID"     = ConvertTo-DisplayValue -Value $q.PolicyId
            }
        }
        $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine Storage QoS Policies konfiguriert.")
    }

    New-HTMLSection -Title "Storage QoS Policies" -Category "Speicher" -Content $html
}

# ---------------------------------------------------------------
# 42. SMB 3.0 STORAGE
# ---------------------------------------------------------------
function Get-SMBStorageInfo {
    <#
    .SYNOPSIS
        Dokumentiert SMB-Client-/Server-Konfiguration, Freigaben und
        Multichannel-Verbindungen (relevant bei SMB-basiertem VM-Storage).
    #>
    Write-Log -Message "=== Sammle SMB Storage Konfiguration ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"

        $smb = Invoke-RemoteCommand -ComputerName $server -ScriptBlock {
            $r = @{}
            try { $r.Client = Get-SmbClientConfiguration -ErrorAction SilentlyContinue | Select-Object EnableMultiChannel, EnableLargeMtu, EnableBandwidthThrottling, ConnectionCountPerRssNetworkInterface, RequireSecuritySignature, EnableSecuritySignature } catch { }
            try { $r.Server = Get-SmbServerConfiguration -ErrorAction SilentlyContinue | Select-Object EnableMultiChannel, EnableSMB1Protocol, EnableSMB2Protocol, EncryptData, RejectUnencryptedAccess, RequireSecuritySignature } catch { }
            try { $r.Shares = Get-SmbShare -ErrorAction SilentlyContinue | Select-Object Name, Path, Description, EncryptData, ContinuouslyAvailable, ShareType } catch { }
            try { $r.Connections = Get-SmbConnection -ErrorAction SilentlyContinue | Select-Object ServerName, ShareName, Dialect, NumOpens, Encrypted } catch { }
            try { $r.MultiChannel = Get-SmbMultichannelConnection -ErrorAction SilentlyContinue | Select-Object Server, ClientIpAddress, ServerIpAddress, ClientRdmaCapable, ServerRdmaCapable, ClientLinkSpeed } catch { }
            [PSCustomObject]$r
        }

        if (-not $smb) {
            $html += "<p class='no-data'>SMB-Informationen konnten nicht abgefragt werden (WinRM erforderlich).</p>"
            continue
        }

        if ($smb.Client) {
            $clientData = foreach ($p in ($smb.Client.PSObject.Properties)) {
                [PSCustomObject]@{ Parameter = $p.Name; Wert = ConvertTo-DisplayValue -Value $p.Value }
            }
            $html += "<h4>SMB-Client Konfiguration</h4>"
            $html += (ConvertTo-HTMLTable -Data $clientData)
        }

        if ($smb.Server) {
            $serverData = foreach ($p in ($smb.Server.PSObject.Properties)) {
                [PSCustomObject]@{ Parameter = $p.Name; Wert = ConvertTo-DisplayValue -Value $p.Value }
            }
            $html += "<h4>SMB-Server Konfiguration</h4>"
            $html += (ConvertTo-HTMLTable -Data $serverData)
        }

        $shareData = foreach ($s in $smb.Shares) {
            [PSCustomObject]@{
                Freigabe             = $s.Name
                Pfad                 = ConvertTo-DisplayValue -Value $s.Path
                Beschreibung         = ConvertTo-DisplayValue -Value $s.Description
                "Verschlüsselt"      = ConvertTo-DisplayValue -Value $s.EncryptData
                "Continuously Avail." = ConvertTo-DisplayValue -Value $s.ContinuouslyAvailable
                Typ                  = ConvertTo-DisplayValue -Value $s.ShareType
            }
        }
        $html += "<h4>SMB-Freigaben</h4>"
        $html += (ConvertTo-HTMLTable -Data $shareData -NoDataMessage "Keine Freigaben vorhanden.")

        $mcData = foreach ($m in $smb.MultiChannel) {
            [PSCustomObject]@{
                Server            = $m.Server
                "Client IP"       = $m.ClientIpAddress
                "Server IP"       = $m.ServerIpAddress
                "Client RDMA"     = ConvertTo-DisplayValue -Value $m.ClientRdmaCapable
                "Server RDMA"     = ConvertTo-DisplayValue -Value $m.ServerRdmaCapable
                "Link Speed"      = ConvertTo-DisplayValue -Value $m.ClientLinkSpeed
            }
        }
        $html += "<h4>SMB Multichannel Verbindungen</h4>"
        $html += (ConvertTo-HTMLTable -Data $mcData -NoDataMessage "Keine Multichannel-Verbindungen aktiv.")
    }

    New-HTMLSection -Title "SMB 3.0 Storage" -Category "Speicher" -Content $html
}

# ---------------------------------------------------------------
# 43. MPIO
# ---------------------------------------------------------------
function Get-MPIOInfo {
    <#
    .SYNOPSIS
        Dokumentiert die MPIO-Konfiguration (Multipath I/O) für SAN-Anbindungen.
    #>
    Write-Log -Message "=== Sammle MPIO-Konfiguration ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"

        $mpio = Invoke-RemoteCommand -ComputerName $server -ScriptBlock {
            $r = @{}
            try { $r.Setting = Get-MPIOSetting -ErrorAction SilentlyContinue } catch { }
            try { $r.Available = Get-MPIOAvailableHW -ErrorAction SilentlyContinue | Select-Object VendorId, ProductId, BusType, IsMultipathed } catch { }
            try { $r.Policy = Get-MSDSMGlobalDefaultLoadBalancePolicy -ErrorAction SilentlyContinue } catch { }
            [PSCustomObject]$r
        }

        if (-not $mpio -or (-not $mpio.Setting -and -not $mpio.Available)) {
            $html += "<p class='no-data'>MPIO ist nicht installiert oder nicht abfragbar.</p>"
            continue
        }

        if ($mpio.Setting) {
            $settingData = foreach ($p in ($mpio.Setting.PSObject.Properties)) {
                [PSCustomObject]@{ Parameter = $p.Name; Wert = ConvertTo-DisplayValue -Value $p.Value }
            }
            $html += "<h4>MPIO-Einstellungen</h4>"
            $html += (ConvertTo-HTMLTable -Data $settingData)
        }

        $hwData = foreach ($h in $mpio.Available) {
            [PSCustomObject]@{
                Hersteller     = ConvertTo-DisplayValue -Value $h.VendorId
                Produkt        = ConvertTo-DisplayValue -Value $h.ProductId
                Bustyp         = ConvertTo-DisplayValue -Value $h.BusType
                "Multipath"    = ConvertTo-DisplayValue -Value $h.IsMultipathed
            }
        }
        $html += "<h4>MPIO-fähige Hardware</h4>"
        $html += (ConvertTo-HTMLTable -Data $hwData -NoDataMessage "Keine MPIO-Hardware erkannt.")

        if ($mpio.Policy) {
            $html += "<h4>Standard-Lastverteilungsrichtlinie</h4>"
            $html += (ConvertTo-HTMLTable -Data @([PSCustomObject]@{ "Load Balance Policy" = ConvertTo-DisplayValue -Value $mpio.Policy }))
        }
    }

    New-HTMLSection -Title "MPIO (Multipath I/O)" -Category "Speicher" -Content $html
}

# ---------------------------------------------------------------
# 44. FAILOVER CLUSTER
# ---------------------------------------------------------------
function Get-FailoverClusterInfo {
    <#
    .SYNOPSIS
        Dokumentiert Cluster-Konfiguration, Knoten, Rollen und Ressourcen.
    #>
    Write-Log -Message "=== Sammle Failover Cluster Informationen ===" -Level "INFO"

    if (-not $script:ClusterModuleLoaded) {
        New-HTMLSection -Title "Failover Cluster" -Category "Failover Cluster" -Content "<p class='no-data'>Das Modul FailoverClusters ist nicht verfügbar. Cluster-Dokumentation wurde übersprungen.</p>"
        return
    }

    $html = ""
    $firstServer = @($HyperVServers)[0]

    try {
        $cluster = Get-Cluster -Name $firstServer -ErrorAction Stop

        $clusterProps = foreach ($p in ($cluster.PSObject.Properties | Sort-Object Name)) {
            [PSCustomObject]@{ Parameter = $p.Name; Wert = ConvertTo-DisplayValue -Value $p.Value }
        }

        $summary = [PSCustomObject]@{
            "Clustername"                = $cluster.Name
            "Domäne"                     = ConvertTo-DisplayValue -Value $cluster.Domain
            "Cluster-Funktionsebene"     = ConvertTo-DisplayValue -Value $cluster.ClusterFunctionalLevel
            "Upgrade-Version"            = ConvertTo-DisplayValue -Value $cluster.ClusterUpgradeVersion
            "Dynamisches Quorum"         = ConvertTo-DisplayValue -Value $cluster.DynamicQuorum
            "S2D aktiviert"              = ConvertTo-DisplayValue -Value $cluster.S2DEnabled
            "CSV Block Cache (MB)"       = ConvertTo-DisplayValue -Value $cluster.BlockCacheSize
            "Preferred Site"             = ConvertTo-DisplayValue -Value $cluster.PreferredSite
            "Shared Volumes Root"        = ConvertTo-DisplayValue -Value $cluster.SharedVolumesRoot
        }

        $html += "<h4>Cluster-Übersicht</h4>"
        $html += (ConvertTo-HTMLTable -Data @($summary))

        # --- Knoten ---
        try {
            $nodes = Get-ClusterNode -Cluster $cluster.Name -ErrorAction Stop
            $nodeData = foreach ($n in $nodes) {
                [PSCustomObject]@{
                    Knoten            = $n.Name
                    Status            = ConvertTo-DisplayValue -Value $n.State
                    "Knotengewicht"   = ConvertTo-DisplayValue -Value $n.NodeWeight
                    "Dyn. Gewicht"    = ConvertTo-DisplayValue -Value $n.DynamicWeight
                    "Drain-Status"    = ConvertTo-DisplayValue -Value $n.DrainStatus
                    "Fehleranzahl"    = ConvertTo-DisplayValue -Value $n.NodeHighestVersion
                    Modell            = ConvertTo-DisplayValue -Value $n.Model
                    Hersteller        = ConvertTo-DisplayValue -Value $n.Manufacturer
                    Seriennummer      = ConvertTo-DisplayValue -Value $n.SerialNumber
                    Bewertung         = if ("$($n.State)" -eq "Up") { "✅ Online" } else { "⚠️ $($n.State)" }
                }
            }
            $html += "<h4>Cluster-Knoten</h4>"
            $html += (ConvertTo-HTMLTable -Data $nodeData)
        }
        catch {
            $html += "<h4>Cluster-Knoten</h4><p class='no-data'>Nicht verfügbar.</p>"
        }

        # --- Cluster-Rollen (VMs) ---
        try {
            $groups = Get-ClusterGroup -Cluster $cluster.Name -ErrorAction Stop
            $groupData = foreach ($g in $groups) {
                [PSCustomObject]@{
                    Rolle              = $g.Name
                    Typ                = ConvertTo-DisplayValue -Value $g.GroupType
                    Status             = ConvertTo-DisplayValue -Value $g.State
                    "Besitzerknoten"   = ConvertTo-DisplayValue -Value $g.OwnerNode
                    Priorität          = ConvertTo-DisplayValue -Value $g.Priority
                    "Fehlerbehandlung" = ConvertTo-DisplayValue -Value $g.FailoverThreshold
                    "Failback"         = ConvertTo-DisplayValue -Value $g.AutoFailbackType
                }
            }
            $html += "<h4>Cluster-Rollen</h4>"
            $html += (ConvertTo-HTMLTable -Data $groupData)
        }
        catch {
            $html += "<h4>Cluster-Rollen</h4><p class='no-data'>Nicht verfügbar.</p>"
        }

        # --- Cluster-Ressourcen ---
        try {
            $resources = Get-ClusterResource -Cluster $cluster.Name -ErrorAction Stop
            $resData = foreach ($r in $resources) {
                [PSCustomObject]@{
                    Ressource         = $r.Name
                    Typ               = ConvertTo-DisplayValue -Value $r.ResourceType
                    Status            = ConvertTo-DisplayValue -Value $r.State
                    "Besitzergruppe"  = ConvertTo-DisplayValue -Value $r.OwnerGroup
                    "Besitzerknoten"  = ConvertTo-DisplayValue -Value $r.OwnerNode
                    Bewertung         = if ("$($r.State)" -eq "Online") { "✅ Online" } else { "⚠️ $($r.State)" }
                }
            }
            $html += "<h4>Cluster-Ressourcen</h4>"
            $html += (ConvertTo-HTMLTable -Data $resData)
        }
        catch {
            $html += "<h4>Cluster-Ressourcen</h4><p class='no-data'>Nicht verfügbar.</p>"
        }

        $html += "<h4>Vollständige Cluster-Parameter</h4>"
        $html += (ConvertTo-HTMLTable -Data $clusterProps)
    }
    catch {
        Write-Log -Message "Kein Failover-Cluster erkannt: $_" -Level "INFO"
        $html = "<p class='no-data'>Kein Failover-Cluster erkannt (Standalone-Host) oder keine Berechtigung.</p>"
    }

    New-HTMLSection -Title "Failover Cluster" -Category "Failover Cluster" -Content $html
}

# ---------------------------------------------------------------
# 45. CLUSTER-NETZWERKE
# ---------------------------------------------------------------
function Get-ClusterNetworkInfo {
    <#
    .SYNOPSIS
        Dokumentiert Cluster-Netzwerke und deren Verwendungszweck.
    #>
    Write-Log -Message "=== Sammle Cluster-Netzwerke ===" -Level "INFO"

    if (-not $script:ClusterModuleLoaded) {
        New-HTMLSection -Title "Cluster-Netzwerke" -Category "Failover Cluster" -Content "<p class='no-data'>Modul FailoverClusters nicht verfügbar.</p>"
        return
    }

    $html = "<div class='summary-box'><h3>Best Practice</h3><p>Es sollten getrennte Netzwerke für <strong>Management</strong>, "
    $html += "<strong>Live Migration</strong>, <strong>Cluster/CSV</strong> und <strong>VM-Traffic</strong> existieren "
    $html += "(bzw. ein konvergentes Design mit QoS).</p></div>"

    $firstServer = @($HyperVServers)[0]

    try {
        $cluster = Get-Cluster -Name $firstServer -ErrorAction Stop
        $networks = Get-ClusterNetwork -Cluster $cluster.Name -ErrorAction Stop

        $netData = foreach ($n in $networks) {
            [PSCustomObject]@{
                Netzwerk         = $n.Name
                Adresse          = ConvertTo-DisplayValue -Value $n.Address
                Subnetzmaske     = ConvertTo-DisplayValue -Value $n.AddressMask
                Status           = ConvertTo-DisplayValue -Value $n.State
                Rolle            = switch ("$($n.Role)") { "0" { "Nicht für Cluster verwenden" } "1" { "Nur Clusterkommunikation" } "3" { "Cluster + Clients" } default { "$($n.Role)" } }
                Metrik           = ConvertTo-DisplayValue -Value $n.Metric
                "Auto Metrik"    = ConvertTo-DisplayValue -Value $n.AutoMetric
                Beschreibung     = ConvertTo-DisplayValue -Value $n.Description
            }
        }
        $html += "<h4>Cluster-Netzwerke</h4>"
        $html += (ConvertTo-HTMLTable -Data $netData)

        # --- Netzwerkschnittstellen ---
        try {
            $interfaces = Get-ClusterNetworkInterface -Cluster $cluster.Name -ErrorAction Stop
            $ifData = foreach ($i in $interfaces) {
                [PSCustomObject]@{
                    Knoten        = ConvertTo-DisplayValue -Value $i.Node
                    Schnittstelle = $i.Name
                    Netzwerk      = ConvertTo-DisplayValue -Value $i.Network
                    Adresse       = ConvertTo-DisplayValue -Value $i.Address
                    Status        = ConvertTo-DisplayValue -Value $i.State
                    Adapter       = ConvertTo-DisplayValue -Value $i.Adapter
                }
            }
            $html += "<h4>Netzwerkschnittstellen der Knoten</h4>"
            $html += (ConvertTo-HTMLTable -Data $ifData)
        }
        catch { }

        # --- Live Migration Netzwerkreihenfolge ---
        try {
            $lmSettings = Get-ClusterResourceType -Cluster $cluster.Name -Name "Virtual Machine" -ErrorAction Stop |
                Get-ClusterParameter -Name MigrationNetworkOrder, MigrationExcludeNetworks -ErrorAction SilentlyContinue
            $lmData = foreach ($p in $lmSettings) {
                [PSCustomObject]@{ Parameter = $p.Name; Wert = ConvertTo-DisplayValue -Value $p.Value }
            }
            $html += "<h4>Live Migration Netzwerkpriorität</h4>"
            $html += (ConvertTo-HTMLTable -Data $lmData -NoDataMessage "Keine spezielle Reihenfolge konfiguriert.")
        }
        catch { }
    }
    catch {
        $html += "<p class='no-data'>Kein Failover-Cluster erkannt.</p>"
    }

    New-HTMLSection -Title "Cluster-Netzwerke" -Category "Failover Cluster" -Content $html
}

# ---------------------------------------------------------------
# 46. CLUSTER SHARED VOLUMES
# ---------------------------------------------------------------
function Get-ClusterSharedVolumeInfo {
    <#
    .SYNOPSIS
        Dokumentiert Cluster Shared Volumes (CSV) inkl. Speicherplatz und Redirected-Mode.
    #>
    Write-Log -Message "=== Sammle Cluster Shared Volumes ===" -Level "INFO"

    if (-not $script:ClusterModuleLoaded) {
        New-HTMLSection -Title "Cluster Shared Volumes (CSV)" -Category "Failover Cluster" -Content "<p class='no-data'>Modul FailoverClusters nicht verfügbar.</p>"
        return
    }

    $html = ""
    $firstServer = @($HyperVServers)[0]

    try {
        $cluster = Get-Cluster -Name $firstServer -ErrorAction Stop
        $csvs = Get-ClusterSharedVolume -Cluster $cluster.Name -ErrorAction Stop

        $data = foreach ($csv in $csvs) {
            foreach ($info in $csv.SharedVolumeInfo) {
                $part = $info.Partition
                $freeGB = [math]::Round($part.FreeSpace / 1GB, 2)
                $sizeGB = [math]::Round($part.Size / 1GB, 2)
                [PSCustomObject]@{
                    CSV                = $csv.Name
                    Status             = ConvertTo-DisplayValue -Value $csv.State
                    "Besitzerknoten"   = ConvertTo-DisplayValue -Value $csv.OwnerNode
                    Pfad               = ConvertTo-DisplayValue -Value $info.FriendlyVolumeName
                    Dateisystem        = ConvertTo-DisplayValue -Value $part.FileSystem
                    "Größe_GB"         = $sizeGB
                    "Frei_GB"          = $freeGB
                    "Frei_%"           = if ($sizeGB -gt 0) { [math]::Round(($freeGB / $sizeGB) * 100, 1) } else { 0 }
                    "Redirected Mode"  = ConvertTo-DisplayValue -Value $info.RedirectedAccess
                    "Fault State"      = ConvertTo-DisplayValue -Value $info.FaultState
                    Bewertung          = if ($info.RedirectedAccess) { "⚠️ Redirected Access aktiv" } elseif ($freeGB -lt $script:WarningDiskSpaceGB) { "⚠️ Wenig freier Speicher" } else { "✅ OK" }
                }
            }
        }

        $html += (ConvertTo-HTMLTable -Data $data -NoDataMessage "Keine Cluster Shared Volumes vorhanden.")

        # --- CSV Block Cache ---
        try {
            $html += "<h4>CSV Block Cache</h4>"
            $html += (ConvertTo-HTMLTable -Data @([PSCustomObject]@{
                "Block Cache Größe (MB)" = ConvertTo-DisplayValue -Value $cluster.BlockCacheSize
                Bewertung                = if ($cluster.BlockCacheSize -gt 0) { "✅ Aktiviert" } else { "ℹ️ Deaktiviert (bei reinem SSD-Storage unkritisch)" }
            }))
        }
        catch { }
    }
    catch {
        $html = "<p class='no-data'>Kein Failover-Cluster erkannt oder keine CSVs vorhanden.</p>"
    }

    New-HTMLSection -Title "Cluster Shared Volumes (CSV)" -Category "Failover Cluster" -Content $html
}

# ---------------------------------------------------------------
# 47. CLUSTER-QUORUM
# ---------------------------------------------------------------
function Get-ClusterQuorumInfo {
    <#
    .SYNOPSIS
        Dokumentiert die Quorum-Konfiguration des Clusters.
    #>
    Write-Log -Message "=== Sammle Cluster-Quorum ===" -Level "INFO"

    if (-not $script:ClusterModuleLoaded) {
        New-HTMLSection -Title "Cluster-Quorum" -Category "Failover Cluster" -Content "<p class='no-data'>Modul FailoverClusters nicht verfügbar.</p>"
        return
    }

    $html = "<div class='summary-box'><h3>Best Practice</h3><p>Bei einer geraden Knotenanzahl sollte ein Zeuge (Disk Witness, File Share Witness "
    $html += "oder Cloud Witness) konfiguriert sein.</p></div>"

    $firstServer = @($HyperVServers)[0]

    try {
        $cluster = Get-Cluster -Name $firstServer -ErrorAction Stop
        $quorum = Get-ClusterQuorum -Cluster $cluster.Name -ErrorAction Stop
        $nodeCount = @(Get-ClusterNode -Cluster $cluster.Name -ErrorAction SilentlyContinue).Count

        $data = [PSCustomObject]@{
            Cluster              = $quorum.Cluster
            "Quorum-Typ"         = ConvertTo-DisplayValue -Value $quorum.QuorumType
            "Zeugen-Ressource"   = ConvertTo-DisplayValue -Value $quorum.QuorumResource -EmptyText "Kein Zeuge konfiguriert"
            "Anzahl Knoten"      = $nodeCount
            Bewertung            = if (-not $quorum.QuorumResource -and ($nodeCount % 2) -eq 0) { "⚠️ Gerade Knotenanzahl ohne Zeuge!" } else { "✅ OK" }
        }

        $html += (ConvertTo-HTMLTable -Data @($data))

        # --- Knotengewichtung ---
        try {
            $nodes = Get-ClusterNode -Cluster $cluster.Name -ErrorAction Stop
            $weightData = foreach ($n in $nodes) {
                [PSCustomObject]@{
                    Knoten          = $n.Name
                    Status          = ConvertTo-DisplayValue -Value $n.State
                    "Stimmgewicht"  = ConvertTo-DisplayValue -Value $n.NodeWeight
                    "Dyn. Gewicht"  = ConvertTo-DisplayValue -Value $n.DynamicWeight
                }
            }
            $html += "<h4>Knoten-Stimmgewichtung</h4>"
            $html += (ConvertTo-HTMLTable -Data $weightData)
        }
        catch { }
    }
    catch {
        $html += "<p class='no-data'>Kein Failover-Cluster erkannt.</p>"
    }

    New-HTMLSection -Title "Cluster-Quorum" -Category "Failover Cluster" -Content $html
}

# ---------------------------------------------------------------
# 48. HYPER-V REPLICA
# ---------------------------------------------------------------
function Get-VMReplicationInfo {
    <#
    .SYNOPSIS
        Dokumentiert die Hyper-V Replica Konfiguration (Server + VMs + Autorisierung).
    #>
    Write-Log -Message "=== Sammle Hyper-V Replica Konfiguration ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"

        # --- Replikations-Serverkonfiguration ---
        try {
            $rs = Get-VMReplicationServer -ComputerName $server -ErrorAction Stop
            $rsData = [PSCustomObject]@{
                "Replikat-Server aktiviert"      = ConvertTo-DisplayValue -Value $rs.RepEnabled
                "Authentifizierungstyp"          = ConvertTo-DisplayValue -Value $rs.AuthenticationType
                "Kerberos Port (HTTP)"           = ConvertTo-DisplayValue -Value $rs.KerberosAuthenticationPort
                "Zertifikat Port (HTTPS)"        = ConvertTo-DisplayValue -Value $rs.CertificateAuthenticationPort
                "Zertifikat Thumbprint"          = ConvertTo-DisplayValue -Value $rs.CertificateThumbprint
                "Alle Server zugelassen"         = ConvertTo-DisplayValue -Value $rs.AllowedAuthenticationType
                "Standard-Speicherort"           = ConvertTo-DisplayValue -Value $rs.DefaultStorageLocation
                "Replikation zugelassen von"     = ConvertTo-DisplayValue -Value $rs.AllowAnyServer
                "Monitoring Intervall"           = ConvertTo-DisplayValue -Value $rs.MonitoringInterval
                "Monitoring Startzeit"           = ConvertTo-DisplayValue -Value $rs.MonitoringStartTime
            }
            $html += "<h4>Replikations-Serverkonfiguration</h4>"
            $html += (ConvertTo-HTMLTable -Data @($rsData))
        }
        catch {
            $html += "<h4>Replikations-Serverkonfiguration</h4><p class='no-data'>Nicht verfügbar: $($_.Exception.Message)</p>"
        }

        # --- Autorisierungseinträge ---
        try {
            $auth = Get-VMReplicationAuthorizationEntry -ComputerName $server -ErrorAction Stop
            $authData = foreach ($a in $auth) {
                [PSCustomObject]@{
                    "Primärserver"        = $a.AllowedPrimaryServer
                    "Speicherort"         = ConvertTo-DisplayValue -Value $a.ReplicaStorageLocation
                    "Vertrauensgruppe"    = ConvertTo-DisplayValue -Value $a.TrustGroup
                }
            }
            $html += "<h4>Autorisierungseinträge</h4>"
            $html += (ConvertTo-HTMLTable -Data $authData -NoDataMessage "Keine Autorisierungseinträge konfiguriert.")
        }
        catch {
            $html += "<h4>Autorisierungseinträge</h4><p class='no-data'>Nicht verfügbar.</p>"
        }

        # --- Replikationen je VM ---
        try {
            $reps = Get-VMReplication -ComputerName $server -ErrorAction Stop
            $repData = foreach ($r in $reps) {
                [PSCustomObject]@{
                    VM                        = $r.VMName
                    "Rolle"                   = ConvertTo-DisplayValue -Value $r.ReplicationMode
                    "Status"                  = ConvertTo-DisplayValue -Value $r.ReplicationState
                    "Zustand"                 = ConvertTo-DisplayValue -Value $r.ReplicationHealth
                    "Replikatserver"          = ConvertTo-DisplayValue -Value $r.CurrentReplicaServerName
                    "Primärserver"            = ConvertTo-DisplayValue -Value $r.PrimaryServerName
                    "Replikationsfrequenz (s)" = ConvertTo-DisplayValue -Value $r.ReplicationFrequencySec
                    "Wiederherstellungspunkte" = ConvertTo-DisplayValue -Value $r.RecoveryHistory
                    "VSS-Snapshot Frequenz"   = ConvertTo-DisplayValue -Value $r.VSSSnapshotFrequencyHour
                    "Authentifizierung"       = ConvertTo-DisplayValue -Value $r.AuthenticationType
                    "Komprimierung"           = ConvertTo-DisplayValue -Value $r.CompressionEnabled
                    "Letzte Replikation"      = if ($r.LastReplicationTime) { $r.LastReplicationTime.ToString("dd.MM.yyyy HH:mm") } else { "-" }
                    Bewertung                 = if ("$($r.ReplicationHealth)" -eq "Normal") { "✅ OK" } else { "⚠️ $($r.ReplicationHealth)" }
                }
            }
            $html += "<h4>Replizierte virtuelle Maschinen</h4>"
            $html += (ConvertTo-HTMLTable -Data $repData -NoDataMessage "Keine Replikationen konfiguriert.")
        }
        catch {
            $html += "<h4>Replizierte virtuelle Maschinen</h4><p class='no-data'>Nicht verfügbar.</p>"
        }
    }

    New-HTMLSection -Title "Hyper-V Replica" -Category "Replikation & Backup" -Content $html
}

# ---------------------------------------------------------------
# 49. VSS WRITER & BACKUP
# ---------------------------------------------------------------
function Get-BackupConfigurationInfo {
    <#
    .SYNOPSIS
        Dokumentiert VSS-Writer-Status sowie eine ggf. vorhandene
        Windows Server Backup Konfiguration.
    #>
    Write-Log -Message "=== Sammle Backup- und VSS-Konfiguration ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Hinweis</h3><p>Für konsistente Sicherungen virtueller Maschinen müssen die VSS-Writer "
    $html += "<strong>Microsoft Hyper-V VSS Writer</strong> und <strong>Cluster Shared Volume VSS Writer</strong> im Zustand "
    $html += "<em>Stable / No error</em> sein.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"

        # --- VSS Writer ---
        $writers = Invoke-RemoteCommand -ComputerName $server -ScriptBlock {
            $output = & vssadmin list writers 2>&1 | Out-String
            $blocks = $output -split "Writer name:"
            $result = @()
            foreach ($b in $blocks) {
                if ($b -match "'(.+?)'") {
                    $name = $Matches[1]
                    $state = if ($b -match "State:\s*(.+)") { $Matches[1].Trim() } else { "-" }
                    $lastError = if ($b -match "Last error:\s*(.+)") { $Matches[1].Trim() } else { "-" }
                    $result += [PSCustomObject]@{ Name = $name; State = $state; LastError = $lastError }
                }
            }
            $result
        }

        if ($writers) {
            $writerData = foreach ($w in $writers) {
                [PSCustomObject]@{
                    "VSS Writer"   = $w.Name
                    Status         = $w.State
                    "Letzter Fehler" = $w.LastError
                    Bewertung      = if ($w.LastError -match "No error") { "✅ OK" } else { "⚠️ $($w.LastError)" }
                }
            }
            $html += "<h4>VSS Writer</h4>"
            $html += (ConvertTo-HTMLTable -Data $writerData)
        }
        else {
            $html += "<h4>VSS Writer</h4><p class='no-data'>Nicht verfügbar (WinRM erforderlich).</p>"
        }

        # --- Windows Server Backup ---
        $wsb = Invoke-RemoteCommand -ComputerName $server -ScriptBlock {
            try {
                if (Get-Command Get-WBPolicy -ErrorAction SilentlyContinue) {
                    $policy = Get-WBPolicy -ErrorAction SilentlyContinue
                    $summary = Get-WBSummary -ErrorAction SilentlyContinue
                    [PSCustomObject]@{
                        HasPolicy       = [bool]$policy
                        Schedule        = if ($policy) { ($policy.Schedule -join ', ') } else { $null }
                        LastBackup      = if ($summary) { $summary.LastBackupTime } else { $null }
                        LastResult      = if ($summary) { $summary.LastBackupResultHR } else { $null }
                        NextBackup      = if ($summary) { $summary.NextBackupTime } else { $null }
                        NumberOfVersions = if ($summary) { $summary.NumberOfVersions } else { $null }
                    }
                }
            }
            catch { $null }
        }

        if ($wsb) {
            $wsbData = [PSCustomObject]@{
                "Sicherungsrichtlinie vorhanden" = ConvertTo-DisplayValue -Value $wsb.HasPolicy
                "Zeitplan"                       = ConvertTo-DisplayValue -Value $wsb.Schedule
                "Letzte Sicherung"               = ConvertTo-DisplayValue -Value $wsb.LastBackup
                "Letztes Ergebnis"               = ConvertTo-DisplayValue -Value $wsb.LastResult
                "Nächste Sicherung"              = ConvertTo-DisplayValue -Value $wsb.NextBackup
                "Anzahl Versionen"               = ConvertTo-DisplayValue -Value $wsb.NumberOfVersions
            }
            $html += "<h4>Windows Server Backup</h4>"
            $html += (ConvertTo-HTMLTable -Data @($wsbData))
        }
        else {
            $html += "<h4>Windows Server Backup</h4><p class='no-data'>Windows Server Backup ist nicht installiert oder nicht abfragbar.</p>"
        }
    }

    New-HTMLSection -Title "Backup &amp; VSS Writer" -Category "Replikation & Backup" -Content $html
}

# ---------------------------------------------------------------
# 50. HYPER-V ADMINISTRATOREN
# ---------------------------------------------------------------
function Get-HyperVAdminGroupInfo {
    <#
    .SYNOPSIS
        Dokumentiert die Mitglieder der lokalen Gruppen "Administratoren" und
        "Hyper-V-Administratoren".
    #>
    Write-Log -Message "=== Sammle Hyper-V Administratoren ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Best Practice</h3><p>Die Gruppe <strong>Hyper-V-Administratoren</strong> erlaubt die Verwaltung von VMs "
    $html += "ohne volle lokale Administratorrechte. Mitglieder haben jedoch weitreichende Rechte am Host - regelmäßig prüfen.</p></div>"

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"

        $groups = @("Administratoren", "Administrators", "Hyper-V-Administratoren", "Hyper-V Administrators")

        foreach ($groupName in $groups) {
            try {
                $members = Invoke-RemoteCommand -ComputerName $server -ScriptBlock {
                    param($g)
                    try {
                        Get-LocalGroupMember -Group $g -ErrorAction Stop |
                            Select-Object Name, ObjectClass, PrincipalSource
                    }
                    catch { $null }
                } -ArgumentList @($groupName)

                if ($members) {
                    $data = foreach ($m in $members) {
                        [PSCustomObject]@{
                            Mitglied  = $m.Name
                            Typ       = ConvertTo-DisplayValue -Value $m.ObjectClass
                            Quelle    = ConvertTo-DisplayValue -Value $m.PrincipalSource
                        }
                    }
                    $html += "<h4>Gruppe: $groupName</h4>"
                    $html += (ConvertTo-HTMLTable -Data $data)
                }
            }
            catch { }
        }
    }

    New-HTMLSection -Title "Hyper-V Administratoren" -Category "Sicherheit" -Content $html
}

# ---------------------------------------------------------------
# 51. CREDENTIAL GUARD
# ---------------------------------------------------------------
function Get-CredentialGuardInfo {
    <#
    .SYNOPSIS
        Dokumentiert den Status von Credential Guard und Device Guard.
    #>
    Write-Log -Message "=== Sammle Credential Guard Status ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Hinweis</h3><p>Credential Guard auf einem Hyper-V Host kann Nested Virtualization und einige "
    $html += "Verwaltungsszenarien beeinflussen. Auf Hosts wird es üblicherweise nicht aktiviert, in den Gästen hingegen schon.</p></div>"

    $data = foreach ($server in $HyperVServers) {
        try {
            $lsaCfg = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SYSTEM\CurrentControlSet\Control\LSA" -ValueName "LsaCfgFlags"
            $vbsReq = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SYSTEM\CurrentControlSet\Control\DeviceGuard" -ValueName "EnableVirtualizationBasedSecurity"
            $lsaCfgPolicy = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SOFTWARE\Policies\Microsoft\Windows\DeviceGuard" -ValueName "LsaCfgFlags"

            [PSCustomObject]@{
                Host                          = $server
                "LsaCfgFlags (Registry)"      = ConvertTo-DisplayValue -Value $lsaCfg
                "LsaCfgFlags (Richtlinie)"    = ConvertTo-DisplayValue -Value $lsaCfgPolicy
                "VBS aktiviert (Registry)"    = ConvertTo-DisplayValue -Value $vbsReq
                "Credential Guard Status"     = switch ("$lsaCfg") {
                    "0" { "Deaktiviert" }
                    "1" { "✅ Aktiviert mit UEFI-Sperre" }
                    "2" { "✅ Aktiviert ohne UEFI-Sperre" }
                    default { "Nicht konfiguriert" }
                }
            }
        }
        catch {
            [PSCustomObject]@{
                Host = $server; "LsaCfgFlags (Registry)" = "Fehler"; "LsaCfgFlags (Richtlinie)" = "-"
                "VBS aktiviert (Registry)" = "-"; "Credential Guard Status" = "❌ $($_.Exception.Message)"
            }
        }
    }

    $html += (ConvertTo-HTMLTable -Data $data)
    New-HTMLSection -Title "Credential Guard" -Category "Sicherheit" -Content $html
}

# ---------------------------------------------------------------
# 52. HOST GUARDIAN SERVICE
# ---------------------------------------------------------------
function Get-HostGuardianServiceInfo {
    <#
    .SYNOPSIS
        Dokumentiert die HGS-Client-Konfiguration (Guarded Fabric / Shielded VMs).
    #>
    Write-Log -Message "=== Sammle Host Guardian Service Konfiguration ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"

        $hgs = Invoke-RemoteCommand -ComputerName $server -ScriptBlock {
            try {
                if (Get-Command Get-HgsClientConfiguration -ErrorAction SilentlyContinue) {
                    Get-HgsClientConfiguration -ErrorAction SilentlyContinue |
                        Select-Object IsHostGuarded, Mode, KeyProtectionServerUrl, AttestationServerUrl, AttestationStatus, AttestationSubStatus, FallbackConfigured
                }
            }
            catch { $null }
        }

        if ($hgs) {
            $data = [PSCustomObject]@{
                "Host ist Guarded"        = ConvertTo-DisplayValue -Value $hgs.IsHostGuarded
                "Modus"                   = ConvertTo-DisplayValue -Value $hgs.Mode
                "Key Protection Server"   = ConvertTo-DisplayValue -Value $hgs.KeyProtectionServerUrl
                "Attestation Server"      = ConvertTo-DisplayValue -Value $hgs.AttestationServerUrl
                "Attestation Status"      = ConvertTo-DisplayValue -Value $hgs.AttestationStatus
                "Attestation Substatus"   = ConvertTo-DisplayValue -Value $hgs.AttestationSubStatus
                "Fallback konfiguriert"   = ConvertTo-DisplayValue -Value $hgs.FallbackConfigured
            }
            $html += (ConvertTo-HTMLTable -Data @($data))
        }
        else {
            $html += "<p class='no-data'>Host Guardian Service Client ist nicht installiert bzw. nicht konfiguriert (kein Guarded Fabric).</p>"
        }
    }

    New-HTMLSection -Title "Host Guardian Service (Guarded Fabric)" -Category "Sicherheit" -Content $html
}

# ---------------------------------------------------------------
# 53. KERBEROS-DELEGIERUNG
# ---------------------------------------------------------------
function Get-KerberosDelegationInfo {
    <#
    .SYNOPSIS
        Prüft die eingeschränkte Kerberos-Delegierung der Hostcomputerkonten,
        die für Live Migration mit Kerberos benötigt wird.
    #>
    Write-Log -Message "=== Sammle Kerberos-Delegierung ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Hinweis</h3><p>Für Live Migration mit Kerberos-Authentifizierung muss auf jedem Hostcomputerkonto "
    $html += "die eingeschränkte Delegierung für die Dienste <strong>cifs</strong> und <strong>Microsoft Virtual System Migration Service</strong> "
    $html += "zu allen anderen Hosts konfiguriert sein.</p></div>"

    if (-not (Test-CommandAvailable -Name "Get-ADComputer")) {
        $html += "<p class='no-data'>Das ActiveDirectory-Modul ist nicht verfügbar - Delegierung konnte nicht geprüft werden.</p>"
        New-HTMLSection -Title "Kerberos-Delegierung (Live Migration)" -Category "Sicherheit" -Content $html
        return
    }

    $data = foreach ($server in $HyperVServers) {
        try {
            $shortName = ($server -split '\.')[0]
            $comp = Get-ADComputer -Identity $shortName -Properties msDS-AllowedToDelegateTo, TrustedForDelegation, TrustedToAuthForDelegation, servicePrincipalName -ErrorAction Stop

            [PSCustomObject]@{
                Host                          = $shortName
                "Delegierung zu"              = ConvertTo-DisplayValue -Value $comp.'msDS-AllowedToDelegateTo' -EmptyText "Keine (nur CredSSP möglich)"
                "Uneingeschränkte Delegierung" = ConvertTo-DisplayValue -Value $comp.TrustedForDelegation
                "Protokollübergang"           = ConvertTo-DisplayValue -Value $comp.TrustedToAuthForDelegation
                Bewertung                     = if ($comp.TrustedForDelegation) { "⚠️ Uneingeschränkte Delegierung (Sicherheitsrisiko)" } elseif ($comp.'msDS-AllowedToDelegateTo') { "✅ Eingeschränkte Delegierung konfiguriert" } else { "ℹ️ Keine Delegierung konfiguriert" }
            }
        }
        catch {
            [PSCustomObject]@{
                Host = $server; "Delegierung zu" = "Fehler"; "Uneingeschränkte Delegierung" = "-"
                "Protokollübergang" = "-"; Bewertung = "❌ $($_.Exception.Message)"
            }
        }
    }

    $html += (ConvertTo-HTMLTable -Data $data)
    New-HTMLSection -Title "Kerberos-Delegierung (Live Migration)" -Category "Sicherheit" -Content $html
}

# ---------------------------------------------------------------
# 54. ANTIVIRUS-AUSSCHLÜSSE
# ---------------------------------------------------------------
function Get-AntivirusExclusionsInfo {
    <#
    .SYNOPSIS
        Dokumentiert Microsoft Defender Ausschlüsse und prüft die von Microsoft
        empfohlenen Hyper-V Ausschlüsse.
    #>
    Write-Log -Message "=== Sammle Antivirus-Ausschlüsse ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Empfohlene Hyper-V Ausschlüsse</h3><p>Prozesse: <code>vmms.exe</code>, <code>vmwp.exe</code>, <code>vmcompute.exe</code>, <code>vmsp.exe</code><br>"
    $html += "Dateitypen: <code>.vhd</code>, <code>.vhdx</code>, <code>.avhd</code>, <code>.avhdx</code>, <code>.vsv</code>, <code>.iso</code>, <code>.rct</code>, <code>.vmcx</code>, <code>.vmrs</code><br>"
    $html += "Verzeichnisse: VM-Konfigurationspfad, VHD-Pfad, Cluster Shared Volumes, Snapshot-Pfad</p></div>"

    $recommendedProcesses = @("vmms.exe", "vmwp.exe", "vmcompute.exe", "vmsp.exe")
    $recommendedExtensions = @(".vhd", ".vhdx", ".avhd", ".avhdx", ".vsv", ".iso", ".rct", ".vmcx", ".vmrs")

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"

        $mp = Invoke-RemoteCommand -ComputerName $server -ScriptBlock {
            try {
                if (Get-Command Get-MpPreference -ErrorAction SilentlyContinue) {
                    $p = Get-MpPreference -ErrorAction SilentlyContinue
                    $s = Get-MpComputerStatus -ErrorAction SilentlyContinue
                    [PSCustomObject]@{
                        ExclusionPath      = $p.ExclusionPath
                        ExclusionExtension = $p.ExclusionExtension
                        ExclusionProcess   = $p.ExclusionProcess
                        RealTimeProtection = if ($s) { $s.RealTimeProtectionEnabled } else { $null }
                        AntivirusEnabled   = if ($s) { $s.AntivirusEnabled } else { $null }
                        Signatur           = if ($s) { $s.AntivirusSignatureVersion } else { $null }
                        SignaturDatum      = if ($s) { $s.AntivirusSignatureLastUpdated } else { $null }
                    }
                }
            }
            catch { $null }
        }

        if ($mp) {
            $status = [PSCustomObject]@{
                "Antivirus aktiviert"   = ConvertTo-DisplayValue -Value $mp.AntivirusEnabled
                "Echtzeitschutz"        = ConvertTo-DisplayValue -Value $mp.RealTimeProtection
                "Signaturversion"       = ConvertTo-DisplayValue -Value $mp.Signatur
                "Signatur aktualisiert" = ConvertTo-DisplayValue -Value $mp.SignaturDatum
            }
            $html += "<h4>Microsoft Defender Status</h4>"
            $html += (ConvertTo-HTMLTable -Data @($status))

            $pathData = foreach ($p in @($mp.ExclusionPath)) {
                if ($p) { [PSCustomObject]@{ "Ausgeschlossener Pfad" = $p } }
            }
            $html += "<h4>Pfad-Ausschlüsse</h4>"
            $html += (ConvertTo-HTMLTable -Data $pathData -NoDataMessage "Keine Pfad-Ausschlüsse konfiguriert.")

            $extData = foreach ($e in $recommendedExtensions) {
                [PSCustomObject]@{
                    Dateityp   = $e
                    Status     = if (@($mp.ExclusionExtension) -contains $e.TrimStart('.') -or @($mp.ExclusionExtension) -contains $e) { "✅ Ausgeschlossen" } else { "⚠️ Fehlt" }
                }
            }
            $html += "<h4>Empfohlene Dateityp-Ausschlüsse</h4>"
            $html += (ConvertTo-HTMLTable -Data $extData)

            $procData = foreach ($p in $recommendedProcesses) {
                [PSCustomObject]@{
                    Prozess    = $p
                    Status     = if (@($mp.ExclusionProcess) -match [regex]::Escape($p)) { "✅ Ausgeschlossen" } else { "⚠️ Fehlt" }
                }
            }
            $html += "<h4>Empfohlene Prozess-Ausschlüsse</h4>"
            $html += (ConvertTo-HTMLTable -Data $procData)
        }
        else {
            $html += "<p class='no-data'>Microsoft Defender ist nicht installiert oder nicht abfragbar (Drittanbieter-AV möglich).</p>"
        }

        # --- Installierte Antivirenlösung (SecurityCenter2) ---
        try {
            $cimSession = New-ServerCimSession -ComputerName $server
            if ($cimSession) {
                $av = Get-CimInstance -CimSession $cimSession -Namespace "root\SecurityCenter2" -ClassName AntiVirusProduct -ErrorAction SilentlyContinue
                $avData = foreach ($a in $av) {
                    [PSCustomObject]@{
                        Produkt        = $a.displayName
                        Pfad           = ConvertTo-DisplayValue -Value $a.pathToSignedProductExe
                        "Status (hex)" = ("0x{0:X}" -f $a.productState)
                    }
                }
                $html += "<h4>Installierte Antivirenprodukte</h4>"
                $html += (ConvertTo-HTMLTable -Data $avData -NoDataMessage "Keine Produkte über SecurityCenter2 registriert (bei Server-Betriebssystemen normal).")
                Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
            }
        }
        catch { }
    }

    New-HTMLSection -Title "Antivirus-Ausschlüsse (Hyper-V)" -Category "Sicherheit" -Content $html
}

# ---------------------------------------------------------------
# 55. FIREWALL
# ---------------------------------------------------------------
function Get-FirewallInfo {
    <#
    .SYNOPSIS
        Dokumentiert Firewallprofile und Hyper-V relevante Firewallregeln.
    #>
    Write-Log -Message "=== Sammle Firewall-Konfiguration ===" -Level "INFO"

    $html = ""

    foreach ($server in $HyperVServers) {
        $html += "<h3 class='server-break'>Host: $server</h3>"

        $fw = Invoke-RemoteCommand -ComputerName $server -ScriptBlock {
            $r = @{}
            try { $r.Profiles = Get-NetFirewallProfile -ErrorAction SilentlyContinue | Select-Object Name, Enabled, DefaultInboundAction, DefaultOutboundAction, LogFileName } catch { }
            try {
                $r.Rules = Get-NetFirewallRule -ErrorAction SilentlyContinue |
                    Where-Object { $_.Enabled -eq 'True' -and ($_.DisplayGroup -match 'Hyper-V|Failover|Remote|Windows-Verwaltungsinstrumentation|Windows Management Instrumentation|Datei- und Druckerfreigabe|File and Printer Sharing|Kerncomputer|Core Networking') } |
                    Select-Object DisplayName, DisplayGroup, Direction, Action, Profile
            } catch { }
            [PSCustomObject]$r
        }

        if (-not $fw) {
            $html += "<p class='no-data'>Firewall-Informationen konnten nicht abgefragt werden (WinRM erforderlich).</p>"
            continue
        }

        $profileData = foreach ($p in $fw.Profiles) {
            [PSCustomObject]@{
                Profil                = $p.Name
                Aktiviert             = ConvertTo-DisplayValue -Value $p.Enabled
                "Eingehend (Standard)" = ConvertTo-DisplayValue -Value $p.DefaultInboundAction
                "Ausgehend (Standard)" = ConvertTo-DisplayValue -Value $p.DefaultOutboundAction
                Protokolldatei        = ConvertTo-DisplayValue -Value $p.LogFileName
            }
        }
        $html += "<h4>Firewall-Profile</h4>"
        $html += (ConvertTo-HTMLTable -Data $profileData -NoDataMessage "Nicht verfügbar.")

        $ruleData = foreach ($r in $fw.Rules) {
            [PSCustomObject]@{
                Regel     = $r.DisplayName
                Gruppe    = ConvertTo-DisplayValue -Value $r.DisplayGroup
                Richtung  = ConvertTo-DisplayValue -Value $r.Direction
                Aktion    = ConvertTo-DisplayValue -Value $r.Action
                Profil    = ConvertTo-DisplayValue -Value $r.Profile
            }
        }
        $html += "<h4>Relevante aktive Firewall-Regeln</h4>"
        $html += (ConvertTo-HTMLTable -Data $ruleData -NoDataMessage "Keine relevanten Regeln gefunden.")
    }

    New-HTMLSection -Title "Firewall-Konfiguration" -Category "Sicherheit" -Content $html
}

# ---------------------------------------------------------------
# 56. SMBv1 STATUS
# ---------------------------------------------------------------
function Get-SMBv1StatusInfo {
    <#
    .SYNOPSIS
        Prüft, ob das unsichere SMBv1-Protokoll aktiviert ist.
    #>
    Write-Log -Message "=== Prüfe SMBv1 Status ===" -Level "INFO"

    $html = "<div class='summary-box'><h3>Best Practice</h3><p>SMBv1 gilt als unsicher und sollte auf Hyper-V Hosts deaktiviert sein.</p></div>"

    $data = foreach ($server in $HyperVServers) {
        try {
            $reg = Get-RemoteRegistryValue -ComputerName $server -RegistryPath "SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" -ValueName "SMB1"
            $smbState = Invoke-RemoteCommand -ComputerName $server -ScriptBlock {
                try {
                    $cfg = Get-SmbServerConfiguration -ErrorAction SilentlyContinue
                    $feature = Get-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -ErrorAction SilentlyContinue
                    [PSCustomObject]@{
                        ServerEnabled  = if ($cfg) { $cfg.EnableSMB1Protocol } else { $null }
                        FeatureState   = if ($feature) { $feature.State } else { $null }
                    }
                }
                catch { $null }
            }

            [PSCustomObject]@{
                Host                     = $server
                "Registry SMB1"          = ConvertTo-DisplayValue -Value $reg -EmptyText "Nicht gesetzt (Standard)"
                "SMB Server Protokoll"   = if ($smbState) { ConvertTo-DisplayValue -Value $smbState.ServerEnabled } else { "-" }
                "Windows Feature"        = if ($smbState) { ConvertTo-DisplayValue -Value $smbState.FeatureState } else { "-" }
                Bewertung                = if (($smbState -and $smbState.ServerEnabled -eq $true) -or "$reg" -eq "1") { "⚠️ SMBv1 AKTIV - deaktivieren!" } else { "✅ SMBv1 deaktiviert" }
            }
        }
        catch {
            [PSCustomObject]@{
                Host = $server; "Registry SMB1" = "Fehler"; "SMB Server Protokoll" = "-"
                "Windows Feature" = "-"; Bewertung = "❌ $($_.Exception.Message)"
            }
        }
    }

    $html += (ConvertTo-HTMLTable -Data $data)
    New-HTMLSection -Title "SMBv1 Status (Sicherheit)" -Category "Sicherheit" -Content $html
}

# ---------------------------------------------------------------
# 57. ACTIVE DIRECTORY
# ---------------------------------------------------------------
function Get-ADInformation {
    <#
    .SYNOPSIS
        Dokumentiert grundlegende Active Directory Informationen der Umgebung.
    #>
    Write-Log -Message "=== Sammle Active Directory Informationen ===" -Level "INFO"

    if (-not (Test-CommandAvailable -Name "Get-ADDomain")) {
        New-HTMLSection -Title "Active Directory" -Category "Active Directory" -Content "<p class='no-data'>Das ActiveDirectory-Modul ist nicht verfügbar.</p>"
        return
    }

    $html = ""

    try {
        $domain = Get-ADDomain -ErrorAction Stop
        $forest = Get-ADForest -ErrorAction Stop

        $domainData = [PSCustomObject]@{
            "Domäne (DNS)"          = $domain.DNSRoot
            "NetBIOS-Name"          = $domain.NetBIOSName
            "Domänenmodus"          = ConvertTo-DisplayValue -Value $domain.DomainMode
            "Gesamtstruktur"        = $forest.Name
            "Gesamtstrukturmodus"   = ConvertTo-DisplayValue -Value $forest.ForestMode
            "Domänencontroller"     = ConvertTo-DisplayValue -Value (@($domain.ReplicaDirectoryServers) -join ', ')
            "Standorte"             = ConvertTo-DisplayValue -Value (@($forest.Sites) -join ', ')
            "Globale Kataloge"      = ConvertTo-DisplayValue -Value (@($forest.GlobalCatalogs) -join ', ')
        }
        $html += "<h4>Domäne &amp; Gesamtstruktur</h4>"
        $html += (ConvertTo-HTMLTable -Data @($domainData))

        # --- Hostcomputerkonten ---
        $compData = foreach ($server in $HyperVServers) {
            try {
                $shortName = ($server -split '\.')[0]
                $comp = Get-ADComputer -Identity $shortName -Properties OperatingSystem, OperatingSystemVersion, LastLogonDate, whenCreated, DistinguishedName, Enabled -ErrorAction Stop
                [PSCustomObject]@{
                    Host                = $comp.Name
                    "Betriebssystem"    = ConvertTo-DisplayValue -Value $comp.OperatingSystem
                    "OS Version"        = ConvertTo-DisplayValue -Value $comp.OperatingSystemVersion
                    "Aktiviert"         = ConvertTo-DisplayValue -Value $comp.Enabled
                    "Letzte Anmeldung"  = if ($comp.LastLogonDate) { $comp.LastLogonDate.ToString("dd.MM.yyyy HH:mm") } else { "-" }
                    "Erstellt"          = if ($comp.whenCreated) { $comp.whenCreated.ToString("dd.MM.yyyy") } else { "-" }
                    "OU / DN"           = $comp.DistinguishedName
                }
            }
            catch { $null }
        }
        $html += "<h4>Hostcomputerkonten</h4>"
        $html += (ConvertTo-HTMLTable -Data $compData -NoDataMessage "Keine Computerkonten gefunden.")
    }
    catch {
        Write-Log -Message "AD-Abfrage fehlgeschlagen: $_" -Level "WARNING"
        $html += "<p class='error'>Fehler: $($_.Exception.Message)</p>"
    }

    New-HTMLSection -Title "Active Directory" -Category "Active Directory" -Content $html
}

# ---------------------------------------------------------------
# 58. FSMO-ROLLEN
# ---------------------------------------------------------------
function Get-FSMORoles {
    <#
    .SYNOPSIS
        Dokumentiert die FSMO-Rolleninhaber der Domäne und Gesamtstruktur.
    #>
    Write-Log -Message "=== Sammle FSMO-Rollen ===" -Level "INFO"

    if (-not (Test-CommandAvailable -Name "Get-ADDomain")) {
        New-HTMLSection -Title "FSMO-Rollen" -Category "Active Directory" -Content "<p class='no-data'>Das ActiveDirectory-Modul ist nicht verfügbar.</p>"
        return
    }

    try {
        $domain = Get-ADDomain -ErrorAction Stop
        $forest = Get-ADForest -ErrorAction Stop

        $data = @(
            [PSCustomObject]@{ Rolle = "Schema Master";          Ebene = "Gesamtstruktur"; Inhaber = $forest.SchemaMaster }
            [PSCustomObject]@{ Rolle = "Domain Naming Master";   Ebene = "Gesamtstruktur"; Inhaber = $forest.DomainNamingMaster }
            [PSCustomObject]@{ Rolle = "PDC Emulator";           Ebene = "Domäne";         Inhaber = $domain.PDCEmulator }
            [PSCustomObject]@{ Rolle = "RID Master";             Ebene = "Domäne";         Inhaber = $domain.RIDMaster }
            [PSCustomObject]@{ Rolle = "Infrastructure Master";  Ebene = "Domäne";         Inhaber = $domain.InfrastructureMaster }
        )

        $html = "<div class='summary-box'><h3>Hinweis</h3><p>Wenn Domänencontroller als VMs auf diesen Hyper-V Hosts laufen, sollte mindestens ein DC "
        $html += "(idealerweise der PDC-Emulator) auf physischer Hardware oder auf einem anderen Host betrieben werden.</p></div>"
        $html += (ConvertTo-HTMLTable -Data $data)

        New-HTMLSection -Title "FSMO-Rollen" -Category "Active Directory" -Content $html
    }
    catch {
        Write-Log -Message "FSMO-Abfrage fehlgeschlagen: $_" -Level "WARNING"
        New-HTMLSection -Title "FSMO-Rollen" -Category "Active Directory" -Content "<p class='error'>Fehler: $($_.Exception.Message)</p>"
    }
}

#endregion

#region ============================================================
# HTML-DOKUMENT GENERIERUNG
#endregion ============================================================

function Build-HTMLDocument {
    <#
    .SYNOPSIS
        Baut das finale HTML-Dokument mit Deckblatt, Inhaltsverzeichnis und allen Sektionen.
    #>
    Write-Log -Message "=== Erstelle HTML-Dokument ===" -Level "INFO"

    $cssStyle = @"
    <style>
        body {
            font-family: 'Segoe UI', Calibri, Arial, sans-serif;
            font-size: 11pt;
            color: #333333;
            margin: 40px;
            line-height: 1.5;
        }
        .cover-page {
            text-align: center;
            padding: 100px 0;
            page-break-after: always;
            border-bottom: 3px solid #0078D4;
        }
        .cover-page h1 {
            font-size: 28pt;
            color: #0078D4;
            margin-bottom: 10px;
        }
        .cover-page h2 {
            font-size: 18pt;
            color: #555555;
            font-weight: normal;
        }
        .cover-page .meta {
            margin-top: 50px;
            font-size: 12pt;
            color: #777777;
        }
        h2 {
            color: #0078D4;
            font-size: 16pt;
            border-bottom: 2px solid #0078D4;
            padding-bottom: 5px;
            margin-top: 30px;
        }
        h3 {
            color: #333333;
            font-size: 13pt;
            border-bottom: 1px solid #CCCCCC;
            padding-bottom: 3px;
            margin-top: 20px;
        }
        h3.server-break {
            page-break-before: avoid;
        }
        h4 {
            color: #555555;
            font-size: 11pt;
            margin-top: 15px;
        }
        .toc {
            page-break-after: always;
        }
        .toc h2 {
            page-break-before: avoid;
        }
        .toc ul {
            list-style-type: none;
            padding-left: 0;
        }
        .toc ul ul {
            padding-left: 0;
            margin-top: 4px;
        }
        .toc > ul > li {
            font-weight: bold;
            font-size: 10pt;
            color: #333333;
            display: block;
            padding: 10px 18px;
            margin: 8px 0 6px 0;
            border: 1px solid #0078D4;
            border-radius: 6px;
            background-color: #E8F4FD;
        }
        .toc li li {
            padding: 2px 0 2px 8px;
            list-style-type: none;
        }
        .toc li li a {
            text-decoration: none;
            color: #0078D4;
        }
        .toc li li a:hover {
            text-decoration: underline;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 10px 0 20px 0;
            font-size: 9pt;
        }
        th {
            background-color: #0078D4;
            color: white;
            text-align: left;
            padding: 8px 10px;
            font-weight: bold;
            border: 1px solid #005A9E;
        }
        td {
            padding: 6px 10px;
            border: 1px solid #DDDDDD;
            vertical-align: top;
            word-wrap: break-word;
            max-width: 300px;
        }
        tr.even { background-color: #F5F5F5; }
        tr.odd { background-color: #FFFFFF; }
        tr:hover { background-color: #E8F4FD; }
        .section { margin-bottom: 20px; }
        .no-data {
            color: #888888;
            font-style: italic;
            padding: 10px;
            background-color: #F9F9F9;
            border-left: 3px solid #CCCCCC;
        }
        .error {
            color: #D32F2F;
            padding: 10px;
            background-color: #FFEBEE;
            border-left: 3px solid #D32F2F;
        }
        .summary-box {
            background-color: #F0F7FF;
            border: 1px solid #0078D4;
            border-radius: 5px;
            padding: 15px;
            margin: 20px 0;
        }
        .summary-box h3 {
            border: none;
            color: #0078D4;
            margin-top: 0;
        }
        code {
            background-color: #F3F3F3;
            padding: 1px 4px;
            border-radius: 3px;
            font-family: Consolas, 'Courier New', monospace;
            font-size: 9pt;
        }
        .footer {
            text-align: center;
            font-size: 8pt;
            color: #999999;
            border-top: 1px solid #CCCCCC;
            padding-top: 10px;
            margin-top: 40px;
        }
        @media print {
            body { margin: 20px; line-height: 1.4; font-size: 9pt; }
            h1, h2, h3, h4 { orphans: 3; widows: 3; }
            h2 { page-break-before: auto; margin-top: 0; }
            h3.server-break { page-break-inside: avoid; page-break-before: auto; }
            .section { page-break-inside: avoid; margin-bottom: 15px; }
            table { page-break-inside: avoid; margin: 8px 0 15px 0; }
            tr { page-break-inside: avoid; }
            thead { display: table-header-group; }
            tfoot { display: table-footer-group; }
            .toc { page-break-after: always; }
            .cover-page { page-break-after: always; }
        }
    </style>
"@

    # Inhaltsverzeichnis (nach Category + Title sortiert)
    $sortedTOC = $script:TOCEntries | Sort-Object { $_.Category }, { $_.Title }
    $tocHTML = "<div class='toc'>`n<h2>$(Get-T 'Inhaltsverzeichnis')</h2>`n<ul>"
    $lastCategory = ""
    foreach ($entry in $sortedTOC) {
        if ($entry.Category -and $entry.Category -ne $lastCategory) {
            if ($lastCategory) { $tocHTML += "</ul></li>`n" }
            $tocHTML += "<li>$($entry.Category)`n<ul>`n"
            $lastCategory = $entry.Category
        }
        $tocHTML += "<li><a href='#$($entry.Anchor)'>$($entry.Title)</a></li>`n"
    }
    if ($lastCategory) { $tocHTML += "</ul></li>`n" }
    $tocHTML += "</ul></div>"

    # Zusammenfassung
    $summaryHTML = @"
    <div class="summary-box">
        <h3>$(Get-T 'Dokumentations-Zusammenfassung')</h3>
        <p><strong>$(Get-T 'Dokumentierte Hosts'):</strong> $($HyperVServers -join ', ')</p>
        <p><strong>$(Get-T 'Host-Betriebssystem'):</strong> $script:HyperVEdition</p>
        <p><strong>$(Get-T 'Failover Cluster'):</strong> $(if ($script:ClusterName) { $script:ClusterName } else { Get-T "Kein Cluster (Standalone)" })</p>
        <p><strong>$(Get-T 'Erstellt am'):</strong> $script:Timestamp</p>
        <p><strong>$(Get-T 'Erstellt von'):</strong> $script:DocAuthor $(Get-T 'auf') $script:DocComputerName</p>
        <p><strong>$(Get-T 'Sektionen'):</strong> $($script:SectionCounter)</p>
        <p><strong>$(Get-T 'Fehler'):</strong> $($script:ErrorCount)</p>
        <p><strong>$(Get-T 'Warnungen'):</strong> $($script:WarningCount)</p>
    </div>
"@

    # Gesamtes HTML-Dokument
    $fullHTML = @"
<!DOCTYPE html>
<html lang="$(if ($script:Language -eq 'EN') { 'en' } else { 'de' })">
<head>
    <meta charset="UTF-8">
    <meta name="author" content="$script:DocAuthor">
    <meta name="generator" content="Hyper-V Documentation Script v$script:ScriptVersion">
    <title>$script:DocTitle - $script:DocSubTitle</title>
    $cssStyle
</head>
<body>

    <div class="cover-page">
        <h1>$script:DocTitle</h1>
        <h2>$script:DocSubTitle</h2>
        <div class="meta">
            <p>$(Get-T 'Erstellt am'): $script:DateOnly</p>
            <p>$(Get-T 'Erstellt von'): $script:DocAuthor</p>
            <p>$(Get-T 'Dokumentierte Hosts'): $($HyperVServers -join ', ')</p>
            <p>Version: $script:ScriptVersion (Hyper-V / CIM-DCOM Fallback)</p>
        </div>
    </div>

    $tocHTML

    $summaryHTML

    $($script:HTMLSections -join "`n`n")

    <div class="footer">
        <p>$script:DocTitle | $script:DocSubTitle | $(Get-T 'Erstellt am'): $script:Timestamp | PowerShell Documentation v$script:ScriptVersion</p>
    </div>

</body>
</html>
"@

    return $fullHTML
}

#endregion

#region ============================================================
# SEKTIONS-REGISTRY, MARKDOWN, EXPORT & GUI
#endregion ============================================================

function Get-DocSectionRegistry {
    <#
    .SYNOPSIS
        Liefert die geordnete Liste aller verfügbaren Dokumentations-Sektionen
        (Schlüssel, Anzeigename, Kategorie, zugehörige Funktion). Basis für GUI und Runner.
    #>
    return @(
        # === HARDWARE & OS ===
        [PSCustomObject]@{ Key = "Hardware";          Label = "Hardware & Host-Details";              Category = "Hardware & OS";          Function = "Get-HardwareInformation" }
        [PSCustomObject]@{ Key = "CpuVirtualization"; Label = "Virtualisierungs-Unterstützung (SLAT)"; Category = "Hardware & OS";         Function = "Get-CpuVirtualizationSupportInfo" }
        [PSCustomObject]@{ Key = "Numa";              Label = "NUMA-Topologie & Spanning";            Category = "Hardware & OS";          Function = "Get-NumaTopologyInfo" }
        [PSCustomObject]@{ Key = "WindowsFeatures";   Label = "Windows Features & Rollen";            Category = "Hardware & OS";          Function = "Get-WindowsFeaturesInfo" }
        [PSCustomObject]@{ Key = "Software";          Label = "Installierte Software";                Category = "Hardware & OS";          Function = "Get-InstalledSoftwareInfo" }
        [PSCustomObject]@{ Key = "PowerPlan";         Label = "Power Plan & Performance";             Category = "Hardware & OS";          Function = "Get-PowerPlanInfo" }
        [PSCustomObject]@{ Key = "PendingReboot";     Label = "Ausstehende Neustarts";                Category = "Hardware & OS";          Function = "Get-PendingRebootInfo" }
        [PSCustomObject]@{ Key = "Patch";             Label = "Windows Updates & Patch-Stand";        Category = "Hardware & OS";          Function = "Get-PatchInformation" }
        [PSCustomObject]@{ Key = "DiskSpace";         Label = "Speicherplatz & VM-Pfade";             Category = "Hardware & OS";          Function = "Get-DiskSpaceInfo" }
        [PSCustomObject]@{ Key = "EventLog";          Label = "Event Logs (7 Tage)";                  Category = "Hardware & OS";          Function = "Get-EventLogInfo" }
        [PSCustomObject]@{ Key = "TimeSync";          Label = "Zeitsynchronisierung (w32tm)";         Category = "Hardware & OS";          Function = "Get-TimeSyncInfo" }
        [PSCustomObject]@{ Key = "Licensing";         Label = "Lizenzierung & Aktivierung";           Category = "Hardware & OS";          Function = "Get-LicensingInfo" }

        # === HYPER-V HOST ===
        [PSCustomObject]@{ Key = "VMHost";            Label = "Hyper-V Host-Konfiguration";           Category = "Hyper-V Host";           Function = "Get-VMHostConfigurationInfo" }
        [PSCustomObject]@{ Key = "HostPaths";         Label = "Standardpfade (VHD & VM)";             Category = "Hyper-V Host";           Function = "Get-VMHostPathsInfo" }
        [PSCustomObject]@{ Key = "LiveMigration";     Label = "Live Migration & Storage Migration";   Category = "Hyper-V Host";           Function = "Get-LiveMigrationInfo" }
        [PSCustomObject]@{ Key = "EnhancedSession";   Label = "Enhanced Session & Metering";          Category = "Hyper-V Host";           Function = "Get-EnhancedSessionModeInfo" }
        [PSCustomObject]@{ Key = "HyperVServices";    Label = "Hyper-V Dienststatus";                 Category = "Hyper-V Host";           Function = "Get-HyperVServiceStatusInfo" }
        [PSCustomObject]@{ Key = "VMVersions";        Label = "VM-Konfigurationsversionen";           Category = "Hyper-V Host";           Function = "Get-HyperVVersionInfo" }
        [PSCustomObject]@{ Key = "HostUsage";         Label = "Ressourcenauslastung & Überbuchung";   Category = "Hyper-V Host";           Function = "Get-HostResourceUsageInfo" }

        # === VIRTUELLE MASCHINEN ===
        [PSCustomObject]@{ Key = "VMOverview";        Label = "VM-Übersicht";                         Category = "Virtuelle Maschinen";    Function = "Get-VMOverviewInfo" }
        [PSCustomObject]@{ Key = "VMProcessor";       Label = "VM Prozessor-Konfiguration";           Category = "Virtuelle Maschinen";    Function = "Get-VMProcessorInfo" }
        [PSCustomObject]@{ Key = "VMMemory";          Label = "VM Arbeitsspeicher & Dynamic Memory";  Category = "Virtuelle Maschinen";    Function = "Get-VMMemoryInfo" }
        [PSCustomObject]@{ Key = "VMStorage";         Label = "VM Speicher (VHD/VHDX)";               Category = "Virtuelle Maschinen";    Function = "Get-VMStorageInfo" }
        [PSCustomObject]@{ Key = "VMIntegration";     Label = "Integrationsdienste";                  Category = "Virtuelle Maschinen";    Function = "Get-VMIntegrationServicesInfo" }
        [PSCustomObject]@{ Key = "VMCheckpoints";     Label = "Prüfpunkte (Checkpoints)";             Category = "Virtuelle Maschinen";    Function = "Get-VMCheckpointInfo" }
        [PSCustomObject]@{ Key = "VMFirmware";        Label = "Firmware / BIOS & Secure Boot";        Category = "Virtuelle Maschinen";    Function = "Get-VMFirmwareInfo" }
        [PSCustomObject]@{ Key = "VMAutomatic";       Label = "Automatische Start-/Stop-Aktionen";    Category = "Virtuelle Maschinen";    Function = "Get-VMAutomaticActionsInfo" }
        [PSCustomObject]@{ Key = "VMController";      Label = "Controller, DVD & COM-Ports";          Category = "Virtuelle Maschinen";    Function = "Get-VMControllerInfo" }
        [PSCustomObject]@{ Key = "VMGroups";          Label = "VM-Gruppen";                           Category = "Virtuelle Maschinen";    Function = "Get-VMGroupInfo" }
        [PSCustomObject]@{ Key = "VMGuestOS";         Label = "Gast-Betriebssysteme (KVP)";           Category = "Virtuelle Maschinen";    Function = "Get-VMGuestOperatingSystemInfo" }
        [PSCustomObject]@{ Key = "VMMetering";        Label = "Ressourcenmessung (Metering)";         Category = "Virtuelle Maschinen";    Function = "Get-VMResourceMeteringInfo" }

        # === VIRTUELLE NETZWERKE ===
        [PSCustomObject]@{ Key = "VMSwitch";          Label = "Virtuelle Switches";                   Category = "Virtuelle Netzwerke";    Function = "Get-VMSwitchInfo" }
        [PSCustomObject]@{ Key = "VMSwitchExt";       Label = "Virtual Switch Extensions";            Category = "Virtuelle Netzwerke";    Function = "Get-VMSwitchExtensionInfo" }
        [PSCustomObject]@{ Key = "VMNetwork";         Label = "VM Netzwerkadapter";                   Category = "Virtuelle Netzwerke";    Function = "Get-VMNetworkAdapterInfo" }
        [PSCustomObject]@{ Key = "VMNetworkAdvanced"; Label = "VLAN, Portsicherheit & ACLs";          Category = "Virtuelle Netzwerke";    Function = "Get-VMNetworkAdvancedInfo" }
        [PSCustomObject]@{ Key = "HostNIC";           Label = "Host-NICs (VMQ/RSS/SR-IOV/RDMA)";      Category = "Virtuelle Netzwerke";    Function = "Get-HostNetworkAdapterInfo" }
        [PSCustomObject]@{ Key = "NetQoS";            Label = "Netzwerk-QoS / DCB";                   Category = "Virtuelle Netzwerke";    Function = "Get-NetworkQoSInfo" }

        # === SPEICHER ===
        [PSCustomObject]@{ Key = "Storage";           Label = "Storage-Konfiguration";                Category = "Speicher";               Function = "Get-StorageConfigurationInfo" }
        [PSCustomObject]@{ Key = "VHDDetails";        Label = "VHD-Analyse & verwaiste Dateien";      Category = "Speicher";               Function = "Get-VHDDetailsInfo" }
        [PSCustomObject]@{ Key = "StorageQoS";        Label = "Storage QoS Policies";                 Category = "Speicher";               Function = "Get-StorageQoSInfo" }
        [PSCustomObject]@{ Key = "SMBStorage";        Label = "SMB 3.0 Storage";                      Category = "Speicher";               Function = "Get-SMBStorageInfo" }
        [PSCustomObject]@{ Key = "MPIO";              Label = "MPIO (Multipath I/O)";                 Category = "Speicher";               Function = "Get-MPIOInfo" }

        # === FAILOVER CLUSTER ===
        [PSCustomObject]@{ Key = "Cluster";           Label = "Failover Cluster";                     Category = "Failover Cluster";       Function = "Get-FailoverClusterInfo" }
        [PSCustomObject]@{ Key = "ClusterNetwork";    Label = "Cluster-Netzwerke";                    Category = "Failover Cluster";       Function = "Get-ClusterNetworkInfo" }
        [PSCustomObject]@{ Key = "CSV";               Label = "Cluster Shared Volumes (CSV)";         Category = "Failover Cluster";       Function = "Get-ClusterSharedVolumeInfo" }
        [PSCustomObject]@{ Key = "ClusterQuorum";     Label = "Cluster-Quorum";                       Category = "Failover Cluster";       Function = "Get-ClusterQuorumInfo" }

        # === REPLIKATION & BACKUP ===
        [PSCustomObject]@{ Key = "Replication";       Label = "Hyper-V Replica";                      Category = "Replikation & Backup";   Function = "Get-VMReplicationInfo" }
        [PSCustomObject]@{ Key = "Backup";            Label = "Backup & VSS Writer";                  Category = "Replikation & Backup";   Function = "Get-BackupConfigurationInfo" }

        # === SICHERHEIT ===
        [PSCustomObject]@{ Key = "VMSecurity";        Label = "VM Sicherheit (vTPM / Shielded)";      Category = "Sicherheit";             Function = "Get-VMSecurityInfo" }
        [PSCustomObject]@{ Key = "HyperVAdmins";      Label = "Hyper-V Administratoren";              Category = "Sicherheit";             Function = "Get-HyperVAdminGroupInfo" }
        [PSCustomObject]@{ Key = "CredentialGuard";   Label = "Credential Guard";                     Category = "Sicherheit";             Function = "Get-CredentialGuardInfo" }
        [PSCustomObject]@{ Key = "HGS";               Label = "Host Guardian Service";                Category = "Sicherheit";             Function = "Get-HostGuardianServiceInfo" }
        [PSCustomObject]@{ Key = "Delegation";        Label = "Kerberos-Delegierung";                 Category = "Sicherheit";             Function = "Get-KerberosDelegationInfo" }
        [PSCustomObject]@{ Key = "AntivirusExclusions"; Label = "Antivirus-Ausschlüsse";              Category = "Sicherheit";             Function = "Get-AntivirusExclusionsInfo" }
        [PSCustomObject]@{ Key = "Firewall";          Label = "Firewall-Konfiguration";               Category = "Sicherheit";             Function = "Get-FirewallInfo" }
        [PSCustomObject]@{ Key = "SMBv1";             Label = "SMBv1 Status (Sicherheit)";            Category = "Sicherheit";             Function = "Get-SMBv1StatusInfo" }

        # === ACTIVE DIRECTORY ===
        [PSCustomObject]@{ Key = "AD";                Label = "Active Directory";                     Category = "Active Directory";       Function = "Get-ADInformation" }
        [PSCustomObject]@{ Key = "FSMO";              Label = "FSMO-Rollen";                          Category = "Active Directory";       Function = "Get-FSMORoles" }
    )
}

function ConvertTo-MarkdownFromHtml {
    <#
    .SYNOPSIS
        Konvertiert die im Skript verwendeten HTML-Fragmente in Markdown.
    .PARAMETER Html
        Der zu konvertierende HTML-String.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Html
    )

    $md = $Html

    # Tabellen separat behandeln
    $tableRegex = [regex]'(?s)<table.*?>(.*?)</table>'
    $md = $tableRegex.Replace($md, {
        param($m)
        $tableContent = $m.Groups[1].Value
        $rowRegex = [regex]'(?s)<tr.*?>(.*?)</tr>'
        $rows = $rowRegex.Matches($tableContent)
        $sb = [System.Text.StringBuilder]::new()
        [void]$sb.AppendLine()
        $rowIndex = 0
        foreach ($row in $rows) {
            $cellRegex = [regex]'(?s)<t[hd].*?>(.*?)</t[hd]>'
            $cells = $cellRegex.Matches($row.Groups[1].Value)
            $cellTexts = foreach ($c in $cells) {
                ($c.Groups[1].Value -replace '<.*?>', '' -replace '\|', '\|').Trim()
            }
            if ($cellTexts.Count -gt 0) {
                [void]$sb.AppendLine("| " + ($cellTexts -join " | ") + " |")
                if ($rowIndex -eq 0) {
                    [void]$sb.AppendLine("| " + (($cellTexts | ForEach-Object { "---" }) -join " | ") + " |")
                }
                $rowIndex++
            }
        }
        [void]$sb.AppendLine()
        return $sb.ToString()
    })

    # Überschriften
    $md = $md -replace '(?s)<h1[^>]*>(.*?)</h1>', "`n# `$1`n"
    $md = $md -replace '(?s)<h2[^>]*>(.*?)</h2>', "`n## `$1`n"
    $md = $md -replace '(?s)<h3[^>]*>(.*?)</h3>', "`n### `$1`n"
    $md = $md -replace '(?s)<h4[^>]*>(.*?)</h4>', "`n#### `$1`n"

    # Inline-Formatierung
    $md = $md -replace '(?s)<strong>(.*?)</strong>', '**$1**'
    $md = $md -replace '(?s)<em>(.*?)</em>', '*$1*'
    $md = $md -replace '(?s)<code>(.*?)</code>', '`$1`'
    $md = $md -replace '<br\s*/?>', "`n"
    $md = $md -replace '(?s)<p[^>]*>(.*?)</p>', "`n`$1`n"

    # Restliche Tags entfernen
    $md = $md -replace '<[^>]+>', ''

    # HTML-Entities dekodieren
    Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
    $md = [System.Web.HttpUtility]::HtmlDecode($md)

    # Mehrfache Leerzeilen reduzieren
    $md = $md -replace '(\r?\n){3,}', "`n`n"

    return $md.Trim()
}

function Build-MarkdownDocument {
    <#
    .SYNOPSIS
        Baut ein Markdown-Dokument aus den gesammelten HTML-Sektionen.
    #>
    Write-Log -Message (Get-T "=== Generiere Markdown-Dokument ===") -Level "INFO"

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("# $script:DocTitle")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("**$script:DocSubTitle**")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("- $(Get-T 'Erstellt am'): $script:DateOnly")
    [void]$sb.AppendLine("- $(Get-T 'Erstellt von'): $script:DocAuthor $(Get-T 'auf') $script:DocComputerName")
    [void]$sb.AppendLine("- $(Get-T 'Dokumentierte Hosts'): $($HyperVServers -join ', ')")
    [void]$sb.AppendLine("- $(Get-T 'Host-Betriebssystem'): $script:HyperVEdition")
    [void]$sb.AppendLine("- $(Get-T 'Failover Cluster'): $(if ($script:ClusterName) { $script:ClusterName } else { Get-T 'Kein Cluster (Standalone)' })")
    [void]$sb.AppendLine("- $(Get-T 'Fehler'): $($script:ErrorCount) | $(Get-T 'Warnungen'): $($script:WarningCount)")
    [void]$sb.AppendLine()

    # Inhaltsverzeichnis (nach Category + Title sortiert)
    [void]$sb.AppendLine("## $(Get-T 'Inhaltsverzeichnis')")
    [void]$sb.AppendLine()
    $sortedTOC = $script:TOCEntries | Sort-Object { $_.Category }, { $_.Title }
    $lastCategory = ""
    foreach ($entry in $sortedTOC) {
        if ($entry.Category -and $entry.Category -ne $lastCategory) {
            [void]$sb.AppendLine("### $($entry.Category)")
            $lastCategory = $entry.Category
        }
        [void]$sb.AppendLine("- **$($entry.Title)**")
    }
    [void]$sb.AppendLine()

    # Sektionen konvertieren
    foreach ($section in $script:HTMLSections) {
        [void]$sb.AppendLine((ConvertTo-MarkdownFromHtml -Html $section))
        [void]$sb.AppendLine()
    }

    return $sb.ToString()
}

function Export-DocumentToPdf {
    <#
    .SYNOPSIS
        Konvertiert eine HTML-Datei in PDF. Nutzt zuerst Microsoft Word (COM),
        bei Nichtverfügbarkeit Microsoft Edge / Chrome im Headless-Modus.
    .PARAMETER HtmlPath
        Pfad zur HTML-Quelldatei.
    .PARAMETER PdfPath
        Ziel-Pfad der PDF-Datei.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string]$HtmlPath,
        [Parameter(Mandatory = $true)] [string]$PdfPath
    )

    # --- Methode 1: Microsoft Word COM ---
    $word = $null
    $doc = $null

    try {
        $word = New-Object -ComObject Word.Application -ErrorAction Stop
        $word.Visible = $false
        $doc = $word.Documents.Open($HtmlPath)
        # 17 = wdFormatPDF
        $doc.SaveAs([ref]$PdfPath, [ref]17)
        $doc.Close($false)
        Write-Log -Message "PDF über Microsoft Word erstellt: $PdfPath" -Level "INFO"
        return $true
    }
    catch {
        Write-Log -Message "Word-COM PDF-Export nicht möglich: $($_.Exception.Message). Versuche Browser-Headless..." -Level "WARNING"
    }
    finally {
        # Cleanup - in separaten Try-Catch um unerwartete Fehler abzufangen
        try {
            if ($doc) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
            }
            if ($word) {
                $word.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
            }
        }
        catch {
            # Fehler beim Cleanup - ignorieren (Word-Prozess läuft trotzdem weiter)
            [void]0
        }
        # Erzwinge Garbage Collection um Word-Prozesse freizugeben
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    # --- Methode 2: Edge / Chrome Headless ---
    $browsers = @(
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
    )
    foreach ($browser in $browsers) {
        if (Test-Path $browser) {
            try {
                $uri = ([System.Uri]$HtmlPath).AbsoluteUri
                $arguments = "--headless --disable-gpu --print-to-pdf=`"$PdfPath`" --no-margins `"$uri`""
                Start-Process -FilePath $browser -ArgumentList $arguments -Wait -NoNewWindow
                if (Test-Path $PdfPath) {
                    Write-Log -Message "PDF über Browser-Headless erstellt: $PdfPath" -Level "INFO"
                    return $true
                }
            }
            catch {
                Write-Log -Message "Browser-Headless PDF-Export fehlgeschlagen ($browser): $_" -Level "WARNING"
            }
        }
    }

    Write-Log -Message "PDF-Export nicht möglich (weder Word noch Edge/Chrome verfügbar)." -Level "ERROR"
    return $false
}

function Show-DocumentationGui {
    <#
    .SYNOPSIS
        Zeigt eine WPF-GUI zur Auswahl von Hosts, Ausgabepfad, Sektionen und
        Ausgabeformaten. Hosts werden automatisch erkannt (lokaler Host + Cluster-Knoten).
    #>
    [CmdletBinding()]
    param(
        [string]$DefaultServers,
        [string]$DefaultCompany,
        [string]$DefaultOutputPath
    )

    try {
        # WPF Assemblies laden
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase -ErrorAction Stop
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    }
    catch {
        Write-Host "FEHLER: WPF-Assemblies konnten nicht geladen werden: $_" -ForegroundColor Red
        return $null
    }

    # --- Hyper-V Hosts automatisch erkennen ---
    $detectedServers = @()
    try {
        if (-not (Get-Module -Name Hyper-V -ErrorAction SilentlyContinue)) {
            if (Get-Module -ListAvailable -Name Hyper-V -ErrorAction SilentlyContinue) {
                Import-Module Hyper-V -ErrorAction SilentlyContinue -DisableNameChecking
            }
        }

        # Lokalen Host prüfen
        try {
            $localHost = Get-VMHost -ErrorAction Stop
            if ($localHost) { $detectedServers += $env:COMPUTERNAME }
        }
        catch { }

        # Cluster-Knoten ergänzen
        try {
            if (Get-Module -ListAvailable -Name FailoverClusters -ErrorAction SilentlyContinue) {
                Import-Module FailoverClusters -ErrorAction SilentlyContinue -DisableNameChecking
                $nodes = Get-ClusterNode -ErrorAction SilentlyContinue
                foreach ($n in $nodes) {
                    if ($detectedServers -notcontains $n.Name) { $detectedServers += $n.Name }
                }
            }
        }
        catch { }
    }
    catch {
        # Keine Hosts gefunden - manuelle Eingabe nötig
    }

    $detectedServers = @($detectedServers | Where-Object { $_ } | Sort-Object -Unique)

    $xamlText = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Hyper-V Dokumentation v1.1" Height="920" Width="1120"
        WindowStartupLocation="CenterScreen" Background="#EDEDF2" ResizeMode="CanResize" MinWidth="900" MinHeight="700"
        FontFamily="Segoe UI, Arial, sans-serif">
    <Window.Resources>
        <!-- Farben -->
        <Color x:Key="PrimaryColor">#1A237E</Color>
        <Color x:Key="PrimaryLight">#3949AB</Color>
        <Color x:Key="AccentColor">#00BCD4</Color>
        <Color x:Key="SuccessColor">#43A047</Color>
        <Color x:Key="ErrorColor">#E53935</Color>

        <SolidColorBrush x:Key="PrimaryBrush" Color="#1A237E"/>
        <SolidColorBrush x:Key="PrimaryLightBrush" Color="#3949AB"/>
        <SolidColorBrush x:Key="PrimaryGradientEnd" Color="#283593"/>
        <SolidColorBrush x:Key="AccentBrush" Color="#00BCD4"/>
        <SolidColorBrush x:Key="SuccessBrush" Color="#43A047"/>
        <SolidColorBrush x:Key="ErrorBrush" Color="#E53935"/>
        <SolidColorBrush x:Key="CardBg" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="PageBg" Color="#EDEDF2"/>
        <SolidColorBrush x:Key="TextPrimary" Color="#212121"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#757575"/>

        <!-- Modernes GroupBox-Style (Card) -->
        <Style x:Key="CardStyle" TargetType="Border">
            <Setter Property="Background" Value="White"/>
            <Setter Property="CornerRadius" Value="10"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="16"/>
            <Setter Property="Margin" Value="0,0,0,14"/>
            <Setter Property="Effect">
                <Setter.Value>
                    <DropShadowEffect BlurRadius="16" ShadowDepth="2" Color="#1A000000" Opacity="0.12"/>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Button Primary Style -->
        <Style x:Key="PrimaryButtonStyle" TargetType="Button">
            <Setter Property="Height" Value="44"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="15"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" CornerRadius="8" Background="#1A237E" BorderThickness="0">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#283593"/>
                                <Setter TargetName="border" Property="Effect">
                                    <Setter.Value>
                                        <DropShadowEffect BlurRadius="12" ShadowDepth="3" Color="#FF1A237E" Opacity="0.4"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#0D154A"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Button Secondary Style -->
        <Style x:Key="SecondaryButtonStyle" TargetType="Button">
            <Setter Property="Height" Value="44"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="BorderThickness" Value="1.5"/>
            <Setter Property="Foreground" Value="#1A237E"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" CornerRadius="8" Background="White" BorderBrush="#1A237E" BorderThickness="1.5">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#F0F0FF"/>
                                <Setter TargetName="border" Property="Effect">
                                    <Setter.Value>
                                        <DropShadowEffect BlurRadius="8" ShadowDepth="2" Color="#1A000000" Opacity="0.1"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#E0E0F0"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Modern CheckBox Style -->
        <Style x:Key="ModernCheckBox" TargetType="CheckBox">
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Foreground" Value="#212121"/>
            <Setter Property="Margin" Value="0,3,0,3"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CheckBox">
                        <BulletDecorator Background="Transparent">
                            <BulletDecorator.Bullet>
                                <Border x:Name="CheckBorder" Width="20" Height="20" CornerRadius="4" Background="White" BorderBrush="#BDBDBD" BorderThickness="1.5">
                                    <Path x:Name="CheckMark" Width="12" Height="12" Stretch="Uniform" Fill="#1A237E" Data="M 0 6 L 4 10 L 10 2" Visibility="Hidden" Margin="0,0,0,0"/>
                                </Border>
                            </BulletDecorator.Bullet>
                            <ContentPresenter Margin="8,0,0,0" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                        </BulletDecorator>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="CheckBorder" Property="Background" Value="#E8EAF6"/>
                                <Setter TargetName="CheckBorder" Property="BorderBrush" Value="#1A237E"/>
                                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="CheckBorder" Property="BorderBrush" Value="#1A237E"/>
                                <Setter TargetName="CheckBorder" Property="Effect">
                                    <Setter.Value>
                                        <DropShadowEffect BlurRadius="6" ShadowDepth="1" Color="#1A000000" Opacity="0.1"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Toggle Button Style für Formate -->
        <Style x:Key="FormatToggleStyle" TargetType="ToggleButton">
            <Setter Property="Height" Value="42"/>
            <Setter Property="Width" Value="110"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Margin" Value="0,0,14,0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ToggleButton">
                        <Border x:Name="border" CornerRadius="8" Background="White" BorderBrush="#BDBDBD" BorderThickness="1.5">
                            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">
                                <TextBlock x:Name="Icon" Text="●" FontSize="10" VerticalAlignment="Center" Margin="0,0,6,0" Foreground="#BDBDBD"/>
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </StackPanel>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#E8EAF6"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#1A237E"/>
                                <Setter TargetName="Icon" Property="Foreground" Value="#1A237E"/>
                                <Setter TargetName="Icon" Property="Text" Value="◆"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="BorderBrush" Value="#3949AB"/>
                                <Setter TargetName="border" Property="Effect">
                                    <Setter.Value>
                                        <DropShadowEffect BlurRadius="8" ShadowDepth="2" Color="#1A000000" Opacity="0.08"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- TextBox Modern Style -->
        <Style x:Key="ModernTextBox" TargetType="TextBox">
            <Setter Property="Height" Value="36"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Padding" Value="10,0"/>
            <Setter Property="BorderThickness" Value="1.5"/>
            <Setter Property="BorderBrush" Value="#BDBDBD"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="border" CornerRadius="6" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="10,0"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="BorderBrush" Value="#3949AB"/>
                            </Trigger>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter TargetName="border" Property="BorderBrush" Value="#1A237E"/>
                                <Setter TargetName="border" Property="Effect">
                                    <Setter.Value>
                                        <DropShadowEffect BlurRadius="8" ShadowDepth="1" Color="#FF1A237E" Opacity="0.15"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ScrollViewer Modern Style -->
        <Style x:Key="ModernScrollViewer" TargetType="ScrollViewer">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollViewer">
                        <Grid>
                            <ScrollContentPresenter Margin="{TemplateBinding Padding}"/>
                            <ScrollBar x:Name="PART_VerticalScrollBar" HorizontalAlignment="Right" Width="8" Visibility="{TemplateBinding ComputedVerticalScrollBarVisibility}">
                                <ScrollBar.Template>
                                    <ControlTemplate TargetType="ScrollBar">
                                        <Grid>
                                            <Track x:Name="PART_Track">
                                                <Track.Thumb>
                                                    <Thumb>
                                                        <Thumb.Template>
                                                            <ControlTemplate TargetType="Thumb">
                                                                <Border CornerRadius="4" Background="#C0C0C0" Margin="0,2"/>
                                                            </ControlTemplate>
                                                        </Thumb.Template>
                                                    </Thumb>
                                                </Track.Thumb>
                                            </Track>
                                        </Grid>
                                    </ControlTemplate>
                                </ScrollBar.Template>
                            </ScrollBar>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="24">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- ===== HEADER ===== -->
        <Border Grid.Row="0" CornerRadius="12" Margin="0,0,0,20" BorderThickness="0">
            <Border.Background>
                <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                    <GradientStop Color="#1A237E" Offset="0.0"/>
                    <GradientStop Color="#283593" Offset="0.5"/>
                    <GradientStop Color="#3949AB" Offset="1.0"/>
                </LinearGradientBrush>
            </Border.Background>
            <Border.Effect>
                <DropShadowEffect BlurRadius="20" ShadowDepth="4" Color="#FF1A237E" Opacity="0.35"/>
            </Border.Effect>
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <!-- Icon Bereich -->
                <Border Grid.Column="0" Width="56" Height="56" CornerRadius="10" Background="#FFFFFF22" Margin="16,16,0,16">
                    <TextBlock Text="🖧" FontSize="28" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
                <!-- Titel -->
                <StackPanel Grid.Column="1" Margin="16,16,0,16" VerticalAlignment="Center">
                    <TextBlock Text="Hyper-V Dokumentation" FontSize="26" FontWeight="SemiBold" Foreground="White"/>
                    <TextBlock Text="Host · VMs · Netzwerke · Speicher · Cluster · Replica · Multi-Format Export" Foreground="#90CAF9" FontSize="13" Margin="0,4,0,0"/>
                </StackPanel>
                <!-- Version Badge & Sprachumschalter -->
                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center" Margin="0,0,16,0">
                    <ToggleButton Name="BtnLangDe" Content="DE" Width="48" Style="{StaticResource FormatToggleStyle}" Margin="0,0,6,0"/>
                    <ToggleButton Name="BtnLangEn" Content="EN" Width="48" Style="{StaticResource FormatToggleStyle}" Margin="0,0,12,0"/>
                    <Border CornerRadius="6" Background="#FFFFFF22" Padding="10,6" VerticalAlignment="Center">
                        <TextBlock Text="v1.1" Foreground="Black" FontWeight="Bold" FontSize="13"/>
                    </Border>
                </StackPanel>
            </Grid>
        </Border>

        <!-- ===== HOST AUSWAHL ===== -->
        <Border Grid.Row="1" Style="{StaticResource CardStyle}">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid Grid.Row="0" Margin="0,0,0,10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0" Orientation="Horizontal">
                        <TextBlock Text="🖥️" FontSize="18" VerticalAlignment="Center" Margin="0,0,10,0"/>
                        <TextBlock Text="Hyper-V Hosts" FontSize="16" FontWeight="SemiBold" Foreground="{StaticResource TextPrimary}" VerticalAlignment="Center"/>
                    </StackPanel>
                    <StackPanel Grid.Column="2" Orientation="Horizontal">
                        <CheckBox Name="ChkSelectAllServers" Content="Alle auswählen" Style="{StaticResource ModernCheckBox}" FontWeight="SemiBold" VerticalAlignment="Center"/>
                    </StackPanel>
                </Grid>
                <Border Grid.Row="1" CornerRadius="8" Background="#F5F5F8" Padding="14,12" MinHeight="40">
                    <Grid>
                        <WrapPanel Grid.Row="0" Name="ServerPanel" Orientation="Horizontal"/>
                        <TextBox Grid.Row="0" Name="TxtServersManual" Height="32" Visibility="Collapsed" VerticalContentAlignment="Center" Margin="0,0,0,0" Style="{StaticResource ModernTextBox}"/>
                    </Grid>
                </Border>
            </Grid>
        </Border>

        <!-- ===== ORGANISATION & AUSGABEPFAD ===== -->
        <Border Grid.Row="2" Style="{StaticResource CardStyle}">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="24"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <!-- Organisation -->
                <StackPanel Grid.Column="0">
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                        <TextBlock Text="🏢" FontSize="16" VerticalAlignment="Center" Margin="0,0,8,0"/>
                        <TextBlock Text="Organisation" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource TextPrimary}" VerticalAlignment="Center"/>
                    </StackPanel>
                    <TextBox Name="TxtCompany" Style="{StaticResource ModernTextBox}"/>
                </StackPanel>
                <!-- Ausgabepfad -->
                <StackPanel Grid.Column="2">
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                        <TextBlock Text="📁" FontSize="16" VerticalAlignment="Center" Margin="0,0,8,0"/>
                        <TextBlock Text="Ausgabepfad" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource TextPrimary}" VerticalAlignment="Center"/>
                    </StackPanel>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox Grid.Column="0" Name="TxtOutputPath" Style="{StaticResource ModernTextBox}"/>
                        <Button Grid.Column="1" Name="BtnBrowse" Content="📂" Width="40" Height="36" Margin="10,0,0,0" Padding="0"
                                FontSize="18" Cursor="Hand" BorderThickness="1.5" BorderBrush="#BDBDBD" Background="White"
                                ToolTip="Ausgabeverzeichnis wählen">
                            <Button.Template>
                                <ControlTemplate TargetType="Button">
                                    <Border x:Name="border" CornerRadius="6" Background="White" BorderBrush="#BDBDBD" BorderThickness="1.5">
                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter TargetName="border" Property="Background" Value="#F5F5FF"/>
                                            <Setter TargetName="border" Property="BorderBrush" Value="#3949AB"/>
                                        </Trigger>
                                        <Trigger Property="IsPressed" Value="True">
                                            <Setter TargetName="border" Property="Background" Value="#E8EAF6"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Button.Template>
                        </Button>
                    </Grid>
                </StackPanel>
            </Grid>
        </Border>

        <!-- ===== AUSGABEFORMATE ===== -->
        <Border Grid.Row="3" Style="{StaticResource CardStyle}">
            <StackPanel>
                <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
                    <TextBlock Text="📄" FontSize="16" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBlock Text="Ausgabeformate" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource TextPrimary}" VerticalAlignment="Center"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <ToggleButton Name="ChkHtml" Content="HTML" Style="{StaticResource FormatToggleStyle}" IsChecked="True"/>
                    <ToggleButton Name="ChkPdf" Content="PDF" Style="{StaticResource FormatToggleStyle}"/>
                    <ToggleButton Name="ChkMarkdown" Content="Markdown" Style="{StaticResource FormatToggleStyle}"/>
                </StackPanel>
            </StackPanel>
        </Border>

        <!-- =Category Header== -->
        <Border Grid.Row="4" Style="{StaticResource CardStyle}" Padding="16,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal">
                    <TextBlock Text="📋" FontSize="16" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBlock Text="Dokumentationsbereiche" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource TextPrimary}" VerticalAlignment="Center"/>
                </StackPanel>
                <CheckBox Grid.Column="2" Name="ChkSelectAll" Content="Alle auswählen" Style="{StaticResource ModernCheckBox}" FontWeight="SemiBold" IsChecked="True"/>
            </Grid>
        </Border>

        <!-- ===== DOKUMENTATIONSBEREICHE ===== -->
        <Border Grid.Row="5" Style="{StaticResource CardStyle}" Padding="0" Margin="0,0,0,14">
            <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Style="{StaticResource ModernScrollViewer}" Padding="16,12">
                <StackPanel Name="CategoryPanel" Orientation="Vertical"/>
            </ScrollViewer>
        </Border>

        <!-- ===== STATUS ===== -->
        <Border Grid.Row="6" CornerRadius="8" Padding="12,10" Margin="0,0,0,12" Background="#FFF3E0" BorderThickness="0" Visibility="Collapsed" Name="StatusBorder">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="ℹ️" FontSize="16" VerticalAlignment="Center" Margin="0,0,10,0"/>
                <TextBlock Grid.Column="1" Name="TxtStatus" Foreground="#E65100" FontWeight="SemiBold" TextWrapping="Wrap" VerticalAlignment="Center"/>
            </Grid>
        </Border>

        <!-- ===== BUTTONS ===== -->
        <StackPanel Grid.Row="7" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,4,0,0">
            <Button Name="BtnCancel" Content="Abbrechen" Width="130" Style="{StaticResource SecondaryButtonStyle}" Margin="0,0,14,0"/>
            <Button Name="BtnStart" Content="  🚀  Dokumentation starten" Width="260" Style="{StaticResource PrimaryButtonStyle}"/>
        </StackPanel>
    </Grid>
</Window>
'@

    # XAML-Texte in die aktive Sprache übersetzen
    if ($script:Language -eq "EN") {
        foreach ($key in @(
            "Hyper-V Dokumentation",
            "Host · VMs · Netzwerke · Speicher · Cluster · Replica · Multi-Format Export",
            "Hyper-V Hosts",
            "Alle auswählen",
            "Organisation",
            "Ausgabepfad",
            "Ausgabeformate",
            "Dokumentationsbereiche",
            "Abbrechen",
            "  🚀  Dokumentation starten"
        )) {
            $xamlText = $xamlText.Replace(">$key<", ">$($script:Dict[$key])<").Replace("=`"$key`"", "=`"$($script:Dict[$key])`"")
        }
    }

    [xml]$xaml = $xamlText

    try {
        $reader = New-Object System.Xml.XmlNodeReader $xaml
        $window = [Windows.Markup.XamlReader]::Load($reader)
    }
    catch {
        Write-Host "FEHLER beim Laden der XAML: $_" -ForegroundColor Red
        return $null
    }

    # Steuerelemente referenzieren
    try {
        $chkSelectAllServers = $window.FindName("ChkSelectAllServers")
        $serverPanel         = $window.FindName("ServerPanel")
        $txtServersManual    = $window.FindName("TxtServersManual")
        $txtCompany          = $window.FindName("TxtCompany")
        $txtOutputPath       = $window.FindName("TxtOutputPath")
        $btnBrowse           = $window.FindName("BtnBrowse")
        $chkHtml             = $window.FindName("ChkHtml")
        $chkPdf              = $window.FindName("ChkPdf")
        $chkMarkdown         = $window.FindName("ChkMarkdown")
        $chkSelectAll        = $window.FindName("ChkSelectAll")
        $categoryPanel       = $window.FindName("CategoryPanel")
        $txtStatus           = $window.FindName("TxtStatus")
        $statusBorder        = $window.FindName("StatusBorder")
        $btnStart            = $window.FindName("BtnStart")
        $btnCancel           = $window.FindName("BtnCancel")
        $btnLangDe           = $window.FindName("BtnLangDe")
        $btnLangEn           = $window.FindName("BtnLangEn")
    }
    catch {
        Write-Host "FEHLER beim Zugriff auf GUI-Steuerelemente: $_" -ForegroundColor Red
        return $null
    }

    # Sprachumschalter: bei Wechsel wird die GUI in der neuen Sprache neu aufgebaut
    $window.Title = "$(Get-T 'Hyper-V Dokumentation') v$script:ScriptVersion"
    $btnLangDe.IsChecked = ($script:Language -eq "DE")
    $btnLangEn.IsChecked = ($script:Language -eq "EN")

    $btnLangDe.Add_Click({
        if ($script:Language -ne "DE") {
            $script:Language = "DE"
            $script:GuiPendingCompany = $txtCompany.Text
            $script:GuiPendingPath    = $txtOutputPath.Text
            $script:GuiRestart = $true
            $window.Close()
        }
        else { $btnLangDe.IsChecked = $true }
    })
    $btnLangEn.Add_Click({
        if ($script:Language -ne "EN") {
            $script:Language = "EN"
            $script:GuiPendingCompany = $txtCompany.Text
            $script:GuiPendingPath    = $txtOutputPath.Text
            $script:GuiRestart = $true
            $window.Close()
        }
        else { $btnLangEn.IsChecked = $true }
    })

    # Vorbelegung
    $txtCompany.Text    = $DefaultCompany
    $txtOutputPath.Text = $DefaultOutputPath

    # Host-Checkboxen
    $serverCheckboxes = @{}
    $modernCheckboxStyle = $window.FindResource("ModernCheckBox")
    if ($detectedServers.Count -gt 0) {
        foreach ($srv in $detectedServers) {
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.Content = $srv
            $cb.Tag = $srv
            $cb.IsChecked = $true
            $cb.Margin = "0,0,20,6"
            $cb.FontSize = 13
            $cb.Style = $modernCheckboxStyle
            [void]$serverPanel.Children.Add($cb)
            $serverCheckboxes[$srv] = $cb
        }
        $chkSelectAllServers.IsChecked = $true
    }
    else {
        $serverPanel.Visibility = "Collapsed"
        $txtServersManual.Visibility = "Visible"
        $txtServersManual.Text = if ($DefaultServers) { $DefaultServers } else { $env:COMPUTERNAME }
        $chkSelectAllServers.Visibility = "Collapsed"
    }

    $chkSelectAllServers.Add_Click({
        $state = $chkSelectAllServers.IsChecked
        foreach ($cb in $serverPanel.Children) {
            if ($cb -is [System.Windows.Controls.CheckBox]) { $cb.IsChecked = $state }
        }
    })

    # Sektions-Checkboxen
    $sectionCheckboxes = @{}
    $registry = Get-DocSectionRegistry
    $categories = $registry | Group-Object -Property Category

    foreach ($cat in $categories) {
        $border = New-Object System.Windows.Controls.Border
        $border.CornerRadius = [System.Windows.CornerRadius]::new(8)
        $border.Background = [System.Windows.Media.Brushes]::White
        $border.BorderThickness = [System.Windows.Thickness]::new(0)
        $border.Margin = [System.Windows.Thickness]::new(0, 0, 0, 12)
        $border.Effect = New-Object System.Windows.Media.Effects.DropShadowEffect -Property @{
            BlurRadius = 12
            ShadowDepth = 1
            Color = [System.Windows.Media.Color]::FromArgb(26, 0, 0, 0)
            Opacity = 0.1
        }

        $outerSp = New-Object System.Windows.Controls.StackPanel
        $outerSp.Orientation = "Vertical"

        # Category Header
        $headerBorder = New-Object System.Windows.Controls.Border
        $headerBorder.CornerRadius = [System.Windows.CornerRadius]::new(8, 8, 0, 0)
        $headerBorder.Background = [System.Windows.Media.Brushes]::White
        $headerBorder.BorderThickness = [System.Windows.Thickness]::new(0)
        $headerBorder.Padding = [System.Windows.Thickness]::new(12, 10, 12, 8)

        $headerGrid = New-Object System.Windows.Controls.Grid
        $headerGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::Auto }))
        $headerGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))

        $iconMap = @{
            "Hardware & OS"        = "🖥️"
            "Hyper-V Host"         = "⚙️"
            "Virtuelle Maschinen"  = "💻"
            "Virtuelle Netzwerke"  = "🌐"
            "Speicher"             = "💾"
            "Failover Cluster"     = "🔗"
            "Replikation & Backup" = "♻️"
            "Sicherheit"           = "🔒"
            "Active Directory"     = "🏛️"
        }
        $catIcon = if ($iconMap.ContainsKey($cat.Name)) { $iconMap[$cat.Name] } else { "📋" }

        $iconText = New-Object System.Windows.Controls.TextBlock
        $iconText.Text = "$catIcon  $(Get-T $cat.Name)"
        $iconText.FontSize = 14
        $iconText.FontWeight = [System.Windows.FontWeights]::SemiBold
        $iconText.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 33, 33, 33))
        $iconText.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        [void]$headerGrid.Children.Add($iconText)

        $headerBorder.Child = $headerGrid
        [void]$outerSp.Children.Add($headerBorder)

        # Separator
        $sep = New-Object System.Windows.Controls.Border
        $sep.Height = 1
        $sep.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(40, 0, 0, 0))
        $sep.Margin = [System.Windows.Thickness]::new(12, 0, 12, 0)
        [void]$outerSp.Children.Add($sep)

        # Content
        $contentBorder = New-Object System.Windows.Controls.Border
        $contentBorder.CornerRadius = [System.Windows.CornerRadius]::new(0, 0, 8, 8)
        $contentBorder.Background = [System.Windows.Media.Brushes]::White
        $contentBorder.Padding = [System.Windows.Thickness]::new(12, 8, 12, 10)

        $sp = New-Object System.Windows.Controls.StackPanel
        $sp.Orientation = "Vertical"

        foreach ($section in $cat.Group) {
            $cb = New-Object System.Windows.Controls.CheckBox
            $cb.Content = Get-T $section.Label
            $cb.Tag = $section.Key
            $cb.IsChecked = $true
            $cb.Margin = "0,3,0,3"
            $cb.FontSize = 12
            $cb.Style = $modernCheckboxStyle
            [void]$sp.Children.Add($cb)
            $sectionCheckboxes[$section.Key] = $cb
        }
        $contentBorder.Child = $sp
        [void]$outerSp.Children.Add($contentBorder)
        $border.Child = $outerSp
        [void]$categoryPanel.Children.Add($border)
    }

    $chkSelectAll.Add_Click({
        $state = $chkSelectAll.IsChecked
        foreach ($key in $sectionCheckboxes.Keys) {
            $sectionCheckboxes[$key].IsChecked = $state
        }
    })

    $btnBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = Get-T "Ausgabeverzeichnis wählen"
        if ($txtOutputPath.Text -and (Test-Path $txtOutputPath.Text)) { $dlg.SelectedPath = $txtOutputPath.Text }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtOutputPath.Text = $dlg.SelectedPath
        }
    })

    $script:GuiResult = $null
    $script:GuiRestart = $false

    # Hilfsfunktion für Status-Anzeige
    $script:ShowStatus = {
        param([string]$Message, [string]$Type = "error")
        $txtStatus.Text = $Message
        $statusBorder.Visibility = "Visible"
        switch ($Type) {
            "error"   { $statusBorder.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 255, 235, 238)); $txtStatus.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 198, 40, 40)) }
            "success" { $statusBorder.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 232, 245, 233)); $txtStatus.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 46, 125, 50)) }
            "info"    { $statusBorder.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 227, 242, 253)); $txtStatus.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 21, 101, 192)) }
        }
    }
    $script:HideStatus = {
        $statusBorder.Visibility = "Collapsed"
    }

    $btnStart.Add_Click({
        # Status verstecken
        $statusBorder.Visibility = "Collapsed"

        $selectedServers = @()
        if ($detectedServers.Count -gt 0) {
            foreach ($srv in $serverCheckboxes.Keys) {
                if ($serverCheckboxes[$srv].IsChecked) { $selectedServers += $srv }
            }
        }
        else {
            $selectedServers = @($txtServersManual.Text -split '[,;\s]+' | Where-Object { $_ })
        }

        if ($selectedServers.Count -eq 0) {
            & $script:ShowStatus -Message (Get-T "Fehler: Bitte mindestens einen Hyper-V Host auswählen.") -Type "error"
            return
        }
        if ([string]::IsNullOrWhiteSpace($txtOutputPath.Text)) {
            & $script:ShowStatus -Message (Get-T "Fehler: Bitte einen Ausgabepfad angeben.") -Type "error"
            return
        }

        $selectedSections = @()
        foreach ($key in $sectionCheckboxes.Keys) {
            if ($sectionCheckboxes[$key].IsChecked) { $selectedSections += $key }
        }
        if ($selectedSections.Count -eq 0) {
            & $script:ShowStatus -Message (Get-T "Fehler: Bitte mindestens einen Dokumentationsbereich auswählen.") -Type "error"
            return
        }

        $formats = @()
        if ($chkHtml.IsChecked)     { $formats += "HTML" }
        if ($chkPdf.IsChecked)      { $formats += "PDF" }
        if ($chkMarkdown.IsChecked) { $formats += "Markdown" }
        if ($formats.Count -eq 0) {
            & $script:ShowStatus -Message (Get-T "Fehler: Bitte mindestens ein Ausgabeformat wählen.") -Type "error"
            return
        }

        $script:GuiResult = @{
            Servers    = $selectedServers
            Company    = $txtCompany.Text
            OutputPath = $txtOutputPath.Text
            Sections   = $selectedSections
            Formats    = $formats
        }
        $window.Close()
    })

    $btnCancel.Add_Click({
        $script:GuiResult = $null
        $window.Close()
    })

    [void]$window.ShowDialog()
    return $script:GuiResult
}

#endregion

#region ============================================================
# HAUPTPROGRAMM
#endregion ============================================================

try {
    # ============================================================
    # 0. EINGABEN ERMITTELN (GUI ODER PARAMETER)
    # ============================================================

    # System.Web Assembly laden (für HtmlEncode in ConvertTo-HTMLTable)
    try { Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue } catch {}

    # Entscheiden ob GUI angezeigt wird:
    # - explizit via -ShowGui ODER
    # - keine Hosts angegeben und nicht -NoGui
    $useGui = $ShowGui -or (-not $HyperVServers -and -not $NoGui)

    # Standardauswahl der Sektionen (alle), falls keine angegeben
    $selectedSectionKeys = if ($Sections) { $Sections } else { (Get-DocSectionRegistry).Key }

    if ($useGui) {
        # Bei Sprachwechsel wird die GUI in der neuen Sprache neu aufgebaut
        do {
            $script:GuiRestart = $false
            $guiConfig = Show-DocumentationGui `
                -DefaultServers ($HyperVServers -join ", ") `
                -DefaultCompany $CompanyName `
                -DefaultOutputPath $OutputPath

            if ($script:GuiRestart) {
                if ($script:GuiPendingCompany) { $CompanyName = $script:GuiPendingCompany }
                if ($script:GuiPendingPath)    { $OutputPath  = $script:GuiPendingPath }
            }
        } while ($script:GuiRestart)

        if (-not $guiConfig) {
            Write-Host (Get-T "Abgebrochen durch Benutzer (GUI).") -ForegroundColor Yellow
            return
        }

        # GUI-Auswahl übernehmen
        $HyperVServers       = $guiConfig.Servers
        $CompanyName         = $guiConfig.Company
        $OutputPath          = $guiConfig.OutputPath
        $selectedSectionKeys = $guiConfig.Sections
        $OutputFormats       = $guiConfig.Formats
    }

    # Validierung
    if (-not $HyperVServers -or $HyperVServers.Count -eq 0) {
        throw (Get-T "Keine Hyper-V Hosts angegeben. Bitte -HyperVServers verwenden oder die GUI nutzen.")
    }

    # ============================================================
    # 0b. PFADE UND VARIABLEN (NEU) BERECHNEN
    # ============================================================
    $script:LogPath        = $OutputPath
    $stamp                 = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:LogFile        = Join-Path -Path $LogPath -ChildPath "HyperV-Dokumentation_$stamp.log"
    $script:HTMLOutputFile = Join-Path -Path $LogPath -ChildPath "HyperV_Dokumentation_$stamp.html"
    $script:PDFOutputFile  = Join-Path -Path $LogPath -ChildPath "HyperV_Dokumentation_$stamp.pdf"
    $script:MDOutputFile   = Join-Path -Path $LogPath -ChildPath "HyperV_Dokumentation_$stamp.md"
    $script:DocSubTitle    = $CompanyName

    # ============================================================
    # 1. VORBEREITUNG
    # ============================================================
    if (-not (Test-Path -Path $LogPath)) {
        New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
        Write-Host "$(Get-T 'Ausgabeverzeichnis erstellt'): $LogPath" -ForegroundColor Cyan
    }

    # Logging starten
    Write-Log -Message "=============================================" -Level "INFO"
    Write-Log -Message "$(Get-T 'Hyper-V Dokumentation gestartet') (v$script:ScriptVersion)" -Level "INFO"
    Write-Log -Message "$(Get-T 'Zielhosts'): $($HyperVServers -join ', ')" -Level "INFO"
    Write-Log -Message "$(Get-T 'Ausgabepfad:') $LogPath" -Level "INFO"
    Write-Log -Message "$(Get-T 'Ausgabeformate:') $($OutputFormats -join ', ')" -Level "INFO"
    Write-Log -Message "$(Get-T 'Gewählte Sektionen'): $($selectedSectionKeys.Count)" -Level "INFO"
    Write-Log -Message (Get-T "Verbindungsmodus: CIM mit automatischem DCOM Fallback") -Level "INFO"
    Write-Log -Message "=============================================" -Level "INFO"

    # ============================================================
    # 2. HYPER-V MODUL LADEN
    # ============================================================
    $hyperVLoaded = Initialize-HyperVEnvironment
    if (-not $hyperVLoaded) {
        Write-Log -Message (Get-T "Hyper-V PowerShell-Modul nicht verfügbar. Skript wird beendet.") -Level "ERROR"
        throw (Get-T "Hyper-V PowerShell-Modul nicht verfügbar. Bitte RSAT-Hyper-V-Tools installieren.")
    }

    # Host-Version und Cluster erkennen
    Get-HyperVEdition

    # DocTitle dynamisch anpassen
    $script:DocTitle = if ($script:ClusterName) {
        if ($script:Language -eq "EN") {
            "Hyper-V Cluster ($($script:ClusterName)) - Environment Documentation"
        }
        else {
            "Hyper-V Cluster ($($script:ClusterName)) - Umgebungsdokumentation"
        }
    }
    else {
        Get-T "Hyper-V - Umgebungsdokumentation"
    }

    # ============================================================
    # 3. KONNEKTIVITÄTSTEST
    # ============================================================
    Write-Log -Message (Get-T "=== Prüfe Erreichbarkeit der Hosts ===") -Level "INFO"
    foreach ($server in $HyperVServers) {
        try {
            if (-not (Test-Connection -ComputerName $server -Count 2 -Quiet -ErrorAction Stop)) {
                Write-Log -Message $(if ($script:Language -eq "EN") { "Host $server is NOT reachable!" } else { "Host $server ist NICHT erreichbar!" }) -Level "WARNING"
            }
            else {
                Write-Log -Message $(if ($script:Language -eq "EN") { "Host $server is reachable (ping OK)." } else { "Host $server ist erreichbar (Ping OK)." }) -Level "INFO"
            }
        }
        catch {
            Write-Log -Message $(if ($script:Language -eq "EN") { "Ping test for ${server} failed: $_" } else { "Ping-Test für ${server} fehlgeschlagen: $_" }) -Level "WARNING"
        }
    }

    # ============================================================
    # 4. DATENSAMMLUNG - NUR AUSGEWÄHLTE SEKTIONEN
    # ============================================================
    Write-Log -Message "$(Get-T '=== Starte Datensammlung') ($($selectedSectionKeys.Count) $(Get-T 'Sektionen')) ===" -Level "INFO"

    $registry = Get-DocSectionRegistry
    foreach ($section in $registry) {
        if ($selectedSectionKeys -contains $section.Key) {
            try {
                Write-Log -Message "$(Get-T 'Sektionen'): $(Get-T $section.Label) [$($section.Key)]" -Level "INFO"
                & $section.Function
            }
            catch {
                Write-Log -Message "$(Get-T 'Fehler') [$($section.Key)]: $_" -Level "ERROR"
            }
        }
    }

    # ============================================================
    # 5. DOKUMENTE ERZEUGEN UND SPEICHERN (PRO FORMAT)
    # ============================================================
    $createdFiles = [System.Collections.ArrayList]::new()

    # HTML wird immer für PDF benötigt - bei PDF ohne HTML temporär erzeugen
    $needHtml = ($OutputFormats -contains "HTML") -or ($OutputFormats -contains "PDF")

    if ($needHtml) {
        Write-Log -Message (Get-T "=== Generiere HTML-Dokument ===") -Level "INFO"
        $finalHTML = Build-HTMLDocument
        $finalHTML | Out-File -FilePath $script:HTMLOutputFile -Encoding UTF8 -Force
        if ($OutputFormats -contains "HTML") {
            [void]$createdFiles.Add($script:HTMLOutputFile)
        }
    }

    if ($OutputFormats -contains "PDF") {
        Write-Log -Message (Get-T "=== Generiere PDF-Dokument ===") -Level "INFO"
        if (Export-DocumentToPdf -HtmlPath $script:HTMLOutputFile -PdfPath $script:PDFOutputFile) {
            [void]$createdFiles.Add($script:PDFOutputFile)
        }
        # Temporäres HTML entfernen, falls HTML nicht gewünscht war
        if ($OutputFormats -notcontains "HTML" -and (Test-Path $script:HTMLOutputFile)) {
            Remove-Item -Path $script:HTMLOutputFile -Force -ErrorAction SilentlyContinue
        }
    }

    if ($OutputFormats -contains "Markdown") {
        Write-Log -Message (Get-T "=== Generiere Markdown-Dokument ===") -Level "INFO"
        $finalMD = Build-MarkdownDocument
        $finalMD | Out-File -FilePath $script:MDOutputFile -Encoding UTF8 -Force
        [void]$createdFiles.Add($script:MDOutputFile)
    }

    Write-Log -Message "=============================================" -Level "INFO"
    Write-Log -Message (Get-T "Dokumentation erfolgreich erstellt!") -Level "INFO"
    foreach ($f in $createdFiles) { Write-Log -Message "$(Get-T 'Datei'): $f" -Level "INFO" }
    Write-Log -Message "$(Get-T 'Log-Datei'):  $($script:LogFile)" -Level "INFO"
    Write-Log -Message "$(Get-T 'Fehler'): $($script:ErrorCount) | $(Get-T 'Warnungen'): $($script:WarningCount)" -Level "INFO"
    Write-Log -Message "=============================================" -Level "INFO"

    # Konsolenausgabe
    Write-Host "`n" -NoNewline
    Write-Host "===============================================================" -ForegroundColor Cyan
    Write-Host "  $(Get-T 'Hyper-V Dokumentation abgeschlossen!') (v$script:ScriptVersion)" -ForegroundColor Cyan
    Write-Host "===============================================================" -ForegroundColor Cyan
    foreach ($f in $createdFiles) {
        Write-Host "  $(Get-T 'Datei'): $f" -ForegroundColor White
    }
    Write-Host "  $(Get-T 'Log-Datei'):  $($script:LogFile)" -ForegroundColor White
    Write-Host "  $(Get-T 'Fehler'):     $($script:ErrorCount) | $(Get-T 'Warnungen'): $($script:WarningCount)" -ForegroundColor White
    Write-Host "===============================================================" -ForegroundColor Cyan

    # Datei optional öffnen
    if ($createdFiles.Count -gt 0) {
        $openFile = Read-Host "`n$(Get-T 'Dokumentation jetzt oeffnen? (J/N)')"
        if ($openFile -in @("J", "j", "Y", "y")) {
            Start-Process $createdFiles[0]
        }
    }
}
catch {
    Write-Log -Message "$(if ($script:Language -eq 'EN') { 'CRITICAL ERROR' } else { 'KRITISCHER FEHLER' }): $_" -Level "ERROR"
    Write-Log -Message "Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
    Write-Host "`n$(Get-T 'Kritischer Fehler! Details:') $($script:LogFile)" -ForegroundColor Red
}
finally {
    Write-Log -Message "$(Get-T 'Skript beendet um') $(Get-Date -Format 'HH:mm:ss')" -Level "INFO"
}

#endregion
