# Changelog

Alle nennenswerten Änderungen an diesem Projekt werden in dieser Datei dokumentiert.

Das Format orientiert sich an [Keep a Changelog](https://keepachangelog.com/de/1.1.0/),
die Versionierung folgt [Semantic Versioning](https://semver.org/lang/de/).

## [1.1.0] - 2026-08-31

### Hinzugefügt

- Neuer Parameter `-Language` mit den Werten `DE` und `EN` (Standard: `DE`).
  Steuert GUI, Konsolenausgaben, Logdatei und den erzeugten Report.
- Sprachumschalter (DE/EN) direkt in der grafischen Oberfläche. Beim Wechsel wird
  die Oberfläche unter Beibehaltung der Eingaben in der neuen Sprache neu aufgebaut.
- Englische Übersetzung für Sektionstitel, Kategorien, Zwischenüberschriften,
  Tabellenspalten, Statuswerte und Hinweistexte in HTML-, PDF- und Markdown-Ausgabe.
- Das HTML-Dokument setzt `lang="en"` bzw. `lang="de"` passend zur gewählten Sprache.

### Geändert

- Der Dokumenttitel wird sprachabhängig gesetzt
  ("Hyper-V - Environment Documentation" bzw. "Hyper-V - Umgebungsdokumentation").
- Die Abfrage zum Öffnen der Dokumentation akzeptiert jetzt `J` und `Y`.
- Beim automatischen Neustart mit erhöhten Rechten wird `-Language` weitergereicht.

## [1.0.0] - 2026-08-31

### Hinzugefügt

- Erstveröffentlichung von `HyperV_Documentation.ps1`.
- Automatische Datensammlung für Standalone-Hosts und Hyper-V Failover-Cluster.
- Verbindungsaufbau über CIM-Sessions mit automatischem DCOM-Fallback, wenn WinRM
  nicht verfügbar ist.
- 58 auswählbare Dokumentations-Sektionen in 9 Kategorien (Hardware & OS,
  Hyper-V Host, Virtuelle Maschinen, Virtuelle Netzwerke, Speicher,
  Failover Cluster, Replikation & Backup, Sicherheit, Active Directory).
- Grafische Oberfläche (WPF) zur Auswahl von Hosts, Firmenname, Ausgabepfad,
  Ausgabeformaten und Sektionen.
- Ausgabeformate HTML (Word-importierbar), PDF und Markdown.
- PDF-Export über Microsoft Word (COM) mit Fallback auf Microsoft Edge bzw.
  Google Chrome im Headless-Modus.
- Automatischer Neustart des Skripts mit erhöhten Rechten, falls es ohne
  Administratorrechte gestartet wird.
- Protokollierung aller Schritte in eine Logdatei inklusive Fehler- und
  Warnungszähler.
- Konnektivitätsprüfung (Ping) aller angegebenen Hosts vor der Datensammlung.

[1.1.0]: https://github.com/RoccoAmmon/Hyper-V-Documentation/releases/tag/v1.1.0
[1.0.0]: https://github.com/RoccoAmmon/Hyper-V-Documentation/releases/tag/v1.0.0
