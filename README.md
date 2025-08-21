# PowerAnalyzer (IP-Symcon Modul)

Ein Modul zur Auswertung von Leistungs-Variablen (Watt) inkl. Heatmap/Dezile usw.

## Installation
Siehe unten in dieser README oder im Repo-Root: über *Modulsteuerung* eine neue URL hinzufügen
oder per `git clone` nach `/var/lib/symcon/modules/`.

## Konfiguration
- **PowerVariableID** (integer): ID der Watt-Variable
- **ArchiveControlID** (integer): Standard 12496 (anpassbar)
- **MonthsBack** (integer): Anzahl Monate für Heatmap
- **UpdateSeconds** (integer): Ausführungsintervall (Default 300s)
- **EnableDebug** (boolean): Debug-Ausgaben im Meldungsfenster

## Nutzung
Nach dem Anlegen der Instanz `PowerAnalyzer` die IDs setzen, übernehmen, fertig.