# Analyse KL200-Powershell – Stand 2026-10-14

Archivierter Stand: Commit `b9f69fc` (funktionierende Version).
Dieses Dokument listet nur Vorschläge – **nichts davon ist eingebaut**.

## Bestand

| Datei | Zeilen | Beschreibung |
|---|---|---|
| `XKC_KL200_Controller.ps1` | 247 | Version 1: COM-Port beim Start, einfaches Menü, keine Antwortprüfung |
| `KL200_controller_version2.ps1` | 969 | Version 2: Port-Auswahl, Timeouts, Antwortprüfung (0x66), Monitoring, Auto-Upload |
| `redme.md` | 38 | Anleitung – beschreibt nur Version 1 |

## Version 2 (aktuelle Hauptversion)

### A. Fehler/Unschönheiten, die heute schon auftreten können

1. **Rückgabewert von `SerialPort.Read()` landet in der Ausgabe** (Z. 128, 472, 510).
   `Read()` liefert die Anzahl gelesener Bytes (9). In PowerShell wird alles, was nicht
   zugewiesen wird, Teil des Rückgabewerts. Folge in `Wait-ForResponse`: Ergebnis ist
   `@(9, 0)` statt `0`. Der `switch` läuft über beide Werte → es erscheint zuerst
   *„Unknown error during …“* und danach *„… successful“*. `Read-Distance` gibt `9` und
   die Distanz zurück, im Menü/Monitoring wird zusätzlich eine „9“ ausgegeben.
   Lösung: `[void]$script:serialPort.Read(...)` bzw. `$null = ...`.
2. **Weitere Rückgabewerte, die auf der Konsole erscheinen**: `Initialize-SerialPort`
   (True/False), alle `Set-*`/`Invoke-*Reset` (True/False), `Process-AutoData` im
   Monitoring. Im Menü mit `$null =` oder `| Out-Null` aufrufen.
3. **Quit wird doppelt verarbeitet**: `switch` ist case-insensitiv, `"q"` und `"Q"`
   treffen beide → „Exiting...“ erscheint zweimal. Ein Zweig reicht.
4. **Ternärer Operator `? :`** (Z. 277, 278, 384, 416) läuft nur unter **PowerShell 7+**.
   In Windows PowerShell 5.1 bricht schon das Parsen des ganzen Skripts ab.
   Entweder `#Requires -Version 7` oben einfügen oder `if/else` verwenden.
5. **Kein Leeren des Empfangspuffers vor dem Senden.** Reste (z. B. Auto-Upload-Frames)
   führen zu „Invalid response“. Vor jedem Befehl `DiscardInBuffer()` aufrufen.
6. **Keine Resynchronisation**: Beginnt der Puffer nicht mit `0x62`, werden 9 Bytes
   verworfen bzw. in `Process-AutoData` nur 1 zusätzliches Byte. Besser: bis zum
   Header `0x62` lesen, dann den Frame prüfen.
7. **Kommunikationsmodus widersprüchlich zu Version 1**: V2 sendet `0x61 0x30 …` und
   erwartet Antwort `0x30` (gleich wie Baudrate), Bedeutung 0=Relais/1=UART.
   V1 sendet `0x62 0x31 …`, Bedeutung 0=UART/1=Relais. Gegen das Datenblatt prüfen.
8. **`autoMode` wird nur lokal gemerkt.** Ist der Sensor beim Verbinden schon im
   Auto-Modus, glaubt das Skript „Manual“ und liest Auto-Frames als Antworten.

### B. Optimierungen (Struktur/Wartbarkeit)

1. **Ein gemeinsamer Befehls-Helfer** statt 8× identischem Muster
   (Frame bauen → Checksumme → senden → warten → `switch`):
   ```powershell
   function Invoke-SensorCommand([byte]$Header = 0x62, [byte]$Cmd, [byte]$Hi = 0, [byte]$Lo = 0, [byte]$Ext = 0, [string]$Label) { ... }
   ```
   Spart ca. 250 Zeilen, Fehlermeldungen nur noch an einer Stelle.
2. **Frame-Parsing zusammenlegen**: `Read-Distance` und `Process-AutoData` enthalten
   denselben Parser → `ConvertFrom-SensorFrame`.
3. **Baudraten-Tabelle nur einmal** (`$script:BaudRates`) – aktuell 3× definiert.
   Außerdem: beim Verbinden 1–10, beim Ändern 0–9 → einheitlich machen.
4. **`Test-Connected`-Helfer** statt 11× `if ($script:serialPort -and ...IsOpen)`.
5. **Eingabe-Helfer** `Read-Number -Prompt -Min -Max -Default` statt TryParse-Blöcke.
6. **`[ValidateRange()]`** an den Parametern statt manueller Bereichsprüfung.
7. **Stopwatch statt `Get-Date`-Differenzen** in den Warteschleifen.
8. **Konstanten** als `Set-Variable -Option Constant` bzw. `$script:` und Fehlercodes
   als `enum`.
9. **Konforme Verben** (`Get-Checksum`, `Read-AutoData`, `Watch-Distance`) –
   relevant, falls daraus ein Modul wird.
10. **Monitoring**: Countdown mit `\r` wird von „Distance: …“-Zeilen überschrieben;
    besser eine Zeile pro Messung mit Zeitstempel, optional CSV-Export.
11. **Kommandozeilenparameter** (`-Port COM3 -Baud 9600`) für Nicht-Interaktiv-Nutzung.
12. Kommentarkopf korrigieren: „ultrasonic“ → Laser-Sensor (KL200 ist ein Laser-ToF-Sensor).

## Version 1 (`XKC_KL200_Controller.ps1`)

- **Menü-Parameter ohne Klammern** (Z. 185–213): `-HardReset [bool]::Parse($x)` wird
  als Text „[bool]::Parse“ übergeben → Fehler. Richtig: `-HardReset ([bool]::Parse($x))`.
  Menüpunkte 1–8 funktionieren daher vermutlich nicht.
- **Unterschiedliche Frame-Längen** (6–8 Bytes + Checksumme), V2 sendet immer 9 Bytes.
- Keine `ReadTimeout` → `Read-Distance` kann hängen; `Read` kann weniger als 9 Bytes liefern.
- Empfehlung: V1 als Altversion stehen lassen, nur V2 weiterentwickeln.

## Repository

- `redme.md` → Tippfehler (`README.md`), beschreibt nur V1; V2 + PowerShell-7-Voraussetzung ergänzen.
- Keine Tests – die Checksummen- und Frame-Logik ließe sich mit Pester ohne Hardware testen.

## Vorgeschlagene Reihenfolge

1. A1–A3 (Ausgabe-Fehler, minimal und risikoarm)
2. A4 (`#Requires -Version 7`), A5 (`DiscardInBuffer`)
3. A7 mit Datenblatt klären
4. B1–B5 Refactoring
