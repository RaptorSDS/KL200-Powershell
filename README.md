# KL200-Powershell

PowerShell-Werkzeug zum Konfigurieren und Auslesen des Laser-Entfernungssensors
**XKC-KL200-2M-UART** (Messbereich 10–2000 mm, ±4 %) über eine serielle Schnittstelle.

- Interaktives Menü für alle Sensoreinstellungen
- Einzelmessung, Dauerüberwachung mit Statistik und CSV-Export
- Aufruf ohne Menü über Parameter (für Skripte und die Aufgabenplanung)
- Läuft unter **Windows PowerShell 5.1** und **PowerShell 7+**

## Inhalt

- [Dateien](#dateien)
- [Schnellstart](#schnellstart)
- [Anschluss](#anschluss)
- [Menü](#menü)
- [Aufruf mit Parametern](#aufruf-mit-parametern)
- [Sensoreinstellungen](#sensoreinstellungen)
- [Protokoll](#protokoll)
- [Fehlerbehebung](#fehlerbehebung)
- [Versionen](#versionen)
- [Lizenz](#lizenz)

## Dateien

| Datei | Beschreibung |
|---|---|
| `KL200_controller_version2.1.ps1` | **Aktuelle Version 2.1** |
| `KL200_controller_version2.ps1` | Version 2 – nur PowerShell 7, fehlerhafte Baudraten-Tabelle |
| `XKC_KL200_Controller.ps1` | Version 1 – Altversion, nicht mehr gepflegt |
| `Archiv/` | Datierte Sicherungen (Ordner und ZIP) mit Analyse |

## Schnellstart

1. Sensor über einen USB-UART-Adapter anschließen (siehe [Anschluss](#anschluss)).
2. PowerShell öffnen und in den Ordner mit dem Skript wechseln.
3. Falls Skripte blockiert sind, einmalig erlauben:
   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
   ```
4. Skript starten:
   ```powershell
   .\KL200_controller_version2.1.ps1
   ```
5. Im Menü `1` wählen, COM-Port auswählen, Baudrate mit Enter bestätigen (9600 = Werkseinstellung).
6. Mit `3` die Entfernung lesen.

Administratorrechte sind nicht nötig.

## Anschluss

| Ader | Signal | Anschluss am USB-UART-Adapter |
|---|---|---|
| Braun | VCC 5–24 V | Versorgung (5 V vom Adapter genügt) |
| Blau | GND | GND |
| Gelb | TX des Sensors | RX |
| Schwarz | RX des Sensors | TX |

- Laut Datenblatt gehört ein Pull-up von ca. **1 kΩ** zwischen die Logikspannung (3–5 V) und Gelb.
- Schnittstelle ab Werk: **9600 Baud, 8 Datenbits, keine Parität, 1 Stoppbit**.
- Der Sensor muss im **UART-Modus** sein (Ausgangsmodus = 1). Im Relais-Modus schaltet Gelb
  nur noch und es kommen keine Daten.

## Menü

| Nr. | Funktion |
|---|---|
| 1 | Verbinden – Port und Baudrate auswählen |
| 2 | Trennen |
| 3 | Entfernung lesen |
| 4 | Entfernung überwachen – Zeitstempel je Messung, Min/Max/Mittel, optional CSV; Abbruch mit `Q` |
| 5 | Hard Reset – Werkseinstellungen |
| 6 | Soft Reset – Benutzereinstellungen |
| 7 | Sensoradresse ändern (0–65534, dezimal oder `0x…`) |
| 8 | Baudrate ändern – das Skript verbindet danach automatisch neu |
| 9 | Upload-Modus: manuelle Abfrage oder automatisches Senden |
| 10 | Upload-Intervall (1–100 = 100 ms bis 10 s) |
| 11 | LED-Modus |
| 12 | Relais-Modus |
| 13 | Ausgangsmodus Relais / UART (mit Sicherheitsabfrage) |
| 14 | Konfiguration des Sensors anzeigen |
| 15 | Ansprechverhalten: hohe Genauigkeit / normal / hohe Stabilität |
| Q | Beenden |

Beim Verbinden liest das Skript die Konfiguration des Sensors. So weiß es, ob der Sensor
selbstständig sendet (automatischer Modus) oder abgefragt werden muss.

## Aufruf mit Parametern

```powershell
# Direkt verbinden und Menü öffnen
.\KL200_controller_version2.1.ps1 -Port COM3

# Einmal messen (Exitcode 0 = ok, 2 = keine Messung)
.\KL200_controller_version2.1.ps1 -Port COM3 -Action ReadDistance

# Konfiguration anzeigen
.\KL200_controller_version2.1.ps1 -Port COM3 -Action ReadConfig

# 60 Sekunden überwachen, alle 200 ms, Werte in CSV speichern
.\KL200_controller_version2.1.ps1 -Port COM3 -Action Monitor -Duration 60 -Interval 200 -CsvPath .\messung.csv

# Gesendete und empfangene Bytes anzeigen (Fehlersuche)
.\KL200_controller_version2.1.ps1 -Port COM3 -Action ReadDistance -Verbose
```

| Parameter | Standard | Beschreibung |
|---|---|---|
| `-Port` | – | COM-Port, z. B. `COM3`. Pflicht für alle Aktionen außer `Menu` |
| `-BaudRate` | 9600 | 2400, 4800, 9600, 14400, 19200, 38400, 56000, 57600, 115200, 128000 |
| `-Address` | 65535 | Zieladresse in Befehlen; 65535 (`0xFFFF`) erreicht jeden Sensor |
| `-Action` | `Menu` | `Menu`, `ReadDistance`, `ReadConfig`, `Monitor` |
| `-Duration` | 30 | Überwachungsdauer in Sekunden |
| `-Interval` | 500 | Abfrageintervall in ms (nur im manuellen Modus) |
| `-CsvPath` | – | CSV-Datei für die Überwachung (Spalten `Timestamp`, `DistanceMm`) |

Ausführliche Hilfe: `Get-Help .\KL200_controller_version2.1.ps1 -Full`

## Sensoreinstellungen

| Einstellung | Werkseinstellung | Werte |
|---|---|---|
| Baudrate | 2 = 9600 | 0 = 2400, 1 = 4800, 2 = 9600, 3 = 14400, 4 = 19200, 5 = 38400, 6 = 56000, 7 = 57600, 8 = 115200, 9 = 128000 |
| Adresse | 0 | 0–65534 (65535 = alle) |
| Erfassungsabstand | 100 mm | 10–2000 mm – nur per Taste auf der Rückseite einstellbar |
| Upload-Modus | – | 0 = manuelle Abfrage, 1 = automatisches Senden |
| Upload-Intervall | – | 1–100 (× 100 ms) |
| LED-Modus | 0 | 0 = an bei Erkennung, 1 = aus bei Erkennung, 2 = immer aus, 3 = immer an |
| Relais-Modus | 0 | 0 = aktiv bei Erkennung, 1 = inaktiv bei Erkennung |
| Ausgangsmodus | – | 0 = Relais, 1 = UART |
| Ansprechverhalten | – | 0 = hohe Genauigkeit, 1 = normal, 2 = hohe Stabilität |

Die Werkseinstellungen für Upload-Modus und Upload-Intervall sind im Datenblatt widersprüchlich
angegeben. Menüpunkt 14 zeigt die tatsächlichen Werte des angeschlossenen Sensors.

**Reset:** Der Hard Reset (`FE`) setzt die Werkseinstellungen, der Soft Reset (`FD`) die
gespeicherten Benutzereinstellungen. Nach dem Hard Reset verbindet das Skript bei Bedarf neu mit 9600 Baud.

## Protokoll

Jeder Befehl ist 9 Byte lang. Die Prüfsumme ist das XOR der Bytes 1–8.

```
Byte:  1     2     3     4      5      6      7      8        9
       62   Cmd    09   AdrH   AdrL   DatH   DatL   Antwort   XOR
```

| Befehl | Funktion | Daten (DatH/DatL) |
|---|---|---|
| `62 30` | Baudrate | Index 0–9 |
| `62 31` | Konfiguration lesen | – (Antwort: 29 Byte) |
| `62 32` | Adresse | neue Adresse |
| `62 33` | Entfernung lesen | – (Antwort: Entfernung in mm) |
| `62 34` | Upload-Modus | 0 / 1 |
| `62 35` | Upload-Intervall | 1–100 |
| `62 37` | LED-Modus | 0–3 |
| `62 38` | Relais-Modus | 0 / 1 |
| `62 39` | Reset | `FF FF`, Antwortbyte `FE` oder `FD` |
| `61 30` | Ausgangsmodus | 0 / 1 – Antwort kommt als `62 30` |
| `61 32` | Ansprechverhalten | 0–2 – Antwort kommt als `62 32` |

Bestätigung eines Befehls: `62 <Cmd> 09 00 00 00 00 66 <XOR>`

Beispiel Entfernung: Anfrage `62 33 09 FF FF 00 00 00 58`, Antwort `62 33 09 00 00 01 ED 00 B4` → `0x01ED` = 493 mm.

Aufbau der Konfigurationsantwort (`62 31`, Byte-Nummern ab 1):

| Byte | Inhalt |
|---|---|
| 12 | Baudrate (Index) |
| 13–14 | Adresse |
| 15–16 | Erfassungsabstand in mm |
| 17 | Upload-Modus |
| 18 | Upload-Intervall |
| 21 | LED-Modus |
| 22 | Relais-Modus |
| 23 | Ausgangsmodus |
| 28 | `66` = ok |
| 29 | XOR |

## Fehlerbehebung

| Problem | Ursache / Lösung |
|---|---|
| `No serial ports found` | Treiber des USB-UART-Adapters fehlt, Adapter im Geräte-Manager prüfen |
| `Error opening serial port` | Port von einem anderen Programm belegt (Arduino IDE, Terminal) – dort schließen |
| `Timeout waiting for distance data` | Falsche Baudrate, TX/RX vertauscht, Sensor im Relais-Modus oder ohne Versorgung |
| `Configuration not readable` | Wie oben; ältere Sensoren kennen eventuell den Befehl `62 31` nicht – Messen funktioniert trotzdem |
| `Checksum error` | Störungen auf der Leitung, Pull-up fehlt, Kabel zu lang |
| Baudrate unbekannt | Alle Baudraten nacheinander durchprobieren oder Hard Reset bei bekannter Baudrate |
| Skript startet nicht | `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` ausführen |
| Werte springen oder zu kurz | Starkes Licht, Staub oder zu kleines Objekt: Mindestdurchmesser ≈ Abstand × 0,47 |

Mit `-Verbose` zeigt das Skript alle gesendeten (`TX`) und empfangenen (`RX`) Bytes an.

## Versionen

### 2.1 (2026-10-14)

- Läuft unter PowerShell 5.1 und 7
- Baudraten-Tabelle laut Datenblatt korrigiert (in Version 2 war Index 3 fälschlich 9600 statt 14400)
- Falsche Meldungen behoben: kein „Unknown error“ mehr vor „successful“, keine zusätzlichen
  Ausgaben wie `True`, `False` oder `9`, „Exiting“ nur noch einmal
- Empfangspuffer wird vor jedem Befehl geleert, Datenpakete werden am Byte `0x62` neu erkannt
- Automatischer Modus wird beim Verbinden aus der Sensorkonfiguration erkannt
- Neu: Konfiguration anzeigen, Ansprechverhalten, Aufruf mit Parametern, CSV-Export, Statistik
- Aufgeräumter Code: ein gemeinsamer Befehls-Helfer, ein Frame-Leser, Eingabe-Helfer

### 2.0

- Port- und Baudratenauswahl, Antwortprüfung, Überwachung, automatischer Modus

### 1.0

- Erste Version mit einfachem Menü

## Lizenz

GPL-3.0 – siehe [`LICENSE`](LICENSE).
