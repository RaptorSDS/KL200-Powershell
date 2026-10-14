# KL200-Powershell

PowerShell-Werkzeug zum Konfigurieren und Auslesen des Laser-Entfernungssensors
**XKC-KL200-2M-UART** (Messbereich 10–2000 mm) über eine serielle Schnittstelle (USB-UART-Adapter).

## Dateien

| Datei | Beschreibung |
|---|---|
| `KL200_controller_version2.1.ps1` | **Aktuelle Version 2.1** – Windows PowerShell 5.1 und PowerShell 7+ |
| `KL200_controller_version2.ps1` | Version 2 (nur PowerShell 7, Baudraten-Tabelle fehlerhaft) |
| `XKC_KL200_Controller.ps1` | Version 1 (Altversion, nicht mehr gepflegt) |
| `Archiv/` | Datierte Sicherungen und Analysen |

## Voraussetzungen

- Windows PowerShell **5.1** oder PowerShell **7+**
- Sensor im **UART-Modus** (Line output mode = 1), angeschlossen über USB-UART-Adapter
- Verdrahtung: Braun = VCC (5–24 V), Blau = GND, Gelb = TX des Sensors → RX des Adapters,
  Schwarz = RX des Sensors ← TX des Adapters. Laut Datenblatt ca. 1 kΩ Pull-up an Gelb.
- Werkseinstellung der Schnittstelle: 9600 Baud, 8 Datenbits, keine Parität, 1 Stoppbit

Falls die Ausführung von Skripten blockiert ist:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

## Verwendung

### Interaktives Menü

```powershell
.\KL200_controller_version2.1.ps1
# oder direkt verbinden:
.\KL200_controller_version2.1.ps1 -Port COM3 -BaudRate 9600
```

| Nr. | Funktion |
|---|---|
| 1 | Verbinden (Port- und Baudratenauswahl) |
| 2 | Trennen |
| 3 | Entfernung lesen (manueller und automatischer Modus) |
| 4 | Entfernung überwachen (Zeitstempel, Min/Max/Mittel, optional CSV) – Abbruch mit `Q` |
| 5 | Hard Reset (Werkseinstellungen, Antwortbyte `FE`) |
| 6 | Soft Reset (Benutzereinstellungen, Antwortbyte `FD`) |
| 7 | Sensoradresse ändern (0–65534, dezimal oder `0x…`) |
| 8 | Baudrate ändern (verbindet danach automatisch neu) |
| 9 | Upload-Modus (0 = manuelle Abfrage, 1 = automatisch) |
| 10 | Upload-Intervall (1–100 = 100 ms – 10 s) |
| 11 | LED-Modus |
| 12 | Relais-Modus |
| 13 | Ausgangsmodus Relais / UART |
| 14 | Konfiguration des Sensors auslesen |
| 15 | Ansprechverhalten (hohe Genauigkeit / normal / hohe Stabilität) |
| Q | Beenden |

Beim Verbinden liest das Skript die Sensorkonfiguration und erkennt so, ob der Sensor im
automatischen Upload-Modus läuft.

### Ohne Menü (z. B. für Skripte / Aufgabenplanung)

```powershell
# Einmal messen – Ausgabe "Distance: 493 mm", Exitcode 2 bei Fehler
.\KL200_controller_version2.1.ps1 -Port COM3 -Action ReadDistance

# Konfiguration anzeigen
.\KL200_controller_version2.1.ps1 -Port COM3 -Action ReadConfig

# 60 Sekunden überwachen, alle 200 ms, in CSV speichern
.\KL200_controller_version2.1.ps1 -Port COM3 -Action Monitor -Duration 60 -Interval 200 -CsvPath .\messung.csv

# Rohdaten (gesendete/empfangene Bytes) anzeigen
.\KL200_controller_version2.1.ps1 -Port COM3 -Action ReadDistance -Verbose
```

| Parameter | Standard | Beschreibung |
|---|---|---|
| `-Port` | – | COM-Port, z. B. `COM3` |
| `-BaudRate` | 9600 | 2400, 4800, 9600, 14400, 19200, 38400, 56000, 57600, 115200, 128000 |
| `-Address` | 65535 | Zieladresse in Befehlen (65535 = `FFFF` = alle Sensoren) |
| `-Action` | Menu | `Menu`, `ReadDistance`, `ReadConfig`, `Monitor` |
| `-Duration` | 30 | Überwachungsdauer in Sekunden |
| `-Interval` | 500 | Abfrageintervall in ms (nur manueller Modus) |
| `-CsvPath` | – | CSV-Datei für die Überwachung |

## Protokoll (Kurzfassung aus dem Datenblatt)

Alle Befehle sind 9 Byte lang, die Prüfsumme ist XOR über die Bytes 1–8:

```
62 <Cmd> 09 <AdrH> <AdrL> <DatH> <DatL> <Antwort> <XOR>
```

| Cmd | Funktion | Werte |
|---|---|---|
| `62 30` | Baudrate | 0=2400, 1=4800, **2=9600**, 3=14400, 4=19200, 5=38400, 6=56000, 7=57600, 8=115200, 9=128000 |
| `62 31` | Konfiguration lesen | Antwort 29 Byte |
| `62 32` | Adresse | 0–65534 |
| `62 33` | Entfernung lesen | Antwort: DatH/DatL = mm |
| `62 34` | Upload-Modus | 0 = manuell, 1 = automatisch |
| `62 35` | Upload-Intervall | 1–100 (× 100 ms) |
| `62 37` | LED-Modus | 0 = an bei Erkennung, 1 = aus bei Erkennung, 2 = immer aus, 3 = immer an |
| `62 38` | Relais-Modus | 0 = aktiv bei Erkennung, 1 = inaktiv bei Erkennung |
| `62 39` | Reset | Daten `FFFF`, Antwortbyte `FE` = Werk, `FD` = Benutzer |
| `61 30` | Ausgangsmodus | 0 = Relais, 1 = UART (Antwort kommt mit `62 30`) |
| `61 32` | Ansprechverhalten | 0 = hohe Genauigkeit, 1 = normal, 2 = hohe Stabilität (Antwort mit `62 32`) |

Erfolgreiche Antwort: `62 <Cmd> 09 00 00 00 00 66 <XOR>`.

> Hinweis: Im Relais-Modus sendet der Sensor keine UART-Daten mehr. Der Erfassungsabstand
> (Kalibrierung) wird über die Taste auf der Rückseite eingestellt, nicht per Befehl.

## Änderungen in Version 2.1

- Läuft unter **PowerShell 5.1 und 7** (kein `? :`-Operator mehr)
- Korrekte **Baudraten-Tabelle** laut Datenblatt (in v2 war Index 3 fälschlich 9600 statt 14400)
- Keine falschen Meldungen mehr wie „Unknown error“ vor „successful“, keine überflüssigen
  `True`/`False`/`9`-Ausgaben, „Exiting“ erscheint nur noch einmal
- Empfangspuffer wird vor jedem Befehl geleert; Frames werden am Header `0x62` neu synchronisiert
- Automatischer Upload-Modus wird beim Verbinden aus der Sensorkonfiguration erkannt
- Neue Funktionen: Konfiguration lesen, Ansprechverhalten, Kommandozeilenparameter, CSV-Export
- Code zusammengefasst: ein gemeinsamer Befehls-Helfer, ein Frame-Parser, Eingabe-Helfer
  (die acht fast gleichen Befehlsfunktionen sind jetzt Einzeiler)

## Lizenz

GPL-3.0 – siehe `LICENSE`.
