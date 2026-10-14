Anleitung zur Nutzung des PowerShell-Skripts

    PowerShell-Skript speichern:
    Speichern Sie das obige Skript in einer Datei namens XKC_KL200_Controller.ps1.

    PowerShell-Konsole öffnen:
    Öffnen Sie eine PowerShell-Konsole mit Administratorrechten.

    Skript ausführen:
    Navigieren Sie zum Speicherort der XKC_KL200_Controller.ps1-Datei und führen Sie das Skript aus:

    powershell

    .\XKC_KL200_Controller.ps1

    COM-Port und Baudrate eingeben:
    Geben Sie den COM-Port und die Baudrate Ihres XKC-KL200 Lasersensors ein, wenn Sie dazu aufgefordert werden.

    Menü nutzen:
    Verwenden Sie das Menü, um verschiedene Einstellungen am Sensor vorzunehmen und den Abstand zu lesen.

Funktionen des Skripts

    Send-Command: Sendet einen Befehl an den Sensor.
    Calculate-Checksum: Berechnet die Prüfsumme eines Befehls.
    Initialize-Sensor: Initialisiert den Sensor mit der angegebenen Baudrate.
    Restore-Factory-Settings: Setzt den Sensor auf die Werkseinstellungen zurück (Hard- oder Soft-Reset).
    Change-Address: Ändert die Adresse des Sensors.
    Change-Baud-Rate: Ändert die Baudrate des Sensors.
    Set-Upload-Mode: Setzt den Upload-Modus (manuell oder automatisch).
    Set-Upload-Interval: Setzt das Upload-Intervall.
    Set-LED-Mode: Setzt den LED-Modus.
    Set-Relay-Mode: Setzt den Relaismodus.
    Set-Communication-Mode: Setzt den Kommunikationsmodus (UART oder Relais).
    Read-Distance: Liest die Entfernung vom Sensor.
    Show-Menu: Zeigt das Hauptmenü an und ermöglicht die Auswahl der verschiedenen Funktionen.

Das Skript ermöglicht eine einfache und interaktive Steuerung des XKC-KL200 Lasersensors über die PowerShell-Konsole.
