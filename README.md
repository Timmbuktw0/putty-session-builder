# PuTTY Session Builder

Kleines Programm (erstellt mit Codex), um PuTTY-Sessions im Bulk aus CSV-, XLS- oder XLSX-Dateien zu importieren.

## Download und Start

**[Fertige Windows-ZIP herunterladen](https://github.com/Timmbuktw0/putty-session-builder/raw/refs/heads/main/downloads/PuTTY-Session-Builder.zip)**

1. ZIP vollständig entpacken.
2. `PuTTYSessionBuilder.exe` starten. Die mitgelieferte DLL im selben Ordner belassen.
3. Eine Kopie der Vorlage im Ordner `Vorlagen` befüllen.
4. **Datei laden → Prüfen → Sessions importieren**.
5. PuTTY neu öffnen.

**Zeile 1 enthält die Überschriften. Zeile 2 enthält Beispieldaten und wird immer übersprungen. Eigene Sessions ab Zeile 3 eintragen.**

Voraussetzung: Windows 10/11 mit .NET Framework 4.8. Excel, Python und Administratorrechte sind nicht erforderlich. PuTTY ist separat erforderlich.

## Funktionen

- Import von CSV, XLS und XLSX.
- Editierbare Vorschau mit Dropdowns für installierte Monospace-Schriften und weitere Auswahlfelder.
- Excel-Vorlage mit Dropdowns bis Zeile 1000 sowie CSV-Vorlage.
- Validierung von Pflichtwerten, Ports, Schriftarten, Booleschen Werten und doppelten Namen.
- Direkter Registry-Import nach `HKEY_CURRENT_USER\Software\SimonTatham\PuTTY\Sessions`.
- Bestehende Sessions standardmäßig überspringen oder auf Wunsch aktualisieren.
- Automatische REG-Sicherung der betroffenen Sessions vor dem direkten Import.
- REG-Export zur Übernahme auf einem anderen PC.

## Tabellenfelder

| Feld | Standard bei leerem Feld |
| --- | --- |
| Anzeigename | Pflichtfeld |
| Hostname/IP | Pflichtfeld |
| Port | 22, bei telnet 23 |
| Font | Consolas |
| Fontvariante | Normal; alternativ Fett |
| Fontgröße | 11 pt |
| Fontquality | Standard; alternativ Antialiasing, Keine Glättung, ClearType |
| Keepalive | TRUE |
| Keepalive Rate | **5 Sekunden** |
| Autologin Username | leer |
| X11 Forwarding | FALSE |
| Protokoll | ssh; alternativ telnet oder raw |
| TCP Keepalive | FALSE |
| X11 Display | leer |
| SSH Keyfile | leer |
| Scrollback | 2000 |

Boolean-Felder akzeptieren TRUE/FALSE, WAHR/FALSCH, JA/NEIN und 1/0. Keepalive bezeichnet PuTTY-Anwendungs-Keepalives; TCP Keepalive ist eine separate Einstellung. X11 Forwarding benötigt SSH und einen lokalen X-Server.

Neue Sessions übernehmen vorhandene PuTTY-Default-Settings, anschließend werden die Tabellenwerte gesetzt. Beim Aktualisieren bestehender Sessions bleiben andere Einstellungen erhalten.

Die Sicherungen liegen unter `%LOCALAPPDATA%\PuTTYSessionBuilder\Backups`. Der direkte Import liest die geschriebenen Werte zur Kontrolle zurück. Die ausführliche Anleitung befindet sich in [Anleitung.html](Anleitung.html).

## Quellcode und Tests

- [Source/Program.cs](Source/Program.cs): Windows-Forms-Oberfläche, Import, Prüfung, Registry und Sicherung.
- [Source/Build.cmd](Source/Build.cmd): Build mit dem .NET-Framework-Compiler unter Windows.
- [Source/Test.cmd](Source/Test.cmd): Tests unter Windows; Registry-Tests verwenden einen separaten Testzweig und entfernen ihn danach.
- [Source/Tests.cs](Source/Tests.cs): enthaltene Tests.

Der ausgelieferte Stand wurde kompiliert und mit **64 Prüfungen** unter Mono in einer Linux-Testumgebung geprüft, einschließlich eines zusätzlichen binären XLS-Testfalls. Der zusätzliche XLS-Testfall stammt aus dem Upstream-Testbestand und ist nicht Teil dieses Repositorys. Der enthaltene Testlauf ohne diesen Testfall umfasst 63 Prüfungen. Ein echter Windows-/PuTTY-Lauf wurde hier nicht durchgeführt. Die EXE ist nicht digital signiert.

## Lizenz

MIT, siehe [LICENSE.txt](LICENSE.txt). Die mitgelieferte ExcelDataReader-Bibliothek 3.7.0 steht ebenfalls unter MIT; siehe [LICENSE-ExcelDataReader.txt](LICENSE-ExcelDataReader.txt).
