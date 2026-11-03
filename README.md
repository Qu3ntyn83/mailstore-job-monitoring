# MailStore Job Monitoring

[![Buy Me a Coffee](https://img.shields.io/badge/Buy%20Me%20a%20Coffee-unterst%C3%BCtzen-FFDD00?logo=buymeacoffee&logoColor=black)](https://buymeacoffee.com/quentyn83)

PowerShell-Paket, das MailStore-Server-Archivierungsjobs überwacht, bei hängenden Jobs alarmiert und stufenweise eingreift. Alarme laufen über den Windows-Ereignisprotokoll (Event Log), sodass jedes Monitoring-Tool (RMM, Zabbix, PRTG, …) sie auswerten kann. Kein SMTP, keine Cloud.

> **Status: in Arbeit.** Die Überwachung (Monitor) ist im Einsatz getestet. Die Eskalation ist es nur teilweise, siehe [Bekannte Einschränkungen](#bekannte-einschränkungen) und [Warnung zu taskkill](#warnung-taskkill-und-safe-mode).

## Warum

Bei MailStore Server 26 bleibt ein Archivierungsprofil (z. B. Microsoft 365 per Graph) gelegentlich über Stunden oder Tage im Zustand „Wird ausgeführt“ hängen. Die Administration API liefert über `GetWorkerResults` nur abgeschlossene Jobs, laufende Jobs tauchen dort nicht auf.

Das Paket erkennt einen hängenden Job deshalb indirekt: Liegt der letzte Abschluss länger zurück als *maximale Laufzeit + Pause zwischen den Läufen*, läuft vermutlich ein Job zu lange.

## Dateien

| Datei | Aufgabe |
| --- | --- |
| `Install.ps1` | Geführter Installations-Wizard (einmalig, als Administrator) |
| `Monitor-MailStoreJobs.ps1` | Prüft alle 15 Minuten, schreibt Event 9000–9003 |
| `Escalate-MailStoreJobs.ps1` | Greift ab 180 Minuten Laufzeit ein, schreibt Event 9010–9016 |
| `Update-Config.ps1` | Ändert Passwort, Server und Schwellenwerte nachträglich |

Beim Installieren entstehen im Zielordner zusätzlich `mailstore-config.xml` (Einstellungen), `mailstore-key.bin` (AES-Schlüssel) und die State-Dateien `mailstore_alert_state.json` und `mailstore_escalation_state.json`. Diese Dateien gehören nicht ins Repository, die `.gitignore` schließt sie aus.

## Voraussetzungen

- MailStore Server 26 (getestet mit 26.2.1)
- Windows Server 2019 oder neuer, PowerShell 5.1, `curl.exe` (in Windows enthalten)
- Administration API aktiviert: Dienst-Konfiguration → Netzwerkeinstellungen → „MailStore Administration API (HTTPS)“ (im Wizard ist Port 8463 voreingestellt, den Port der eigenen Installation dort ablesen)
- Ausführung lokal auf dem MailStore-Server als Administrator

> **Hinweis zu TLS:** Die Skripte rufen die API mit `curl.exe -k` auf. Im getesteten Setup akzeptierte der Server nur TLS 1.0 und ein selbstsigniertes Zertifikat, was PowerShell/.NET 5.1 ablehnt. `-k` schaltet die Zertifikatsprüfung ab. Das ist bei einem Aufruf auf `localhost` vertretbar, für einen entfernten Server solltest du ein gültiges Zertifikat einrichten und `-k` entfernen.

## Installation

```powershell
# PowerShell als Administrator
cd C:\Temp\mailstore-job-monitoring
.\Install.ps1
```

Der Wizard fragt Installationspfad, Server, Port, API-Benutzer und Passwort ab, testet die Verbindung, legt die Schwellenwerte fest, erkennt den MailStore-Dienst, erzeugt Schlüssel und Konfiguration, richtet die Event-Log-Quelle `MailStore Monitor` ein, registriert zwei Scheduled Tasks (alle 15 Minuten, als SYSTEM) und führt einen Testlauf aus.

Standardwerte:

| Einstellung | Standard |
| --- | --- |
| Alarm nach (Laufzeit) | 90 Min. |
| Pause zwischen Jobs | 5 Min. (`automaticPauseBetweenExecutions` ÷ 60) |
| Eskalation nach | 180 Min. |
| Wiederholungsalarm | 60 Min. |
| Suchfenster für Jobs | 12 Std. |

Die tatsächlichen Schwellen sind *Wert + Pause*, also 95 Minuten für den Alarm und 185 Minuten für die Eskalation.

## Event IDs

Quelle `MailStore Monitor`, Protokoll `Application`.

**Monitor**

| ID | Typ | Bedeutung |
| --- | --- | --- |
| 9000 | Information | Alles in Ordnung |
| 9001 | Warning | **Alarm:** Job läuft zu lange (darauf im Monitoring-Tool alarmieren) |
| 9002 | Error | MailStore-API nicht erreichbar |
| 9003 | Information | Wiederholungsalarm unterdrückt (Throttling) |

**Eskalation**

| ID | Typ | Bedeutung |
| --- | --- | --- |
| 9010 | Information | Eskalation gestartet |
| 9011 | Information | Job per API abgebrochen |
| 9012 | Warning | MailStore-Dienst gestoppt |
| 9013 | Warning | Prozess per `taskkill` beendet |
| 9014 | Information | Eskalation abgeschlossen, Profil wieder auf automatisch |
| 9015 | Error | Eskalation fehlgeschlagen, manueller Eingriff nötig |
| 9016 | Warning | Profil auf manuell gesetzt, Abbruch läuft |

Im Monitoring-Tool einen Event-Log-Check auf Protokoll `Application`, Quelle `MailStore Monitor`, Event ID `9001` (und `9015`) einrichten.

## Eskalationsablauf

```
Stufe 1  Profil auf manuell setzen
Stufe 2  CancelJobAsync (API), 30 Sek. warten, Abschluss prüfen
           ok  -> Profil auf automatisch, fertig
Stufe 3  Stop-Service (max. 180 Sek.)
           ok  -> Start-Service, Profil auf automatisch, fertig
         Timeout -> taskkill /F /T, 30 Sek. warten
           Dienst gestoppt  -> Start-Service, Profil auf automatisch
           Dienst läuft     -> Dienst neu starten, Profil auf automatisch
           sonst            -> Event 9015
```

## Sicherheit der Zugangsdaten

Das API-Passwort steht nicht im Klartext in den Skripten. `Install.ps1` erzeugt einen zufälligen 256-Bit-AES-Schlüssel (`mailstore-key.bin`, NTFS-Rechte nur SYSTEM und Administratoren) und speichert das Passwort damit verschlüsselt in `mailstore-config.xml`.

Das schützt vor beiläufigem Mitlesen der Konfiguration. Es schützt **nicht** vor einem Administrator auf demselben Server, denn Schlüssel und Konfiguration liegen nebeneinander. Verwende am besten ein eigenes API-Konto mit möglichst wenigen Rechten und rotiere das Passwort mit `Update-Config.ps1`.

## Warnung: taskkill und Safe Mode

Bei einem Test beim Kunden führte `taskkill /F` auf den MailStore-Prozess dazu, dass MailStore Server im Safe Mode startete und sich kein Benutzer mehr verbinden konnte. Der Abbruch wird offenbar als Absturz gewertet. `taskkill` ist deshalb nur noch die letzte Stufe, wenn der Dienst auf `Stop-Service` nicht reagiert. Ob der Safe-Mode-Effekt mit dem aktuellen Ablauf (Profil vorher auf manuell, Neustart über den Dienst) ausbleibt, ist **nicht verifiziert**. Wer dieses Risiko nicht eingehen will, entfernt den Abschnitt „STUFE 3 Fallback: taskkill“ in `Escalate-MailStoreJobs.ps1`; dann endet die Eskalation bei Event 9015 und ein Techniker greift ein.

## Bekannte Einschränkungen

- **Profil auf manuell/automatisch setzen ist ungetestet.** Die API-Methode dafür ist nicht eindeutig dokumentiert. Das Skript probiert nacheinander `SetProfileServerSideExecution`, `SetProfileSchedule` und `SetServerSideExecution` und arbeitet ohne diesen Schritt weiter, wenn keine passt. Danach muss das Profil im MailStore-Client geprüft werden.
- **`CancelJobAsync`** (Parameter `id` = Profil-ID) hat im Test einen hängenden Job beendet. Das ist der bestätigte Teil der Eskalation.
- **Laufende Jobs sind nur indirekt erkennbar.** Dauert ein Hänger länger als das Suchfenster (Standard 12 Std.) ohne dass vorher ein Alarm ausgelöst wurde, kennt das Skript die Profil-ID nicht mehr und die Eskalation bricht mit Hinweis ab.
- Mehrere Profile werden nicht einzeln, sondern über den jeweils letzten Abschluss insgesamt betrachtet.
- Nicht auf anderen MailStore-Versionen als 26.2.1 getestet.

## Konfiguration ändern

```powershell
.\Update-Config.ps1   # Passwort, Server/Port oder Schwellenwerte
```

## Deinstallation

```powershell
Unregister-ScheduledTask -TaskName 'MailStore Job Monitor' -Confirm:$false
Unregister-ScheduledTask -TaskName 'MailStore Job Eskalation' -Confirm:$false
Remove-EventLog -Source 'MailStore Monitor'
Remove-Item -Recurse -Force <Installationspfad>
```

## Unterstützen

Wenn dir das Projekt hilft und du mir einen Kaffee ausgeben möchtest, freue ich mich sehr:

<a href="https://buymeacoffee.com/quentyn83"><img src="https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png" alt="Buy Me a Coffee" height="45"></a>

## Lizenz

[MIT](LICENSE). Nutzung auf eigene Verantwortung, besonders die Eskalationsstufen greifen in einen produktiven Archivierungsdienst ein.
