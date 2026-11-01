#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installations-Wizard fuer MailStore Job Monitor.
    Richtet Konfiguration, Verschluesselung, Event Log Source und Scheduled Tasks ein.

.NOTES
    Ausfuehren als Administrator auf dem MailStore-Server.
    Erstellt:
      <Installpfad>\mailstore-config.xml   - Konfigurationsdatei
      <Installpfad>\mailstore-key.bin      - AES-Schluesseldatei (NTFS-geschuetzt)
      <Installpfad>\mailstore_alert_state.json
      Scheduled Task "MailStore Job Monitor"
      Scheduled Task "MailStore Job Eskalation"
#>

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# ── Hilfsfunktionen ───────────────────────────────────────────────────────────

function Write-Header {
    param([string]$Text)
    $Line = "=" * 60
    Write-Host ""
    Write-Host $Line -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host $Line -ForegroundColor Cyan
    Write-Host ""
}

function Read-InputWithDefault {
    param([string]$Prompt, [string]$Default)
    $Display = if ($Default) { "$Prompt [$Default]" } else { $Prompt }
    $Value = Read-Host $Display
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Default }
    return $Value.Trim()
}

function Test-MailStoreConnection {
    param([string]$Server, [int]$Port, [string]$User, [string]$Password)
    $CredB64 = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("${User}:${Password}"))
    $Url = "https://${Server}:${Port}/api/invoke/GetServerInfo"
    try {
        $Raw = curl.exe -k -s --max-time 10 -X POST -H "Authorization: Basic $CredB64" -H "Content-Length: 0" $Url 2>&1
        if (-not $Raw) { return $false }
        $Json = ($Raw | Out-String).TrimStart([char]0xFEFF).Trim() | ConvertFrom-Json
        return ($Json.statusCode -eq "succeeded")
    }
    catch { return $false }
}

function New-AesKey {
    $Key = New-Object byte[] 32
    [System.Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($Key)
    return $Key
}



# ── Willkommen ────────────────────────────────────────────────────────────────

Clear-Host
Write-Header "MailStore Job Monitor - Installations-Wizard"
Write-Host "Dieser Wizard richtet die automatische Ueberwachung des MailStore" -ForegroundColor White
Write-Host "Archivierungsdienstes ein und konfiguriert die Eskalation." -ForegroundColor White
Write-Host ""
Write-Host "Voraussetzungen:" -ForegroundColor Yellow
Write-Host "  - Ausfuehrung als Administrator auf dem MailStore-Server" -ForegroundColor Yellow
Write-Host "  - curl.exe verfuegbar (Windows 10/Server 2019+)" -ForegroundColor Yellow
Write-Host "  - MailStore Administration API aktiviert (Port 8463)" -ForegroundColor Yellow
Write-Host ""
$Continue = Read-Host "Weiter? (J/N)"
if ($Continue -notmatch "^[Jj]") { Write-Host "Abgebrochen."; exit 0 }

# ── Installationspfad ─────────────────────────────────────────────────────────

Write-Header "Schritt 1/5: Installationspfad"
$InstallPath = Read-InputWithDefault "Installationspfad" "C:\Scripts\MailStore-Monitor"

if (-not (Test-Path $InstallPath)) {
    New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
    Write-Host "  Verzeichnis erstellt: $InstallPath" -ForegroundColor Green
} else {
    Write-Host "  Verzeichnis vorhanden: $InstallPath" -ForegroundColor Green
}

# Pruefen ob Scripts vorhanden
$MonitorScript   = Join-Path $InstallPath "Monitor-MailStoreJobs.ps1"
$EscalateScript  = Join-Path $InstallPath "Escalate-MailStoreJobs.ps1"
$ScriptSource    = Split-Path $MyInvocation.MyCommand.Path

foreach ($Script in @("Monitor-MailStoreJobs.ps1", "Escalate-MailStoreJobs.ps1")) {
    $Src = Join-Path $ScriptSource $Script
    $Dst = Join-Path $InstallPath $Script
    if (Test-Path $Src) {
        Copy-Item $Src $Dst -Force
        Write-Host "  Kopiert: $Script" -ForegroundColor Green
    } elseif (-not (Test-Path $Dst)) {
        Write-Host "  WARNUNG: $Script nicht gefunden in $ScriptSource" -ForegroundColor Yellow
        Write-Host "  Bitte manuell in $InstallPath ablegen." -ForegroundColor Yellow
    }
}

# ── MailStore Verbindung ──────────────────────────────────────────────────────

Write-Header "Schritt 2/5: MailStore Verbindung"

$Server   = Read-InputWithDefault "MailStore Servername oder IP" "localhost"
$Port     = [int](Read-InputWithDefault "API-Port" "8463")
$ApiUser  = Read-InputWithDefault "API-Benutzername" "admin"
$ApiPass  = Read-Host "API-Passwort" -AsSecureString
$ApiPassPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ApiPass)
)

Write-Host ""
Write-Host "  Teste Verbindung zu ${Server}:${Port} ..." -NoNewline
if (Test-MailStoreConnection -Server $Server -Port $Port -User $ApiUser -Password $ApiPassPlain) {
    Write-Host " OK" -ForegroundColor Green
} else {
    Write-Host " FEHLGESCHLAGEN" -ForegroundColor Red
    Write-Host "  Verbindung konnte nicht hergestellt werden." -ForegroundColor Red
    Write-Host "  Moegliche Ursachen: falsches Passwort, Port gesperrt, API nicht aktiv." -ForegroundColor Yellow
    $Force = Read-Host "  Trotzdem fortfahren? (J/N)"
    if ($Force -notmatch "^[Jj]") { exit 1 }
}

# ── Ueberwachungsparameter ────────────────────────────────────────────────────

Write-Header "Schritt 3/5: Ueberwachungsparameter"

Write-Host "  Alarm wenn kein Job-Abschluss seit X Minuten (Laufzeit-Schwelle):"
$MaxRuntime   = [int](Read-InputWithDefault "  Alarm nach Minuten" "90")
Write-Host ""
Write-Host "  Pause zwischen Job-Laeufen laut MailStore-Profil (automaticPauseBetweenExecutions):"
$JobPause     = [int](Read-InputWithDefault "  Pause in Minuten" "5")
Write-Host ""
Write-Host "  Eskalation (API-Stop / Dienst-Neustart) nach:"
$EscalateAfter = [int](Read-InputWithDefault "  Eskalation nach Minuten" "180")
Write-Host ""
Write-Host "  Wiederholungsalarm-Intervall (Throttling):"
$ThrottleMin  = [int](Read-InputWithDefault "  Erneuter Alarm nach Minuten" "60")
Write-Host ""
$LookbackH    = [int](Read-InputWithDefault "  Suchfenster fuer Jobs in Stunden" "12")

# ── MailStore Dienst ──────────────────────────────────────────────────────────

Write-Header "Schritt 4/5: Windows-Dienst fuer Eskalation"

$Services = Get-Service | Where-Object { $_.DisplayName -like "*MailStore*" } |
    Select-Object Name, DisplayName, Status
if ($Services) {
    Write-Host "  Gefundene MailStore-Dienste:" -ForegroundColor Green
    $Services | ForEach-Object { Write-Host "    $($_.Name) - $($_.DisplayName) [$($_.Status)]" }
    $ServiceName = Read-InputWithDefault "  Dienstname fuer Eskalation" $Services[0].Name
} else {
    Write-Host "  Kein MailStore-Dienst gefunden." -ForegroundColor Yellow
    $ServiceName = Read-InputWithDefault "  Dienstname manuell eingeben" "MailStore Server"
}

$ProcessName = Read-InputWithDefault "  Prozessname (ohne .exe) fuer taskkill" "MailStoreServer"

# ── Konfiguration speichern ───────────────────────────────────────────────────

Write-Header "Schritt 5/5: Konfiguration speichern und einrichten"

# AES-Schluessel generieren und speichern
$KeyFile = Join-Path $InstallPath "mailstore-key.bin"
$AesKey  = New-AesKey
[System.IO.File]::WriteAllBytes($KeyFile, $AesKey)

# NTFS-Berechtigungen: nur SYSTEM und Administratoren
$Acl = Get-Acl $KeyFile
$Acl.SetAccessRuleProtection($true, $false)
$Acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
    "SYSTEM", "FullControl", "Allow")))
$Acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
    "Administrators", "FullControl", "Allow")))
Set-Acl -Path $KeyFile -AclObject $Acl
Write-Host "  AES-Schluesseldatei erstellt (NTFS-geschuetzt): $KeyFile" -ForegroundColor Green

# Passwort verschluesseln
$SecStr = ConvertTo-SecureString $ApiPassPlain -AsPlainText -Force
$EncryptedPassword = $SecStr | ConvertFrom-SecureString -Key $AesKey

# Config-XML schreiben
$ConfigFile = Join-Path $InstallPath "mailstore-config.xml"
$Config = [ordered]@{
    MailStoreServer        = $Server
    ApiPort                = $Port
    ApiUser                = $ApiUser
    ApiPasswordEncrypted   = $EncryptedPassword
    MaxRuntimeMinutes      = $MaxRuntime
    PauseBetweenJobs       = $JobPause
    LookbackHours          = $LookbackH
    EscalateAfterMinutes   = $EscalateAfter
    AlertIntervalMinutes   = $ThrottleMin
    ServiceName            = $ServiceName
    ProcessName            = $ProcessName
    StopServiceTimeoutSec  = 60
    StartServiceTimeoutSec = 90
    EventLogName           = "Application"
    EventLogSource         = "MailStore Monitor"
}

$XmlDoc = New-Object System.Xml.XmlDocument
$Root = $XmlDoc.CreateElement("MailStoreMonitor")
$XmlDoc.AppendChild($Root) | Out-Null
foreach ($Key in $Config.Keys) {
    $El = $XmlDoc.CreateElement($Key)
    $El.InnerText = [string]$Config[$Key]
    $Root.AppendChild($El) | Out-Null
}
$XmlDoc.Save($ConfigFile)
Write-Host "  Konfigurationsdatei gespeichert: $ConfigFile" -ForegroundColor Green

# Event Log Source
if (-not [System.Diagnostics.EventLog]::SourceExists("MailStore Monitor")) {
    New-EventLog -LogName "Application" -Source "MailStore Monitor"
    Write-Host "  Event Log Source 'MailStore Monitor' angelegt." -ForegroundColor Green
} else {
    Write-Host "  Event Log Source bereits vorhanden." -ForegroundColor Green
}

# Scheduled Tasks
$TaskArgs = "-NonInteractive -ExecutionPolicy Bypass -File"

foreach ($Task in @(
    @{ Name = "MailStore Job Monitor";    Script = $MonitorScript },
    @{ Name = "MailStore Job Eskalation"; Script = $EscalateScript }
)) {
    $Existing = Get-ScheduledTask -TaskName $Task.Name -ErrorAction SilentlyContinue
    if ($Existing) {
        Unregister-ScheduledTask -TaskName $Task.Name -Confirm:$false
        Write-Host "  Alter Task '$($Task.Name)' entfernt." -ForegroundColor Yellow
    }
    $Action  = New-ScheduledTaskAction -Execute "powershell.exe" `
                 -Argument "$TaskArgs `"$($Task.Script)`""
    $Trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Minutes 15) `
                 -Once -At (Get-Date)
    $Settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
                  -MultipleInstances IgnoreNew
    Register-ScheduledTask -TaskName $Task.Name -Action $Action -Trigger $Trigger `
        -Settings $Settings -RunLevel Highest -User "SYSTEM" | Out-Null
    Write-Host "  Scheduled Task '$($Task.Name)' erstellt." -ForegroundColor Green
}

# Erster Testlauf
Write-Host ""
Write-Host "  Fuehre ersten Testlauf durch ..." -ForegroundColor Cyan
& powershell.exe -NonInteractive -ExecutionPolicy Bypass -File $MonitorScript
Write-Host ""

# Zusammenfassung
Write-Header "Installation abgeschlossen"
Write-Host "  Installationspfad : $InstallPath" -ForegroundColor White
Write-Host "  Konfiguration     : $ConfigFile" -ForegroundColor White
Write-Host "  Schluesseldatei   : $KeyFile" -ForegroundColor White
Write-Host ""
Write-Host "  Scheduled Tasks (alle 15 Minuten):" -ForegroundColor White
Write-Host "    - MailStore Job Monitor" -ForegroundColor White
Write-Host "    - MailStore Job Eskalation" -ForegroundColor White
Write-Host ""
Write-Host "  Monitoring-Tool (RMM) einrichten:" -ForegroundColor Yellow
Write-Host "    Event Log : Application" -ForegroundColor Yellow
Write-Host "    Source    : MailStore Monitor" -ForegroundColor Yellow
Write-Host "    Event ID  : 9001 (Warning) -> Alarm ausloesen" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Zum Aendern der Konfiguration: Update-Config.ps1 ausfuehren" -ForegroundColor Cyan
Write-Host ""
