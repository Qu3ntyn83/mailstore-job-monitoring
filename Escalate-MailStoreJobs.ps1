#Requires -Version 5.1
<#
.SYNOPSIS
    Eskalations-Workflow fuer haengende MailStore-Jobs.

.NOTES
    Ablauf:
      Stufe 1: Profil auf manuell setzen
      Stufe 2: CancelJobAsync
               OK + Job bestaetigt -> Profil automatisch -> Fertig
               nOK -> weiter
      Stufe 3: Stop-Service (max. 180 Sek.)
               OK -> Start-Service -> Profil automatisch -> Fertig
               Timeout -> taskkill -> warten -> Service pruefen
                 Service gestoppt -> Start-Service -> Profil automatisch -> Fertig
                 Service laeuft noch -> Restart-Service -> Profil automatisch -> Fertig
                 Alles fehlgeschlagen -> Event 9015 (manuell)

    Event IDs:
      9010 - Information : Eskalation gestartet
      9011 - Information : Job per API gestoppt, Profil automatisch
      9012 - Warning     : Dienst gestoppt (Stufe 3)
      9013 - Warning     : Prozess per taskkill beendet (Stufe 3 Fallback)
      9014 - Information : Dienst gestartet, Profil automatisch
      9015 - Error       : Eskalation fehlgeschlagen, manueller Eingriff noetig
      9016 - Warning     : Profil auf manuell gesetzt, Abbruch laeuft
#>

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$OutputEncoding = [System.Text.Encoding]::UTF8

# ── Konfiguration laden ───────────────────────────────────────────────────────

$ConfigFile          = Join-Path $PSScriptRoot "mailstore-config.xml"
$KeyFile             = Join-Path $PSScriptRoot "mailstore-key.bin"
$StateFile           = Join-Path $PSScriptRoot "mailstore_alert_state.json"
$EscalationStateFile = Join-Path $PSScriptRoot "mailstore_escalation_state.json"

if (-not (Test-Path $ConfigFile) -or -not (Test-Path $KeyFile)) {
    Write-Error "Konfigurationsdatei nicht gefunden. Bitte Install.ps1 ausfuehren."
    exit 1
}

$AesKey = [System.IO.File]::ReadAllBytes($KeyFile)
$Xml    = [xml](Get-Content $ConfigFile -Raw)
$Cfg    = $Xml.MailStoreMonitor

$SecurePass  = ConvertTo-SecureString ([string]$Cfg.ApiPasswordEncrypted) -Key $AesKey
$ApiPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePass)
)

$MailStoreServer        = $Cfg.MailStoreServer
$ApiPort                = [int]$Cfg.ApiPort
$ApiUser                = $Cfg.ApiUser
$EscalateAfterMinutes   = [int]$Cfg.EscalateAfterMinutes
$PauseBetweenJobs       = [int]$Cfg.PauseBetweenJobs
$ServiceName            = $Cfg.ServiceName
$ProcessName            = $Cfg.ProcessName
$StopServiceTimeoutSec  = 180
$StartServiceTimeoutSec = [int]$Cfg.StartServiceTimeoutSec
$EventLogName           = $Cfg.EventLogName
$EventLogSource         = $Cfg.EventLogSource

$CancelWaitSec          = 30   # Warten nach CancelJobAsync
$TaskkillWaitSec        = 30   # Warten nach taskkill

# ── Funktionen ────────────────────────────────────────────────────────────────

function ts { Get-Date -f 'HH:mm:ss' }

function Write-MonitorEvent {
    param([int]$EventId, [System.Diagnostics.EventLogEntryType]$EntryType, [string]$Message)
    try {
        Write-EventLog -LogName $EventLogName -Source $EventLogSource `
            -EventId $EventId -EntryType $EntryType -Message $Message -ErrorAction Stop
        Write-Host "$(ts) [EVENTLOG] ID $EventId ($EntryType): $($Message.Split([char]10)[0])"
    }
    catch { Write-Warning "Event Log Schreibfehler: $_" }
}

function Invoke-MailStoreApi {
    param([string]$Method, [string]$Body = "")
    $CredB64 = [Convert]::ToBase64String(
        [System.Text.Encoding]::ASCII.GetBytes("${ApiUser}:${ApiPassword}")
    )
    $Url = "https://${MailStoreServer}:${ApiPort}/api/invoke/${Method}"
    try {
        if ($Body) {
            $Raw = curl.exe -k -s -X POST -H "Authorization: Basic $CredB64" -H "Content-Type: application/x-www-form-urlencoded" -d $Body $Url 2>&1
        } else {
            $Raw = curl.exe -k -s -X POST -H "Authorization: Basic $CredB64" -H "Content-Length: 0" $Url 2>&1
        }
        if (-not $Raw) { return $null }
        return (($Raw | Out-String).TrimStart([char]0xFEFF).Trim() | ConvertFrom-Json)
    }
    catch { return $null }
}

function Get-MonitorState {
    if (Test-Path $StateFile) {
        try { return (Get-Content $StateFile -Raw | ConvertFrom-Json) } catch { }
    }
    return $null
}

function Get-EscalationState {
    if (Test-Path $EscalationStateFile) {
        try { return (Get-Content $EscalationStateFile -Raw | ConvertFrom-Json) } catch { }
    }
    return [PSCustomObject]@{ LastEscalationUtc = $null; EscalationDone = $false }
}

function Save-EscalationState { param([bool]$Done)
    @{ LastEscalationUtc = (Get-Date).ToUniversalTime().ToString("o"); EscalationDone = $Done } |
        ConvertTo-Json | Set-Content $EscalationStateFile
}

function Set-ProfileManual {
    param([int]$Id)
    foreach ($M in @("SetProfileServerSideExecution", "SetProfileSchedule", "SetServerSideExecution")) {
        $R = Invoke-MailStoreApi -Method $M -Body "id=$Id&automatic=false"
        if ($R -and $R.statusCode -eq "succeeded") {
            Write-Host "$(ts) [API] Profil $Id -> manuell ($M)."
            return $M
        }
    }
    Write-Host "$(ts) [API] Profil-Manuell-Methode nicht gefunden - ohne gesetzt."
    return $null
}

function Set-ProfileAutomatic {
    param([int]$Id, [string]$Method)
    if (-not $Method) {
        Write-Host "$(ts) [API] Profil $Id bitte manuell im MailStore Client auf automatisch setzen!"
        return
    }
    $PauseSec = $PauseBetweenJobs * 60
    $R = Invoke-MailStoreApi -Method $Method -Body "id=$Id&automatic=true&automaticPauseBetweenExecutions=$PauseSec"
    if ($R -and $R.statusCode -eq "succeeded") {
        Write-Host "$(ts) [API] Profil $Id -> automatisch (Pause: $PauseSec Sek.)."
    } else {
        Write-Host "$(ts) [API] WARNUNG: Profil $Id nicht auf automatisch gesetzt!"
        Write-Host "$(ts) [API] Bitte manuell im MailStore Client korrigieren."
    }
}

function Test-JobStopped {
    param([DateTime]$Since)
    $From = $Since.ToString("yyyy-MM-ddTHH:mm:ss")
    $To   = (Get-Date).AddHours(1).ToString("yyyy-MM-ddTHH:mm:ss")
    $R    = Invoke-MailStoreApi -Method "GetWorkerResults" -Body "fromIncluding=$From&toExcluding=$To&timeZoneID=`$Local"
    return ($R -and $R.statusCode -eq "succeeded" -and @($R.result).Count -gt 0)
}

function Get-ServiceStatus {
    try { return (Get-Service -Name $ServiceName -ErrorAction Stop).Status }
    catch { return "Unknown" }
}

function Start-MSService {
    Write-Host "$(ts) [SVC] Starte '$ServiceName' ..."
    try {
        Start-Service -Name $ServiceName -ErrorAction Stop
        $Waited = 0
        while ($Waited -lt $StartServiceTimeoutSec) {
            Start-Sleep -Seconds 5; $Waited += 5
            if ((Get-ServiceStatus) -eq "Running") {
                Write-Host "$(ts) [SVC] Dienst laeuft nach $Waited Sek."
                return $true
            }
        }
        Write-Host "$(ts) [SVC] Dienst-Start Timeout nach $StartServiceTimeoutSec Sek."
        return $false
    }
    catch { Write-Warning "Start-Service Fehler: $_"; return $false }
}

function Stop-MSService {
    Write-Host "$(ts) [SVC] Stoppe '$ServiceName' (max. $StopServiceTimeoutSec Sek.) ..."
    try {
        if ((Get-ServiceStatus) -eq "Stopped") {
            Write-Host "$(ts) [SVC] Bereits gestoppt."
            return $true
        }
        Stop-Service -Name $ServiceName -Force -ErrorAction Stop
        $Waited = 0
        while ($Waited -lt $StopServiceTimeoutSec) {
            Start-Sleep -Seconds 10; $Waited += 10
            $Status = Get-ServiceStatus
            Write-Host "$(ts) [SVC] Warte ... $Waited/$StopServiceTimeoutSec Sek. (Status: $Status)"
            if ($Status -eq "Stopped") {
                Write-Host "$(ts) [SVC] Gestoppt nach $Waited Sek."
                return $true
            }
        }
        Write-Host "$(ts) [SVC] Stop-Service Timeout nach $StopServiceTimeoutSec Sek."
        return $false
    }
    catch { Write-Warning "Stop-Service Fehler: $_"; return $false }
}

function Invoke-Taskkill {
    Write-Host "$(ts) [KILL] taskkill /F /IM $ProcessName.exe /T ..."
    try {
        $Out = taskkill.exe /F /IM "$ProcessName.exe" /T 2>&1
        Write-Host "$(ts) [KILL] $Out"
        Write-Host "$(ts) [KILL] Warte $TaskkillWaitSec Sek. ..."
        Start-Sleep -Seconds $TaskkillWaitSec
    }
    catch { Write-Warning "taskkill Fehler: $_" }
}

function Invoke-ServiceRestart {
    Write-Host "$(ts) [SVC] Restart: Stop dann Start ..."
    Stop-MSService | Out-Null
    Start-Sleep -Seconds 10
    return (Start-MSService)
}

function Complete-Escalation {
    param([int]$ProfileId, [string]$Method, [string]$Reason)
    Write-Host "$(ts) ========================================"
    Write-Host "$(ts) Eskalation abgeschlossen: $Reason"
    Write-Host "$(ts) ========================================"
    Set-ProfileAutomatic -ProfileId $ProfileId -Method $Method
    Write-MonitorEvent -EventId 9014 -EntryType Information -Message "MailStore Eskalation abgeschlossen.
Server : $MailStoreServer
Grund  : $Reason
Profil : $ProfileId zurueck auf automatisch."
    Save-EscalationState -Done $true
}

# ── Hauptlogik ────────────────────────────────────────────────────────────────

Write-Host "$(ts) MailStore Eskalations-Monitor gestartet."

$MonitorState = Get-MonitorState
if (-not $MonitorState -or -not $MonitorState.LastCompleteTimeUtc) {
    Write-Host "$(ts) Kein Monitor-State vorhanden."
    exit 0
}

$LastCompleteLocal    = [DateTime]::Parse($MonitorState.LastCompleteTimeUtc).ToLocalTime()
$MinutesSinceComplete = [Math]::Round(((Get-Date) - $LastCompleteLocal).TotalMinutes, 1)
$EstimatedRuntime     = [Math]::Round($MinutesSinceComplete - $PauseBetweenJobs, 1)
$ProfileId            = [int]$MonitorState.AlertedProfileId

Write-Host "$(ts) Seit Abschluss  : $MinutesSinceComplete Min."
Write-Host "$(ts) Gesch. Laufzeit : $EstimatedRuntime Min. (Eskalation ab: $EscalateAfterMinutes Min.)"

if ($MinutesSinceComplete -le ($EscalateAfterMinutes + $PauseBetweenJobs)) {
    Write-Host "$(ts) [OK] Keine Eskalation notwendig."
    if (Test-Path $EscalationStateFile) { Remove-Item $EscalationStateFile -Force }
    exit 0
}

$EscState = Get-EscalationState
if ($EscState.EscalationDone) {
    $LastEsc = [DateTime]::Parse($EscState.LastEscalationUtc).ToLocalTime()
    Write-Host "$(ts) [INFO] Eskalation bereits um $($LastEsc.ToString('HH:mm:ss')) durchgefuehrt."
    exit 0
}

if ($ProfileId -le 0) {
    Write-Warning "Keine Profil-ID im State."
    exit 1
}

Write-Host "$(ts) ========================================"
Write-Host "$(ts) ESKALATION GESTARTET"
Write-Host "$(ts) ========================================"

Write-MonitorEvent -EventId 9010 -EntryType Information -Message "MailStore Eskalation gestartet.
Server         : $MailStoreServer
Profil-ID      : $ProfileId
Seit Abschluss : $MinutesSinceComplete Min.
Gesch. Laufzeit: ca. $EstimatedRuntime Min."

# ═══════════════════════════════════════════════════════════
# STUFE 1: Profil auf manuell
# ═══════════════════════════════════════════════════════════
Write-Host "$(ts) ── STUFE 1: Profil auf manuell ──"
$WorkingMethod = Set-ProfileManual -ProfileId $ProfileId

Write-MonitorEvent -EventId 9016 -EntryType Warning -Message "Profil $ProfileId auf manuell gesetzt.
Server: $MailStoreServer | Gesch. Laufzeit: ca. $EstimatedRuntime Min."

# ═══════════════════════════════════════════════════════════
# STUFE 2: CancelJobAsync
# ═══════════════════════════════════════════════════════════
Write-Host "$(ts) ── STUFE 2: CancelJobAsync ──"
$CancelResult = Invoke-MailStoreApi -Method "CancelJobAsync" -Body "id=$ProfileId"

if ($CancelResult -and $CancelResult.statusCode -eq "succeeded") {
    Write-Host "$(ts) CancelJobAsync: succeeded - warte $CancelWaitSec Sek. ..."
    Start-Sleep -Seconds $CancelWaitSec

    if (Test-JobStopped -Since $LastCompleteLocal) {
        Write-Host "$(ts) Job bestaetigt abgeschlossen."
        Write-MonitorEvent -EventId 9011 -EntryType Information -Message "Job per API gestoppt (CancelJobAsync).
Server    : $MailStoreServer | Profil: $ProfileId | Laufzeit: ca. $EstimatedRuntime Min."
        Complete-Escalation -ProfileId $ProfileId -Method $WorkingMethod -Reason "CancelJobAsync erfolgreich"
        exit 0
    }
    Write-Host "$(ts) Kein Abschluss-Eintrag nach Warten - weiter mit Stufe 3."
} else {
    Write-Host "$(ts) CancelJobAsync: $($CancelResult.error.message) - weiter mit Stufe 3."
}

# ═══════════════════════════════════════════════════════════
# STUFE 3: Stop-Service
# ═══════════════════════════════════════════════════════════
Write-Host "$(ts) ── STUFE 3: Stop-Service ──"
$ServiceStopped = Stop-MSService

if ($ServiceStopped) {
    Write-MonitorEvent -EventId 9012 -EntryType Warning -Message "MailStore Dienst gestoppt (Stop-Service).
Server: $MailStoreServer | Dienst: $ServiceName"

    Write-Host "$(ts) Warte 15 Sek. vor Start ..."
    Start-Sleep -Seconds 15

    if (Start-MSService) {
        Complete-Escalation -ProfileId $ProfileId -Method $WorkingMethod -Reason "Stop-Service + Start-Service erfolgreich"
        exit 0
    } else {
        Write-MonitorEvent -EventId 9015 -EntryType Error -Message "Dienst nicht gestartet nach Stop-Service!
Server: $MailStoreServer | Manueller Eingriff erforderlich!"
        Save-EscalationState -Done $true
        exit 1
    }
}

# ═══════════════════════════════════════════════════════════
# STUFE 3 Fallback: taskkill
# ═══════════════════════════════════════════════════════════
Write-Host "$(ts) ── STUFE 3 Fallback: taskkill ──"
Invoke-Taskkill

Write-MonitorEvent -EventId 9013 -EntryType Warning -Message "MailStore Prozess per taskkill beendet.
Server: $MailStoreServer | Prozess: $ProcessName.exe"

# Service-Status nach taskkill pruefen
$StatusAfterKill = Get-ServiceStatus
Write-Host "$(ts) Service-Status nach taskkill: $StatusAfterKill"

if ($StatusAfterKill -eq "Stopped") {
    # Erwarteter Zustand - sauber starten
    Write-Host "$(ts) Dienst gestoppt - starte neu ..."
    if (Start-MSService) {
        Complete-Escalation -ProfileId $ProfileId -Method $WorkingMethod -Reason "taskkill + Start-Service erfolgreich"
        exit 0
    } else {
        Write-MonitorEvent -EventId 9015 -EntryType Error -Message "Dienst nicht gestartet nach taskkill!
Server: $MailStoreServer | Manueller Eingriff erforderlich!"
        Save-EscalationState -Done $true
        exit 1
    }
} elseif ($StatusAfterKill -eq "Running") {
    # Dienst hat sich selbst neu gestartet oder SCM hat ihn neugestartet
    Write-Host "$(ts) Dienst laeuft bereits wieder - Restart durchfuehren ..."
    if (Invoke-ServiceRestart) {
        Complete-Escalation -ProfileId $ProfileId -Method $WorkingMethod -Reason "Restart-Service nach taskkill erfolgreich"
        exit 0
    } else {
        Write-MonitorEvent -EventId 9015 -EntryType Error -Message "Restart-Service nach taskkill fehlgeschlagen!
Server: $MailStoreServer | Manueller Eingriff erforderlich!"
        Save-EscalationState -Done $true
        exit 1
    }
} else {
    # Unbekannter Zustand
    Write-MonitorEvent -EventId 9015 -EntryType Error -Message "Unbekannter Service-Status nach taskkill: $StatusAfterKill
Server: $MailStoreServer | Manueller Eingriff erforderlich!"
    Save-EscalationState -Done $true
    exit 1
}
