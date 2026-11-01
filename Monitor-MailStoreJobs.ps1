#Requires -Version 5.1
<#
.SYNOPSIS
    Ueberwacht MailStore-Archivierungsjobs und schreibt Ergebnisse in den Windows Event Log.

.NOTES
    Konfiguration wird aus mailstore-config.xml im Script-Verzeichnis gelesen.
    Passwort ist AES-verschluesselt (Schluessel in mailstore-key.bin).
    Event IDs:
      9000 - Information : OK
      9001 - Warning     : Job laeuft zu lange (Alarm fuer das Monitoring-Tool)
      9002 - Error       : API-Fehler
      9003 - Information : Throttling aktiv
#>

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$OutputEncoding = [System.Text.Encoding]::UTF8

# ── Konfiguration laden ───────────────────────────────────────────────────────

$ConfigFile = Join-Path $PSScriptRoot "mailstore-config.xml"
$KeyFile    = Join-Path $PSScriptRoot "mailstore-key.bin"
$StateFile  = Join-Path $PSScriptRoot "mailstore_alert_state.json"

if (-not (Test-Path $ConfigFile) -or -not (Test-Path $KeyFile)) {
    Write-Error "Konfigurationsdatei nicht gefunden. Bitte Install.ps1 ausfuehren."
    exit 1
}

$AesKey = [System.IO.File]::ReadAllBytes($KeyFile)
$Xml    = [xml](Get-Content $ConfigFile -Raw)
$Cfg    = $Xml.MailStoreMonitor

# Passwort entschluesseln
$SecurePass = ConvertTo-SecureString ([string]$Cfg.ApiPasswordEncrypted) -Key $AesKey
$ApiPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePass)
)

$MailStoreServer      = $Cfg.MailStoreServer
$ApiPort              = [int]$Cfg.ApiPort
$ApiUser              = $Cfg.ApiUser
$MaxRuntimeMinutes    = [int]$Cfg.MaxRuntimeMinutes
$PauseBetweenJobs     = [int]$Cfg.PauseBetweenJobs
$LookbackHours        = [int]$Cfg.LookbackHours
$AlertIntervalMinutes = [int]$Cfg.AlertIntervalMinutes
$EventLogName         = $Cfg.EventLogName
$EventLogSource       = $Cfg.EventLogSource

# ── Funktionen ────────────────────────────────────────────────────────────────

function ts { Get-Date -f 'HH:mm:ss' }

function Initialize-EventLogSource {
    if (-not [System.Diagnostics.EventLog]::SourceExists($EventLogSource)) {
        try {
            New-EventLog -LogName $EventLogName -Source $EventLogSource -ErrorAction Stop
            Write-Host "$(ts) [INIT] Event Log Source '$EventLogSource' angelegt."
        }
        catch { Write-Warning "Event Log Source nicht anlegbar: $_" }
    }
}

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
    $CredB64 = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("${ApiUser}:${ApiPassword}"))
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

function Get-State {
    if (Test-Path $StateFile) {
        try { return (Get-Content $StateFile -Raw | ConvertFrom-Json) } catch { }
    }
    return [PSCustomObject]@{ LastAlertUtc = $null; AlertedProfileId = $null; LastCompleteTimeUtc = $null }
}

function Save-AlertState { param($ProfileId, $CompleteTimeUtc)
    @{ LastAlertUtc = (Get-Date).ToUniversalTime().ToString("o"); AlertedProfileId = $ProfileId; LastCompleteTimeUtc = $CompleteTimeUtc } |
        ConvertTo-Json | Set-Content $StateFile
}

function Update-LastComplete { param($CompleteTimeUtc)
    $s = Get-State
    @{ LastAlertUtc = $s.LastAlertUtc; AlertedProfileId = $s.AlertedProfileId; LastCompleteTimeUtc = $CompleteTimeUtc } |
        ConvertTo-Json | Set-Content $StateFile
}

# ── Start ─────────────────────────────────────────────────────────────────────

Write-Host "$(ts) MailStore Job Monitor gestartet auf ${MailStoreServer}:${ApiPort}"
Initialize-EventLogSource

# ── API-Abfrage ───────────────────────────────────────────────────────────────

$QueryBody = "fromIncluding=$((Get-Date).AddHours(-$LookbackHours).ToString('yyyy-MM-ddTHH:mm:ss'))&toExcluding=$((Get-Date).ToString('yyyy-MM-ddTHH:mm:ss'))&timeZoneID=`$Local"
$ApiResult = Invoke-MailStoreApi -Method "GetWorkerResults" -Body $QueryBody

if (-not $ApiResult -or $ApiResult.statusCode -ne "succeeded") {
    $ErrMsg = "MailStore API nicht erreichbar.`nServer: ${MailStoreServer}:${ApiPort}`nFehler: $($ApiResult.error.message)"
    Write-Warning $ErrMsg
    Write-MonitorEvent -EventId 9002 -EntryType Error -Message $ErrMsg
    exit 1
}

# ── Jobs auswerten ────────────────────────────────────────────────────────────

$Jobs = @($ApiResult.result)

if ($Jobs.Count -gt 0) {
    $LatestJob = $Jobs | Sort-Object { [DateTime]::Parse($_.completeTime) } -Descending | Select-Object -First 1
    $LastComplete = [DateTime]::Parse($LatestJob.completeTime)
    $ProfileId    = $LatestJob.profileId
    Write-Host "$(ts) Letzter Job   : '$($LatestJob.profileName)' (Profil $ProfileId)"
    Write-Host "$(ts) Abgeschlossen : $($LastComplete.ToString('dd.MM.yyyy HH:mm:ss'))"
    Update-LastComplete -CompleteTimeUtc $LastComplete.ToUniversalTime().ToString("o")
} else {
    Write-Host "$(ts) Keine abgeschlossenen Jobs im Fenster (letzte ${LookbackHours}h)."
}

# ── Zeitbasierte Pruefung ─────────────────────────────────────────────────────

$State = Get-State

if (-not $State.LastCompleteTimeUtc) {
    $Msg = "MailStore Monitor: Erster Lauf auf $MailStoreServer - State initialisiert."
    Write-Host "$(ts) [INIT] $Msg"
    Write-MonitorEvent -EventId 9000 -EntryType Information -Message $Msg
    exit 0
}

$LastCompleteLocal    = [DateTime]::Parse($State.LastCompleteTimeUtc).ToLocalTime()
$MinutesSinceComplete = [Math]::Round(((Get-Date) - $LastCompleteLocal).TotalMinutes, 1)
$ThresholdMinutes     = $MaxRuntimeMinutes + $PauseBetweenJobs
$EstimatedRuntime     = [Math]::Round($MinutesSinceComplete - $PauseBetweenJobs, 1)

Write-Host "$(ts) Seit Abschluss : $MinutesSinceComplete Min. (Schwelle: $ThresholdMinutes Min.)"

if ($MinutesSinceComplete -le $ThresholdMinutes) {
    $OkMsg = "MailStore Monitor: OK`nServer: $MailStoreServer`nLetzter Abschluss: $($LastCompleteLocal.ToString('dd.MM.yyyy HH:mm:ss'))`nSeit Abschluss: $MinutesSinceComplete Min."
    Write-Host "$(ts) [OK] Job laeuft im Normalbereich."
    Write-MonitorEvent -EventId 9000 -EntryType Information -Message $OkMsg
    exit 0
}

# ── Throttling ────────────────────────────────────────────────────────────────

if ($AlertIntervalMinutes -gt 0 -and $State.LastAlertUtc) {
    $LastAlert        = [DateTime]::Parse($State.LastAlertUtc).ToLocalTime()
    $MinutesSinceLast = [Math]::Round(((Get-Date) - $LastAlert).TotalMinutes, 1)
    if ($MinutesSinceLast -lt $AlertIntervalMinutes) {
        $NextIn  = [Math]::Round($AlertIntervalMinutes - $MinutesSinceLast, 0)
        $SkipMsg = "MailStore Monitor: Throttling aktiv`nServer: $MailStoreServer`nNaechster Alarm in $NextIn Min."
        Write-Host "$(ts) [SKIP] Naechster Alarm in $NextIn Min."
        Write-MonitorEvent -EventId 9003 -EntryType Information -Message $SkipMsg
        exit 0
    }
}

# ── Alarm ─────────────────────────────────────────────────────────────────────

$AlarmMsg = "MailStore Monitor: ALARM - Job laeuft zu lange
Server           : $MailStoreServer
Letzter Abschluss: $($LastCompleteLocal.ToString('dd.MM.yyyy HH:mm:ss'))
Gesch. Jobstart  : $($LastCompleteLocal.AddMinutes($PauseBetweenJobs).ToString('dd.MM.yyyy HH:mm:ss'))
Gesch. Laufzeit  : ca. $EstimatedRuntime Min. (Schwelle: $MaxRuntimeMinutes Min.)
Seit Abschluss   : $MinutesSinceComplete Min.
Alarmzeit        : $(Get-Date -f 'dd.MM.yyyy HH:mm:ss')"

Write-Host "$(ts) [ALARM] Job laeuft ca. $EstimatedRuntime Min."
Write-MonitorEvent -EventId 9001 -EntryType Warning -Message $AlarmMsg
Save-AlertState -ProfileId $ProfileId -CompleteTimeUtc $State.LastCompleteTimeUtc
exit 0
