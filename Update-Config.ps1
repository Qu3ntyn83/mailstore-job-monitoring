#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Aendert einzelne Einstellungen in der MailStore Monitor Konfiguration.
    Ermoeglicht insbesondere das Aktualisieren des verschluesselten Passworts.
#>

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$InstallPath = Split-Path $MyInvocation.MyCommand.Path
$ConfigFile  = Join-Path $InstallPath "mailstore-config.xml"
$KeyFile     = Join-Path $InstallPath "mailstore-key.bin"

if (-not (Test-Path $ConfigFile) -or -not (Test-Path $KeyFile)) {
    Write-Error "Konfiguration nicht gefunden. Bitte zuerst Install.ps1 ausfuehren."
    exit 1
}

# Laden
$AesKey = [System.IO.File]::ReadAllBytes($KeyFile)
$Xml    = [xml](Get-Content $ConfigFile)
$Cfg    = $Xml.MailStoreMonitor

Write-Host ""
Write-Host "Aktuelle Konfiguration:" -ForegroundColor Cyan
Write-Host "  Server           : $($Cfg.MailStoreServer):$($Cfg.ApiPort)"
Write-Host "  API-Benutzer     : $($Cfg.ApiUser)"
Write-Host "  API-Passwort     : [verschluesselt]"
Write-Host "  Alarm nach       : $($Cfg.MaxRuntimeMinutes) Min."
Write-Host "  Eskalation nach  : $($Cfg.EscalateAfterMinutes) Min."
Write-Host "  Throttling       : $($Cfg.AlertIntervalMinutes) Min."
Write-Host "  Dienst           : $($Cfg.ServiceName)"
Write-Host ""

Write-Host "Was moechten Sie aendern?" -ForegroundColor Yellow
Write-Host "  1) Passwort aktualisieren"
Write-Host "  2) Server / Port aendern"
Write-Host "  3) Schwellenwerte aendern"
Write-Host "  4) Abbrechen"
Write-Host ""
$Choice = Read-Host "Auswahl (1-4)"

switch ($Choice) {
    "1" {
        $NewPass = Read-Host "Neues API-Passwort" -AsSecureString
        $PlainPass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($NewPass)
        )
        $SecureString = ConvertTo-SecureString $PlainPass -AsPlainText -Force
        $Encrypted = [string]($SecureString | ConvertFrom-SecureString -Key $AesKey)
        $Xml.SelectSingleNode("//ApiPasswordEncrypted").InnerText = $Encrypted
        Write-Host "Passwort aktualisiert." -ForegroundColor Green
    }
    "2" {
        $NewServer = (Read-Host "Server [$($Cfg.MailStoreServer)]").Trim()
        if ($NewServer) { $Xml.SelectSingleNode("//MailStoreServer").InnerText = $NewServer }
        $NewPort = (Read-Host "Port [$($Cfg.ApiPort)]").Trim()
        if ($NewPort) { $Xml.SelectSingleNode("//ApiPort").InnerText = $NewPort }
        Write-Host "Server/Port aktualisiert." -ForegroundColor Green
    }
    "3" {
        $NewMax = (Read-Host "Alarm nach Minuten [$($Cfg.MaxRuntimeMinutes)]").Trim()
        if ($NewMax) { $Xml.SelectSingleNode("//MaxRuntimeMinutes").InnerText = $NewMax }
        $NewEsc = (Read-Host "Eskalation nach Minuten [$($Cfg.EscalateAfterMinutes)]").Trim()
        if ($NewEsc) { $Xml.SelectSingleNode("//EscalateAfterMinutes").InnerText = $NewEsc }
        $NewThr = (Read-Host "Throttling Minuten [$($Cfg.AlertIntervalMinutes)]").Trim()
        if ($NewThr) { $Xml.SelectSingleNode("//AlertIntervalMinutes").InnerText = $NewThr }
        Write-Host "Schwellenwerte aktualisiert." -ForegroundColor Green
    }
    default {
        Write-Host "Abgebrochen." -ForegroundColor Yellow
        exit 0
    }
}

$Xml.Save($ConfigFile)
Write-Host "Konfiguration gespeichert: $ConfigFile" -ForegroundColor Green
