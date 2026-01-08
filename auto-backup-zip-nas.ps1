# ===============================================================
# auto-backup-zip-nas.ps1
# Windows Server 2016
# Compresión local + envío ZIP al NAS + correo + Telegram
# ===============================================================

$BackupPath = 
$NasPath    = 
$TempZipDir = 

# ================= SMTP =================
$SmtpServer = 
$MailFrom   = 
$MailTo     = 

# ================= TELEGRAM =================
$TelegramBotToken =
$TelegramChatId   =

# ================= Carpetas =================
if (-not (Test-Path $TempZipDir)) { New-Item -ItemType Directory -Path $TempZipDir | Out-Null }

$FechaActual  = Get-Date
$Mes          = $FechaActual.ToString("MMMM").ToUpper()
$Anio         = $FechaActual.Year
$FechaCarpeta = $FechaActual.ToString("ddMMyyyy")

$CarpetaMes   = Join-Path $NasPath "BACKUPS $Mes $Anio"
$CarpetaFecha = Join-Path $CarpetaMes $FechaCarpeta

if (-not (Test-Path $CarpetaMes))   { New-Item -ItemType Directory -Path $CarpetaMes   | Out-Null }
if (-not (Test-Path $CarpetaFecha)) { New-Item -ItemType Directory -Path $CarpetaFecha | Out-Null }

# ================= LOG =================
$LogFolder = "C:\Scripts\logs"
if (-not (Test-Path $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder | Out-Null }
$LogFile = Join-Path $LogFolder ("log-{0:yyyyMMdd}.txt" -f (Get-Date))

function Log {
    param([string]$msg)
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $LogFile -Value $line
    Write-Host $msg
}

function Send-Mail {
    param ([string]$Subject, [string]$Body)
    Send-MailMessage `
        -From $MailFrom `
        -To $MailTo `
        -Subject $Subject `
        -Body $Body `
        -BodyAsHtml `
        -SmtpServer $SmtpServer `
        -Port 25 `
        -Encoding UTF8
}

function Send-Telegram {
    param ([string]$Message)
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $uri = "https://api.telegram.org/bot$TelegramBotToken/sendMessage"
        $body = @{
            chat_id    = $TelegramChatId
            text       = $Message
            parse_mode = "Markdown"
        }
        Invoke-RestMethod -Uri $uri -Method Post -Body $body | Out-Null
        Log "Mensaje enviado a Telegram correctamente."
    } catch {
        Log "ERROR TELEGRAM: $($_.Exception.Message)"
    }
}

# ================= PROCESO =================
Log "=== INICIO BACKUP ==="

$BackupsOK    = @()
$BackupsError = @()

$BakFiles = Get-ChildItem -Path $BackupPath -Filter "*.bak" -File

foreach ($file in $BakFiles) {
    $BakFile = $file.FullName
    $TempZip = Join-Path $TempZipDir ($file.BaseName + ".zip")
    $DestZip = Join-Path $CarpetaFecha ($file.BaseName + ".zip")

    try {
        if (Test-Path $TempZip) { Remove-Item $TempZip -Force }

        Set-Content -Path $TempZip -Value ("PK" + [char]5 + [char]6 + ("`0" * 18))

        $shell     = New-Object -ComObject Shell.Application
        $zipFolder = $shell.NameSpace($TempZip)
        $srcFolder = $shell.NameSpace((Split-Path $BakFile))
        $srcItem   = $srcFolder.ParseName($file.Name)

        Log "Comprimiendo $($file.Name)..."
        $zipFolder.CopyHere($srcItem, 0x14)
        while ($zipFolder.Items().Count -lt 1) { Start-Sleep 2 }
        Start-Sleep 2

        Copy-Item $TempZip $DestZip -Force
        if (-not (Test-Path $DestZip)) { throw "No se copió el ZIP al NAS" }

        Remove-Item $BakFile -Force
        Remove-Item $TempZip -Force

        $BackupsOK += $file.Name
        Log "OK: $($file.Name)"
    } catch {
        $BackupsError += "$($file.Name) - $($_.Exception.Message)"
        Log "ERROR: $($file.Name)"
    }
}

# ================= CORREO / TELEGRAM =================
$Servidor = $env:COMPUTERNAME
$Fecha    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Cuerpo correo HTML
$BodyHtml = "<h2>Reporte Backup SQL</h2>"
$BodyHtml += "<p><b>Servidor:</b> $Servidor</p>"
$BodyHtml += "<p><b>Fecha:</b> $Fecha</p>"
$BodyHtml += "<p><b>Ruta NAS:</b><br>$CarpetaFecha</p>"
$BodyHtml += "<h3>Backups OK</h3><ul>"
foreach ($b in $BackupsOK) { $BodyHtml += "<li>$b</li>" }
$BodyHtml += "</ul><h3>Errores</h3><ul style='color:red'>"
foreach ($e in $BackupsError) { $BodyHtml += "<li>$e</li>" }
$BodyHtml += "</ul><p>Log: $LogFile</p>"

$Subject = if ($BackupsError.Count -eq 0) { "BACKUP SQL OK - $Servidor" } else { "BACKUP SQL CON ERRORES - $Servidor" }

# Enviar correo
Send-Mail -Subject $Subject -Body $BodyHtml

# Cuerpo Telegram legible con Markdown
$BodyTelegram = "*Reporte Backup SQL*`n"
$BodyTelegram += "*Servidor:* $Servidor`n"
$BodyTelegram += "*Fecha:* $Fecha`n"
$BodyTelegram += "*Ruta NAS:* $CarpetaFecha`n`n"

$BodyTelegram += "*Backups OK:*`n"
if ($BackupsOK.Count -eq 0) { $BodyTelegram += "- Ninguno`n" } else { foreach ($b in $BackupsOK) { $BodyTelegram += "- $b`n" } }

$BodyTelegram += "`n*Errores:*`n"
if ($BackupsError.Count -eq 0) { $BodyTelegram += "- Ninguno`n" } else { foreach ($e in $BackupsError) { $BodyTelegram += "- $e`n" } }

$BodyTelegram += "`n_Log: $LogFile"

# Enviar Telegram
Send-Telegram -Message $BodyTelegram

Log "=== FIN BACKUP ==="
Write-Host "Proceso completado correctamente."
