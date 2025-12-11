# ===============================================================
# auto-backup-zip-nas.ps1
# Windows Server 2016 - Compresión local + envío ZIP al NAS
# Usa ZIP nativo (Shell.Application) compatible con archivos grandes
# ===============================================================

$BackupPath = "C:\Program Files\Microsoft SQL Server\MSSQL14.SERQPRDDB\MSSQL\Backup"
$NasPath    = "\x\Respaldos"
$TempZipDir = "C:\Scripts\tempzip"

# Crear carpeta temporal
if (-not (Test-Path $TempZipDir)) {
    New-Item -Path $TempZipDir -ItemType Directory | Out-Null
}

# Carpetas en NAS
$FechaActual = Get-Date
$Mes = $FechaActual.ToString("MMMM").ToUpper()
$Anio = $FechaActual.Year
$FechaCarpeta = $FechaActual.ToString("ddMMyyyy")

$CarpetaMes   = Join-Path $NasPath "BACKUPS $Mes $Anio"
$CarpetaFecha = Join-Path $CarpetaMes $FechaCarpeta

if (-not (Test-Path $CarpetaMes))   { New-Item -ItemType Directory -Path $CarpetaMes   | Out-Null }
if (-not (Test-Path $CarpetaFecha)) { New-Item -ItemType Directory -Path $CarpetaFecha | Out-Null }

# LOG
$LogFolder = "C:\Scripts\logs"
if (-not (Test-Path $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder | Out-Null }
$LogFile = Join-Path $LogFolder ("log-{0:yyyyMMdd}.txt" -f (Get-Date))

function Log {
    param([string]$msg)
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $LogFile -Value $line
    Write-Host $msg
}

Log "=== INICIO ==="

$BakFiles = Get-ChildItem -Path $BackupPath -Filter "*.bak" -File

foreach ($file in $BakFiles) {

    $BakFile = $file.FullName
    $TempZip = Join-Path $TempZipDir ($file.BaseName + ".zip")
    $DestZip = Join-Path $CarpetaFecha ($file.BaseName + ".zip")

    $FileSizeGB = [math]::Round($file.Length / 1GB, 2)
    Log "COMPRIMIENDO EN SERVIDOR: $($file.Name) ($FileSizeGB GB)"

    try {

        if (Test-Path $TempZip) { Remove-Item $TempZip -Force }

        # Crear ZIP vacío
        Set-Content -Path $TempZip -Value ("PK" + [char]5 + [char]6 + ("`0" * 18))

        # Shell zip
        $shell     = New-Object -ComObject Shell.Application
        $zipFolder = $shell.NameSpace($TempZip)
        $srcFolder = $shell.NameSpace((Split-Path $BakFile))
        $srcItem   = $srcFolder.ParseName($file.Name)

        Log "   → Comprimir localmente... esto puede tardar..."

        $zipFolder.CopyHere($srcItem, 0x14)

        while ($zipFolder.Items().Count -lt 1) { Start-Sleep -Seconds 2 }
        Start-Sleep -Seconds 3

        if ((Get-Item $TempZip).Length -eq 0) {
            throw "ZIP vacío."
        }

        Log "   ✔ ZIP creado local: $TempZip"

        Log "   → Enviando ZIP al NAS..."
        Copy-Item -Path $TempZip -Destination $DestZip -Force

        Log "✔ OK: $($file.Name) → $DestZip"

        Remove-Item $TempZip -Force
    }
    catch {
        Log "ERROR: $($file.Name) - $($_.Exception.Message)"
    }

} # cierre foreach

Log "=== FIN ==="
Write-Host ("Proceso completado. Log: " + $LogFile)
