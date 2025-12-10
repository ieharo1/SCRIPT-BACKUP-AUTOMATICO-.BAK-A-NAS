# ----------------- CONFIGURACIÓN -----------------

# Carpeta donde SQL Server guarda los .bak
$origen = "C:\"

# Ruta base del NAS
$destinoNAS = "\\182.1.1.23"

# Carpeta temporal para crear ZIPs
$tempZip = "C:\TempZip"
if (!(Test-Path $tempZip)) { New-Item -ItemType Directory -Path $tempZip | Out-Null }

# ----------------- FECHAS DINÁMICAS -----------------

$hoy = Get-Date
$mesNombre = $hoy.ToString("MMMM", [System.Globalization.CultureInfo]::CreateSpecificCulture("es-ES")).ToUpper()
$anio = $hoy.ToString("yyyy")
$carpetaMes = "BACKUPS $mesNombre $anio"

$fechaCarpeta = $hoy.ToString("ddMMyyyy")

# Ruta final donde se copiarán los ZIPs
$rutaMes = Join-Path $destinoNAS $carpetaMes
$rutaDia = Join-Path $rutaMes $fechaCarpeta

# Crear carpetas en el NAS si no existen
if (!(Test-Path $rutaMes)) { New-Item -ItemType Directory -Path $rutaMes | Out-Null }
if (!(Test-Path $rutaDia)) { New-Item -ItemType Directory -Path $rutaDia | Out-Null }

# ----------------- PROCESAR BACKUPS -----------------

Get-ChildItem -Path $origen -Filter "*.bak" | ForEach-Object {

    $bakFile = $_.FullName
    $fileName = $_.BaseName  # Sin extensión
    $zipFile = "$tempZip\$fileName.zip"

    # Crear ZIP con el mismo nombre
    Compress-Archive -Path $bakFile -DestinationPath $zipFile -Force

    # Copiar el ZIP al NAS
    Copy-Item -Path $zipFile -Destination $rutaDia -Force
}

# ----------------- LIMPIEZA TEMPORAL -----------------
Remove-Item "$tempZip\*.zip" -Force
