param(
    [int]$Fila = 0,
    [switch]$Todas
)

$ErrorActionPreference = "Stop"

# ==========================================================
# CRONOMETRO TOTAL
# ==========================================================

$CronometroTotal = [System.Diagnostics.Stopwatch]::StartNew()

$CronometroProceso = New-Object System.Diagnostics.Stopwatch

$FechaInicio = Get-Date


# ==========================================================
# CONFIGURACION
# ==========================================================

$Base = "C:\AutomatizacionAzure"

$IniciarPortal = Join-Path `
    $Base `
    "IniciarPortalTemporal.ps1"

$Proceso = Join-Path `
    $Base `
    "ProcesoUsuariosV4.ps1"

$CerrarPortal = Join-Path `
    $Base `
    "CerrarPortalTemporal.py"

$PerfilesTemporales = Join-Path `
    $Base `
    "PerfilesTemporales"

$LogsPath = Join-Path `
    $Base `
    "Logs"

$PowerShellLimpio = Join-Path `
    $env:SystemRoot `
    "System32\WindowsPowerShell\v1.0\powershell.exe"


# ==========================================================
# VALIDACION
# ==========================================================

if ($Todas -and $Fila -gt 0) {

    throw "Usa -Fila o -Todas, no ambos."

}

if (-not $Todas -and $Fila -lt 3) {

    throw "Indica -Fila N o usa -Todas."

}


# ==========================================================
# CREAR CARPETAS
# ==========================================================

New-Item `
    -ItemType Directory `
    -Path $PerfilesTemporales `
    -Force | Out-Null

New-Item `
    -ItemType Directory `
    -Path $LogsPath `
    -Force | Out-Null


# ==========================================================
# IDENTIFICADOR DE EJECUCION
# ==========================================================

$RunId = (
    Get-Date -Format "yyyyMMdd_HHmmss"
) + "_" + (
    [guid]::NewGuid().ToString("N").Substring(0,8)
)

$PerfilTemporal = Join-Path `
    $PerfilesTemporales `
    $RunId

New-Item `
    -ItemType Directory `
    -Path $PerfilTemporal `
    -Force | Out-Null

$env:AZURE_AUTOMATION_PROFILE = $PerfilTemporal

$ResultadoProceso = Join-Path `
    $PerfilTemporal `
    "resultado_proceso.txt"

$env:AZURE_AUTOMATION_RESULT = $ResultadoProceso


# ==========================================================
# CONTROL DE RESULTADO
# ==========================================================

$huboError = $false

$huboAdvertencias = $false

$MensajeError = ""

$MensajeAdvertencia = ""

Remove-Item Env:\AZURE_AUTOMATION_ROW_WARNINGS `
    -ErrorAction SilentlyContinue


try {

    # ======================================================
    # INFORMACION INICIAL
    # ======================================================

    Write-Host ""
    Write-Host "========================================"
    Write-Host "INICIO DE AUTOMATIZACION"
    Write-Host "========================================"

    Write-Host "Fecha inicio:"
    Write-Host $FechaInicio

    Write-Host ""
    Write-Host "Ejecucion:"
    Write-Host $RunId


    # ======================================================
    # INICIAR CHROME
    # ======================================================

    Write-Host ""
    Write-Host "========================================"
    Write-Host "INICIANDO SESION TEMPORAL AZURE"
    Write-Host "========================================"

    & $IniciarPortal


    # ======================================================
    # CRONOMETRO DEL PROCESAMIENTO
    # ======================================================

    $CronometroProceso.Start()


    # ======================================================
    # AUTOMATIZACION
    # ======================================================

    Write-Host ""
    Write-Host "========================================"
    Write-Host "INICIANDO PROCESAMIENTO"
    Write-Host "========================================"

    # ProcesoUsuariosV4 siempre se ejecuta en un PowerShell 5.1 nuevo.
    # Esto impide reutilizar ensamblados o credenciales de Graph que hayan
    # quedado cargados por una ejecucion anterior en la consola del operador.
    $argumentosProceso = @(
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        $Proceso
    )

    if ($Todas) {
        $argumentosProceso += "-Todas"
    }
    else {
        $argumentosProceso += @("-Fila", [string]$Fila)
    }

    & $PowerShellLimpio @argumentosProceso

    if ($LASTEXITCODE -ne 0) {
        throw "ProcesoUsuariosV4 finalizo con codigo $LASTEXITCODE."
    }

    if (Test-Path -LiteralPath $ResultadoProceso) {
        $MensajeAdvertencia = (
            Get-Content -LiteralPath $ResultadoProceso -Raw
        ).Trim()
    }

    if (-not [string]::IsNullOrWhiteSpace($MensajeAdvertencia)) {
        $huboAdvertencias = $true
    }


    # ======================================================
    # DETENER TIEMPO DE PROCESAMIENTO
    # ======================================================

    $CronometroProceso.Stop()

}
catch {

    $huboError = $true

    $MensajeError = $_.Exception.Message

    if ($CronometroProceso.IsRunning) {

        $CronometroProceso.Stop()

    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "ERROR DE AUTOMATIZACION" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red

    Write-Host $MensajeError -ForegroundColor Red

}
finally {

    # ======================================================
    # CERRAR CHROME Y ELIMINAR SESION TEMPORAL
    # ======================================================

    Write-Host ""
    Write-Host "========================================"
    Write-Host "CERRANDO CHROME"
    Write-Host "========================================"

    & py $CerrarPortal |
        Out-Host


    Write-Host ""
    Remove-Item Env:\AZURE_AUTOMATION_PROFILE `
        -ErrorAction SilentlyContinue

    Remove-Item Env:\AZURE_AUTOMATION_RESULT `
        -ErrorAction SilentlyContinue

    Write-Host ""
    Write-Host "Eliminando perfil temporal:"
    Write-Host $PerfilTemporal

    $perfilEliminado = $false

    for ($intento = 1; $intento -le 5; $intento++) {

        try {

            if (Test-Path -LiteralPath $PerfilTemporal) {
                Remove-Item `
                    -LiteralPath $PerfilTemporal `
                    -Recurse `
                    -Force `
                    -ErrorAction Stop
            }

            $perfilEliminado = $true
            break

        }
        catch {

            Start-Sleep -Milliseconds 500

        }

    }

    if ($perfilEliminado) {
        Write-Host "Perfil temporal eliminado." -ForegroundColor Green
    }
    else {
        Write-Host "ADVERTENCIA: no se pudo eliminar completamente el perfil temporal." -ForegroundColor Yellow
    }


    # ======================================================
    # DETENER CRONOMETRO TOTAL
    # ======================================================

    $CronometroTotal.Stop()

}


# ==========================================================
# CALCULAR TIEMPOS
# ==========================================================

$FechaFin = Get-Date

$TiempoTotal = $CronometroTotal.Elapsed

$TiempoProceso = $CronometroProceso.Elapsed


$TextoTiempoTotal = "{0:00}:{1:00}:{2:00}.{3:000}" -f `
    [math]::Floor($TiempoTotal.TotalHours),
    $TiempoTotal.Minutes,
    $TiempoTotal.Seconds,
    $TiempoTotal.Milliseconds


$TextoTiempoProceso = "{0:00}:{1:00}:{2:00}.{3:000}" -f `
    [math]::Floor($TiempoProceso.TotalHours),
    $TiempoProceso.Minutes,
    $TiempoProceso.Seconds,
    $TiempoProceso.Milliseconds


# ==========================================================
# MOSTRAR RESULTADO
# ==========================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "RESUMEN DE TIEMPOS" -ForegroundColor Cyan
Write-Host "========================================"

Write-Host "Inicio:               $FechaInicio"
Write-Host "Fin:                  $FechaFin"

Write-Host ""
Write-Host "Tiempo procesamiento: $TextoTiempoProceso"
Write-Host "Tiempo TOTAL:         $TextoTiempoTotal"

Write-Host "========================================"


# ==========================================================
# GUARDAR LOG DE TIEMPO
# ==========================================================

$EstadoFinal = if ($huboError) {

    "ERROR"

}
elseif ($huboAdvertencias) {

    "COMPLETADO CON OBSERVACIONES"

}
else {

    "COMPLETADO"

}


$LogPath = Join-Path `
    $LogsPath `
    "TiempoEjecucion_$RunId.txt"


$contenidoLog = @"
========================================
AUTOMATIZACION AZURE
========================================

Ejecucion:
$RunId

Estado:
$EstadoFinal

Inicio:
$FechaInicio

Fin:
$FechaFin

Tiempo de procesamiento:
$TextoTiempoProceso

Tiempo total:
$TextoTiempoTotal

Modo:
$(if ($Todas) { "Todas las filas" } else { "Fila $Fila" })

Error:
$MensajeError

Observaciones:
$MensajeAdvertencia

========================================
"@


Set-Content `
    -LiteralPath $LogPath `
    -Value $contenidoLog `
    -Encoding UTF8


Write-Host ""
Write-Host "Log guardado:"
Write-Host $LogPath


# ==========================================================
# RESULTADO FINAL
# ==========================================================

if ($huboError) {

    Write-Host ""
    Write-Host "PROCESO FINALIZADO CON ERROR." -ForegroundColor Red

    exit 1

}


Write-Host ""

if ($huboAdvertencias) {
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "AUTOMATIZACION COMPLETADA CON OBSERVACIONES" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host $MensajeAdvertencia -ForegroundColor Yellow
}
else {
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "AUTOMATIZACION COMPLETADA" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
}

Write-Host "Chrome cerrado."
Write-Host "La sesion temporal de Chrome no se conserva."

Write-Host ""
Write-Host "TIEMPO TOTAL: $TextoTiempoTotal" -ForegroundColor Green

exit 0



