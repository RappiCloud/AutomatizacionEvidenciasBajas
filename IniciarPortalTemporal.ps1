$ErrorActionPreference = "Stop"

# ==========================================================
# CHROME - SESION TEMPORAL POR EJECUCION
# ==========================================================

$Puerto = 9223

$ArchivoTenant = Join-Path $PSScriptRoot "TenantLaboratorio.txt"

if (-not (Test-Path -LiteralPath $ArchivoTenant)) {
    throw "No se encontro TenantLaboratorio.txt."
}

$TenantPortal = (Get-Content -LiteralPath $ArchivoTenant -Raw).Trim()
$GuidTenant = [guid]::Empty

if (-not [guid]::TryParse($TenantPortal, [ref]$GuidTenant)) {
    throw "TenantLaboratorio.txt no contiene un TenantId valido."
}

# La ruta con TenantId evita que Azure use el ultimo directorio recordado.
$Portal = "https://portal.azure.com/$TenantPortal/"

# El wrapper debe crear esta variable
$Perfil = $env:AZURE_AUTOMATION_PROFILE

if ([string]::IsNullOrWhiteSpace($Perfil)) {
    throw "No se definio AZURE_AUTOMATION_PROFILE."
}

New-Item `
    -ItemType Directory `
    -Path $Perfil `
    -Force | Out-Null


# ==========================================================
# BUSCAR CHROME
# ==========================================================

$rutasChrome = @(
    "C:\Program Files\Google\Chrome\Application\chrome.exe",
    "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    (Join-Path $env:LOCALAPPDATA "Google\Chrome\Application\chrome.exe")
)

$Chrome = $null

foreach ($ruta in $rutasChrome) {

    if (Test-Path -LiteralPath $ruta) {
        $Chrome = $ruta
        break
    }

}

if ($null -eq $Chrome) {
    throw "No se encontro Google Chrome."
}


# ==========================================================
# COMPROBAR CDP
# ==========================================================

function Test-CDP {

    try {

        $null = Invoke-RestMethod `
            -Uri "http://127.0.0.1:$Puerto/json/version" `
            -TimeoutSec 1 `
            -ErrorAction Stop

        return $true

    }
    catch {
        return $false
    }

}


# ==========================================================
# NO DEBE QUEDAR UNA SESION ANTERIOR
# ==========================================================

if (Test-CDP) {

    throw @"
Ya existe un Chrome de automatizacion usando el puerto $Puerto.

Cierra la ejecucion anterior antes de iniciar una nueva.
"@

}


# ==========================================================
# INICIAR CHROME
# ==========================================================

Write-Host ""
Write-Host "Iniciando sesion temporal de Azure..." -ForegroundColor Cyan
Write-Host "Perfil temporal:"
Write-Host $Perfil

$argumentos = @(
    "--remote-debugging-port=$Puerto",
    "--remote-debugging-address=127.0.0.1",
    "--user-data-dir=$Perfil",
    "--start-maximized",
    "--no-first-run",
    "--no-default-browser-check",
    $Portal
)

Start-Process `
    -FilePath $Chrome `
    -ArgumentList $argumentos


# ==========================================================
# ESPERAR CDP
# ==========================================================

$listo = $false

for ($i = 1; $i -le 40; $i++) {

    Start-Sleep -Milliseconds 250

    if (Test-CDP) {

        $listo = $true
        break

    }

}

if (-not $listo) {
    throw "Chrome no habilito el puerto de automatizacion."
}

Write-Host ""
Write-Host "CHROME LISTO" -ForegroundColor Green
Write-Host "La sesion se eliminara al terminar la ejecucion."


# ==========================================================
# ESPERAR INICIO DE SESION EN EL PORTAL
# ==========================================================

Write-Host ""
Write-Host "Esperando autenticacion en Azure Portal..." -ForegroundColor Yellow
Write-Host "Completa el inicio de sesion en Chrome si Microsoft lo solicita."

# Evita aceptar el URL inicial antes de que Azure redirija al login.
Start-Sleep -Seconds 2

$portalAutenticado = $false
$confirmacionesPortal = 0

for ($i = 1; $i -le 600; $i++) {

    try {

        $paginas = @(
            Invoke-RestMethod `
                -Uri "http://127.0.0.1:$Puerto/json/list" `
                -TimeoutSec 2 `
                -ErrorAction Stop |
            Where-Object {
                $_.type -eq "page"
            }
        )

        $paginaPortal = $paginas |
            Where-Object {
                $_.url -like "https://portal.azure.com/*" -and
                $_.url -notlike "https://portal.azure.com/auth/login*"
            } |
            Select-Object -First 1

        if ($null -ne $paginaPortal) {

            $confirmacionesPortal++

            if ($confirmacionesPortal -ge 3) {
                $portalAutenticado = $true
                break
            }

        }
        else {
            $confirmacionesPortal = 0
        }

    }
    catch {
        $confirmacionesPortal = 0
    }

    Start-Sleep -Milliseconds 500

}

if (-not $portalAutenticado) {
    throw "No se completo el inicio de sesion en Azure Portal dentro de 5 minutos."
}

Write-Host "AZURE PORTAL AUTENTICADO" -ForegroundColor Green


