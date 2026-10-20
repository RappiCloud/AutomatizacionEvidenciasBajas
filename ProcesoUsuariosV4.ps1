param(
    [int]$Fila = 0,
    [switch]$Todas
)

# ==========================================================
# AUTOMATIZACION AZURE V4
# PROCESO COMPLETAMENTE AUTOMATICO
#
# Excel:
# B = Nombre apellido
# C = UPN
# D = Antes
# E = Despues
# F = Accion tomada
#
# Desde fila 3
# ==========================================================

$ErrorActionPreference = "Stop"

# ==========================================================
# CONFIGURACION
# ==========================================================

$Base = "C:\AutomatizacionAzure"

$ExcelPath = Join-Path $Base "Evidencias.xlsx"

$CapturaPy = Join-Path $Base "CapturaAzureRapido.py"

$TenantPath = Join-Path $Base "TenantLaboratorio.txt"

$EjecucionesPath = Join-Path $Base "Ejecuciones"

$BackupsPath = Join-Path $Base "Backups"

$HojaNombre = "EVIDENCIAS"

$EliminacionIniciada = $false

# Una sola instancia de Excel durante el lote. Abrir y cerrar el libro para
# cada lectura, imagen y estado hacía que el tiempo creciera a medida que el
# archivo acumulaba capturas.
$script:ExcelApp = $null
$script:ExcelLibro = $null
$script:ExcelHoja = $null
$script:MetricasCaptura = New-Object 'System.Collections.Generic.List[object]'
$script:CapturadorProceso = $null


# ==========================================================
# VALIDAR PARAMETROS
# ==========================================================

if ($Todas -and $Fila -gt 0) {
    throw "Usa -Fila o -Todas, no ambos."
}

if (-not $Todas -and $Fila -lt 3) {
    throw "Debes indicar -Fila N o utilizar -Todas."
}


# ==========================================================
# FUNCIONES GENERALES
# ==========================================================

function Limpiar-NombreArchivo {

    param(
        [string]$Nombre
    )

    $resultado = $Nombre -replace '[\\/:*?"<>|]', '_'

    $resultado = $resultado -replace '\s+', '_'

    $resultado = $resultado.Trim(' ','.','_')

    if ([string]::IsNullOrWhiteSpace($resultado)) {
        $resultado = "Usuario"
    }

    if ($resultado.Length -gt 80) {
        $resultado = $resultado.Substring(0,80)
    }

    return $resultado
}


function Normalizar-TextoComparacion {

    param(
        [string]$Texto
    )

    if ($null -eq $Texto) {
        return ""
    }

    $descompuesto = $Texto.Normalize(
        [Text.NormalizationForm]::FormD
    )

    $sinTildes = New-Object Text.StringBuilder

    foreach ($caracter in $descompuesto.ToCharArray()) {

        $categoria = [Globalization.CharUnicodeInfo]::GetUnicodeCategory(
            $caracter
        )

        if ($categoria -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sinTildes.Append($caracter)
        }

    }

    return (
        ($sinTildes.ToString() -replace '\s+', ' ').Trim()
    ).ToLowerInvariant()
}


function Obtener-ExcepcionesAnidadas {

    param(
        [System.Exception]$Excepcion
    )

    $resultado = New-Object `
        'System.Collections.Generic.List[System.Exception]'

    $pendientes = New-Object System.Collections.Stack

    if ($null -ne $Excepcion) {
        $pendientes.Push($Excepcion)
    }

    while ($pendientes.Count -gt 0) {

        $actual = [System.Exception]$pendientes.Pop()

        if ($resultado.Contains($actual)) {
            continue
        }

        $resultado.Add($actual)

        if ($actual -is [System.AggregateException]) {

            foreach ($interna in $actual.Flatten().InnerExceptions) {

                if ($null -ne $interna) {
                    $pendientes.Push($interna)
                }

            }

        }

        if ($null -ne $actual.InnerException) {
            $pendientes.Push($actual.InnerException)
        }

    }

    return $resultado.ToArray()
}


function Obtener-DetalleError {

    param(
        $ErrorRecord
    )

    $mensajes = New-Object `
        'System.Collections.Generic.List[string]'

    try {

        $detalleGraph = [string]$ErrorRecord.ErrorDetails.Message

        if (-not [string]::IsNullOrWhiteSpace($detalleGraph)) {
            $mensajes.Add($detalleGraph.Trim())
        }

    }
    catch {}

    foreach (
        $excepcion in (
            Obtener-ExcepcionesAnidadas `
                -Excepcion $ErrorRecord.Exception
        )
    ) {

        $mensaje = [string]$excepcion.Message

        if (
            -not [string]::IsNullOrWhiteSpace($mensaje) -and
            -not $mensajes.Contains($mensaje.Trim())
        ) {
            $mensajes.Add($mensaje.Trim())
        }

    }

    if ($mensajes.Count -eq 0) {
        return [string]$ErrorRecord
    }

    return ($mensajes -join " | ")
}


function Es-Error404 {

    param(
        $ErrorRecord
    )

    $textos = New-Object `
        'System.Collections.Generic.List[string]'

    try {

        if ($null -ne $ErrorRecord.ErrorDetails.Message) {
            $textos.Add([string]$ErrorRecord.ErrorDetails.Message)
        }

    }
    catch {}

    try {

        if ($null -ne $ErrorRecord.FullyQualifiedErrorId) {
            $textos.Add([string]$ErrorRecord.FullyQualifiedErrorId)
        }

    }
    catch {}

    foreach (
        $excepcion in (
            Obtener-ExcepcionesAnidadas `
                -Excepcion $ErrorRecord.Exception
        )
    ) {

        foreach ($propiedad in @(
            "ResponseStatusCode",
            "StatusCode"
        )) {

            try {

                $estado = $excepcion.$propiedad

                if ($null -ne $estado) {
                    $textos.Add([string]$estado)
                }

            }
            catch {}

        }

        if (-not [string]::IsNullOrWhiteSpace($excepcion.Message)) {
            $textos.Add([string]$excepcion.Message)
        }

    }

    $detalle = $textos -join " | "

    if (
        $detalle -match '(^|[^0-9])404([^0-9]|$)' -or
        $detalle -match 'Request_ResourceNotFound' -or
        $detalle -match 'ResourceNotFound' -or
        $detalle -match 'NotFound' -or
        $detalle -match 'Resource.*not found' -or
        $detalle -match 'does not exist'
    ) {
        return $true
    }

    return $false
}


# ==========================================================
# MICROSOFT GRAPH
# ==========================================================

function Conectar-Graph {

    # Hay instalaciones separadas para Windows PowerShell 5.1
    # y PowerShell 7. Fijar una sola version evita que PowerShell
    # combine Users 2.38.1 con Authentication 2.40.0.

    $versionGraph = if (
        $PSVersionTable.PSEdition -eq "Desktop"
    ) {
        "2.38.1"
    }
    else {
        "2.39.0"
    }

    # Todas las operaciones se realizan con Invoke-MgGraphRequest para que
    # compartan el mismo token. Cargar los modulos Users y
    # Identity.DirectoryManagement creaba proveedores de autenticacion
    # adicionales que podian volver a mostrar el inicio de sesion por fila.
    $modulosGraph = @(
        "Microsoft.Graph.Authentication"
    )

    foreach ($moduloGraph in $modulosGraph) {

        $disponible = Get-Module `
            -ListAvailable `
            -Name $moduloGraph |
            Where-Object {
                $_.Version -eq [version]$versionGraph
            } |
            Select-Object -First 1

        if ($null -eq $disponible) {

            throw (
                "No esta instalada la version $versionGraph " +
                "del modulo $moduloGraph."
            )

        }

        Import-Module `
            -Name $disponible.Path `
            -Force `
            -ErrorAction Stop

    }

    $versionesCargadas = @(
        Get-Module $modulosGraph |
        Select-Object `
            Name,
            Version
    )

    foreach ($moduloCargado in $versionesCargadas) {

        if ($moduloCargado.Version -ne [version]$versionGraph) {

            throw (
                "Microsoft Graph tiene versiones mezcladas en memoria. " +
                "Cierra PowerShell y vuelve a ejecutar. " +
                "Esperada: $versionGraph. " +
                "Cargada: $($moduloCargado.Name) " +
                "$($moduloCargado.Version)."
            )

        }

    }

    if (-not (Test-Path -LiteralPath $TenantPath)) {
        throw "No existe TenantLaboratorio.txt."
    }

    $tenantPermitido = (
        Get-Content -LiteralPath $TenantPath -Raw
    ).Trim()

    if ($tenantPermitido -notmatch '^[0-9a-fA-F-]{36}$') {
        throw "TenantLaboratorio.txt no contiene un Tenant ID valido."
    }

    $scopesNecesarios = @(
        "User.Read.All",
        "User.DeleteRestore.All"
    )

    $contexto = Get-MgContext

    $reconectar = $false

    if ($null -eq $contexto) {

        $reconectar = $true

    }
    else {

        if ($contexto.TenantId -ine $tenantPermitido) {
            $reconectar = $true
        }

        if ($contexto.ContextScope -ne "CurrentUser") {
            $reconectar = $true
        }

        foreach ($scope in $scopesNecesarios) {

            if ($contexto.Scopes -notcontains $scope) {
                $reconectar = $true
            }

        }

    }

    if ($reconectar) {

        if ($null -ne $contexto) {

            Disconnect-MgGraph `
                -ErrorAction SilentlyContinue |
                Out-Null

        }

        Connect-MgGraph `
            -TenantId $tenantPermitido `
            -Scopes $scopesNecesarios `
            -UseDeviceCode `
            -ContextScope CurrentUser `
            -NoWelcome `
            -ErrorAction Stop

    }

    $contexto = Get-MgContext

    if ($null -eq $contexto) {
        throw "No hay conexion con Microsoft Graph."
    }

    if ($contexto.AuthType -ne "Delegated") {
        throw "La sesion Graph no es Delegated."
    }

    foreach ($scope in $scopesNecesarios) {

        if ($contexto.Scopes -notcontains $scope) {
            throw "Falta el permiso $scope."
        }

    }

    if ($contexto.TenantId -ne $tenantPermitido) {

        throw @"
TENANT INCORRECTO.

Actual:
$($contexto.TenantId)

Permitido:
$tenantPermitido

Proceso detenido.
"@

    }

    # Get-MgContext solamente confirma que existe un contexto. Se validan aqui
    # las mismas operaciones de lectura que se utilizaran durante las filas.
    # Si Microsoft necesita renovar o ampliar el token, debe hacerlo antes de
    # abrir Excel y nunca en medio del procesamiento de un usuario.
    try {

        $perfilGraph = Invoke-MgGraphRequest `
            -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/me?`$select=id,userPrincipalName" `
            -ErrorAction Stop

        $null = Invoke-MgGraphRequest `
            -Method GET `
            -Uri (
                "https://graph.microsoft.com/v1.0/users" +
                "?`$select=id&`$top=1"
            ) `
            -ErrorAction Stop

        $null = Invoke-MgGraphRequest `
            -Method GET `
            -Uri (
                "https://graph.microsoft.com/v1.0/directory/" +
                "deletedItems/microsoft.graph.user" +
                "?`$select=id&`$top=1"
            ) `
            -ErrorAction Stop

    }
    catch {

        throw (
            "No se pudo validar la sesion persistente de Microsoft Graph. " +
            "Completa una sola vez el codigo de dispositivo y vuelve a " +
            "ejecutar. Detalle: " +
            (Obtener-DetalleError -ErrorRecord $_)
        )

    }

    if (
        $null -eq $perfilGraph -or
        [string]::IsNullOrWhiteSpace([string]$perfilGraph.id)
    ) {
        throw "Microsoft Graph no devolvio la identidad autenticada."
    }

    Write-Host ""
    Write-Host "GRAPH CONECTADO" -ForegroundColor Green
    Write-Host "Cuenta: $($contexto.Account)"
    Write-Host "Tenant: $($contexto.TenantId)"

    return $contexto
}


function Obtener-UsuarioActivo {

    param(
        [string]$Identificador
    )

    try {

        $idUri = [Uri]::EscapeDataString($Identificador)

        return Invoke-MgGraphRequest `
            -Method GET `
            -Uri (
                "https://graph.microsoft.com/v1.0/users/${idUri}" +
                "?`$select=id,displayName,userPrincipalName,userType"
            ) `
            -ErrorAction Stop

    }
    catch {

        if (Es-Error404 $_) {
            return $null
        }

        throw

    }
}


function Buscar-UsuarioEliminadoPorNombre {

    param(
        [string]$Nombre
    )

    $eliminados = New-Object 'System.Collections.Generic.List[object]'

    $uri = (
        "https://graph.microsoft.com/v1.0/directory/deletedItems/" +
        "microsoft.graph.user" +
        "?`$select=id,displayName,userPrincipalName,deletedDateTime" +
        "&`$top=999"
    )

    while (-not [string]::IsNullOrWhiteSpace($uri)) {

        $respuesta = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $uri `
            -ErrorAction Stop

        foreach ($usuarioEliminado in @($respuesta.value)) {
            $eliminados.Add($usuarioEliminado)
        }

        $uri = [string]$respuesta.'@odata.nextLink'

    }

    # Primero coincidencia exacta

    $coincidencias = @(
        $eliminados |
        Where-Object {
            $_.DisplayName -and
            (Normalizar-TextoComparacion $_.DisplayName) -eq
            (Normalizar-TextoComparacion $Nombre)
        }
    )

    # Si no existe exacta, permitir que el nombre del Excel
    # este contenido dentro del nombre de Azure

    if ($coincidencias.Count -eq 0) {

        $coincidencias = @(
            $eliminados |
            Where-Object {

                $_.DisplayName -and

                (Normalizar-TextoComparacion $_.DisplayName).IndexOf(
                    (Normalizar-TextoComparacion $Nombre),
                    [StringComparison]::OrdinalIgnoreCase
                ) -ge 0

            }
        )

    }

    if ($coincidencias.Count -gt 0) {

        Write-Host (
            "Deleted users coincidentes: " +
            $coincidencias.Count
        ) -ForegroundColor Yellow

        # En este punto solamente se determina que el usuario esta
        # inactivo. No se restaura ni se elimina ningun objeto de la papelera,
        # por lo que varias coincidencias no impiden continuar.
        return [PSCustomObject]@{
            Cantidad = $coincidencias.Count
            Usuarios = $coincidencias
        }

    }

    return $null
}


# ==========================================================
# EXCEL - SESION UNICA DEL LOTE
# ==========================================================

function Iniciar-SesionExcel {

    if ($null -ne $script:ExcelLibro) {
        return
    }

    try {

        $script:ExcelApp = New-Object -ComObject Excel.Application
        $script:ExcelApp.Visible = $false
        $script:ExcelApp.DisplayAlerts = $false

        $script:ExcelLibro = $script:ExcelApp.Workbooks.Open($ExcelPath)

        if ($script:ExcelLibro.ReadOnly) {
            throw "El Excel esta abierto o bloqueado."
        }

        $script:ExcelHoja = $script:ExcelLibro.Worksheets.Item($HojaNombre)

        Write-Host "Sesion unica de Excel iniciada." -ForegroundColor Green

    }
    catch {

        Cerrar-SesionExcel
        throw

    }
}


function Guardar-SesionExcel {

    if ($null -eq $script:ExcelLibro) {
        throw "La sesion de Excel no esta iniciada."
    }

    $script:ExcelLibro.Save()
}


function Cerrar-SesionExcel {

    if ($null -ne $script:ExcelHoja) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject(
            $script:ExcelHoja
        )
        $script:ExcelHoja = $null
    }

    if ($null -ne $script:ExcelLibro) {
        try {
            $script:ExcelLibro.Close($false)
        }
        finally {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject(
                $script:ExcelLibro
            )
            $script:ExcelLibro = $null
        }
    }

    if ($null -ne $script:ExcelApp) {
        try {
            $script:ExcelApp.Quit()
        }
        finally {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject(
                $script:ExcelApp
            )
            $script:ExcelApp = $null
        }
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}


# ==========================================================
# EXCEL - LEER
# ==========================================================

function Leer-FilaExcel {

    param(
        [int]$NumeroFila
    )

    Iniciar-SesionExcel

    $nombre = [string]$script:ExcelHoja.Cells.Item(
        $NumeroFila,
        2
    ).Value2

    $upn = [string]$script:ExcelHoja.Cells.Item(
        $NumeroFila,
        3
    ).Value2

    $accion = [string]$script:ExcelHoja.Cells.Item(
        $NumeroFila,
        6
    ).Value2

    return [PSCustomObject]@{

        Nombre = $nombre.Trim()

        UPN = $upn.Trim()

        Accion = $accion.Trim()

    }
}


# ==========================================================
# EXCEL - ACCION
# ==========================================================

function Escribir-AccionExcel {

    param(
        [int]$NumeroFila,
        [string]$Accion
    )

    Iniciar-SesionExcel

    $script:ExcelHoja.Cells.Item(
        $NumeroFila,
        6
    ).Value2 = $Accion

    Write-Host "F$NumeroFila = $Accion"
}


# ==========================================================
# EXCEL - INSERTAR IMAGEN
# ==========================================================

function Insertar-ImagenExcel {

    param(
        [int]$NumeroFila,
        [ValidateSet("D","E")]
        [string]$Columna,
        [string]$ImagenPath
    )

    if (-not (Test-Path -LiteralPath $ImagenPath)) {
        throw "No existe la imagen: $ImagenPath"
    }

    Add-Type -AssemblyName System.Drawing

    $imagen = [System.Drawing.Image]::FromFile(
        $ImagenPath
    )

    try {

        [double]$anchoOriginal = $imagen.Width

        [double]$altoOriginal = $imagen.Height

    }
    finally {

        $imagen.Dispose()

    }

    Iniciar-SesionExcel

    $celda = $null
    $shape = $null

    try {

        $celdaTexto = "$Columna$NumeroFila"

        $direccionEsperada = '$' + $Columna + '$' + $NumeroFila

        $celda = $script:ExcelHoja.Range($celdaTexto)

        [double]$anchoCelda = $celda.Width

        [double]$altoCelda = $celda.Height

        [double]$anchoDisponible = $anchoCelda - 6

        [double]$altoDisponible = $altoCelda - 6

        if (
            $anchoOriginal -le 0 -or
            $altoOriginal -le 0 -or
            $anchoDisponible -le 0 -or
            $altoDisponible -le 0
        ) {

            throw "Dimensiones invalidas para $celdaTexto."

        }

        [double]$escala = [Math]::Min(
            1.0,
            [Math]::Min(
                $anchoDisponible / $anchoOriginal,
                $altoDisponible / $altoOriginal
            )
        )

        [double]$anchoFinal = $anchoOriginal * $escala

        [double]$altoFinal = $altoOriginal * $escala

        [double]$posX = $celda.Left +
            (($anchoCelda - $anchoFinal) / 2)

        [double]$posY = $celda.Top +
            (($altoCelda - $altoFinal) / 2)

        $shape = $script:ExcelHoja.Shapes.AddPicture(
            $ImagenPath,
            0,
            -1,
            [single]$posX,
            [single]$posY,
            [single]$anchoFinal,
            [single]$altoFinal
        )

        if (
            $shape.Width -le 0 -or
            $shape.Height -le 0
        ) {

            $shape.Delete()

            throw "Excel creo una imagen de tamano cero."

        }

        if ($Columna -eq "D") {

            $nombreShape = "Evidencia_Antes_Fila$NumeroFila"

        }
        else {

            $nombreShape = "Evidencia_Despues_Fila$NumeroFila"

        }

        # Borrar evidencia anterior de esta misma celda

        # Shapes.Item puede devolver el mismo RCW que ya esta guardado en
        # $shape para la imagen recien creada. FinalReleaseComObject sobre ese
        # alias invalida tambien $shape y rompe la sesion antes de continuar
        # con Deleted users. Conservamos el RCW de la imagen actual y solo
        # liberamos las demas figuras recorridas.

        $shapeIdActual = $shape.ID

        for ($i = $script:ExcelHoja.Shapes.Count; $i -ge 1; $i--) {

            $anterior = $script:ExcelHoja.Shapes.Item($i)

            if ($anterior.ID -eq $shapeIdActual) {
                $anterior = $null
                continue
            }

            if (
                $anterior.Name -eq $nombreShape
            ) {

                $anterior.Delete()

            }

            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject(
                $anterior
            )

        }

        $shape.Name = $nombreShape

        # Place over cells
        $shape.Placement = 1

        $shape.Visible = -1

        # Verificar la imagen en la sesion activa. La persistencia en disco se
        # confirma al guardar la fila completa.
        if (
            $shape.Name -ne $nombreShape -or
            $shape.Width -le 0 -or
            $shape.Height -le 0 -or
            $shape.TopLeftCell.Address() -ne $direccionEsperada
        ) {

            throw "No se pudo verificar la evidencia en $celdaTexto."

        }

        Write-Host "Evidencia verificada: $celdaTexto"

    }
    finally {

        if ($null -ne $shape) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject(
                $shape
            )
        }

        if ($null -ne $celda) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject(
                $celda
            )
        }

    }
}


# ==========================================================
# CAPTURADOR PYTHON PERSISTENTE
# ==========================================================

function Leer-LineaCapturador {

    param(
        [int]$TimeoutMs = 60000
    )

    if (
        $null -eq $script:CapturadorProceso -or
        $script:CapturadorProceso.HasExited
    ) {
        throw "El capturador persistente no esta disponible."
    }

    $tarea = $script:CapturadorProceso.StandardOutput.ReadLineAsync()

    if (-not $tarea.Wait($TimeoutMs)) {
        throw "El capturador persistente excedio el tiempo de respuesta."
    }

    $linea = $tarea.Result

    if ($null -eq $linea) {
        throw "El capturador persistente cerro su salida inesperadamente."
    }

    return $linea
}


function Iniciar-CapturadorPersistente {

    if (
        $null -ne $script:CapturadorProceso -and
        -not $script:CapturadorProceso.HasExited
    ) {
        return
    }

    $python = (Get-Command py -ErrorAction Stop).Source

    $inicioProceso = New-Object Diagnostics.ProcessStartInfo
    $inicioProceso.FileName = $python
    $inicioProceso.Arguments = (
        '-X utf8 -u "' + $CapturaPy.Replace('"','\"') + '" --servidor'
    )
    $inicioProceso.UseShellExecute = $false
    $inicioProceso.RedirectStandardInput = $true
    $inicioProceso.RedirectStandardOutput = $true
    $inicioProceso.RedirectStandardError = $false
    $inicioProceso.CreateNoWindow = $true
    $script:CapturadorProceso = New-Object Diagnostics.Process
    $script:CapturadorProceso.StartInfo = $inicioProceso

    if (-not $script:CapturadorProceso.Start()) {
        throw "No se pudo iniciar el capturador persistente."
    }

    while ($true) {

        $linea = Leer-LineaCapturador -TimeoutMs 30000

        if ($linea -eq "__CAPTURA_READY__") {
            break
        }

        Write-Host $linea
    }

    Write-Host "Capturador Playwright persistente iniciado." -ForegroundColor Green
}


function Invocar-CapturadorPersistente {

    param(
        [string]$Modo,
        [string]$Busqueda,
        [string]$Nombre,
        [string]$Esperado,
        [string]$Salida,
        [string]$ResultadoJson
    )

    Iniciar-CapturadorPersistente

    $solicitud = [ordered]@{
        modo = $Modo
        busqueda = $Busqueda
        nombre = $Nombre
        esperado = $Esperado
        salida = $Salida
        resultado_json = $ResultadoJson
    } | ConvertTo-Json -Compress

    # Windows PowerShell 5.1 no permite configurar StandardInputEncoding
    # en ProcessStartInfo. Se envia el JSON como Base64 ASCII para conservar
    # correctamente nombres y rutas con tildes u otros caracteres Unicode.
    $solicitudBytes = [Text.Encoding]::UTF8.GetBytes($solicitud)
    $solicitudBase64 = [Convert]::ToBase64String($solicitudBytes)

    $script:CapturadorProceso.StandardInput.WriteLine(
        "BASE64:$solicitudBase64"
    )
    $script:CapturadorProceso.StandardInput.Flush()

    while ($true) {

        $linea = Leer-LineaCapturador -TimeoutMs 120000

        if ($linea.StartsWith("__CAPTURA_RESULT__")) {

            $jsonRespuesta = $linea.Substring(
                "__CAPTURA_RESULT__".Length
            )

            $respuesta = $jsonRespuesta | ConvertFrom-Json

            if (-not $respuesta.ok) {
                throw [string]$respuesta.error
            }

            return $respuesta
        }

        Write-Host $linea
    }
}


function Cerrar-CapturadorPersistente {

    if ($null -eq $script:CapturadorProceso) {
        return
    }

    try {

        if (-not $script:CapturadorProceso.HasExited) {

            $cierre = @{ comando = "cerrar" } | ConvertTo-Json -Compress
            $cierreBytes = [Text.Encoding]::UTF8.GetBytes($cierre)
            $cierreBase64 = [Convert]::ToBase64String($cierreBytes)

            $script:CapturadorProceso.StandardInput.WriteLine(
                "BASE64:$cierreBase64"
            )
            $script:CapturadorProceso.StandardInput.Flush()

            if (-not $script:CapturadorProceso.WaitForExit(5000)) {
                $script:CapturadorProceso.Kill()
                [void]$script:CapturadorProceso.WaitForExit(5000)
            }
        }

    }
    finally {

        $script:CapturadorProceso.Dispose()
        $script:CapturadorProceso = $null

    }
}


# ==========================================================
# CAPTURA AZURE V4
# ==========================================================

function Capturar-Azure {

    param(
        [ValidateSet("all","deleted")]
        [string]$Modo,

        [string]$Busqueda,

        [string]$Nombre,

        [ValidateSet("encontrado","no_encontrado","auto")]
        [string]$Esperado,

        [string]$Salida,

        [string]$ResultadoJson,

        [int]$Reintentos = 3
    )

    $cronometroCaptura = [Diagnostics.Stopwatch]::StartNew()

    for ($intento = 1; $intento -le $Reintentos; $intento++) {

        Write-Host ""
        Write-Host "Captura Azure intento $intento / $Reintentos"
        Write-Host "Modo: $Modo"
        Write-Host "Busqueda: $Busqueda"
        Write-Host "Esperado: $Esperado"

        if (Test-Path -LiteralPath $Salida) {
            Remove-Item -LiteralPath $Salida -Force
        }

        if (
            -not [string]::IsNullOrWhiteSpace($ResultadoJson) -and
            (Test-Path -LiteralPath $ResultadoJson)
        ) {
            Remove-Item -LiteralPath $ResultadoJson -Force
        }

        try {

            [void](
                Invocar-CapturadorPersistente `
                    -Modo $Modo `
                    -Busqueda $Busqueda `
                    -Nombre $Nombre `
                    -Esperado $Esperado `
                    -Salida $Salida `
                    -ResultadoJson $ResultadoJson
            )

            $codigo = 0
            $salidaCaptura = @()

        }
        catch {

            $codigo = 1
            $salidaCaptura = @(
                Obtener-DetalleError -ErrorRecord $_
            )

            # Una falla puede dejar la conexion CDP en un estado incompleto.
            # El siguiente intento iniciara un trabajador limpio.
            Cerrar-CapturadorPersistente

            $salidaCaptura |
                ForEach-Object {
                    Write-Host $_ -ForegroundColor Red
                }

        }

        if (
            $codigo -eq 0 -and
            (Test-Path -LiteralPath $Salida)
        ) {

            $archivo = Get-Item -LiteralPath $Salida

            if ($archivo.Length -gt 10000) {

                Write-Host "Captura valida: $Salida" -ForegroundColor Green

                if (-not [string]::IsNullOrWhiteSpace($ResultadoJson)) {

                    if (-not (Test-Path -LiteralPath $ResultadoJson)) {
                        throw "No se genero el resultado estructurado del portal."
                    }

                    $resultadoEstructurado = (
                        Get-Content `
                            -LiteralPath $ResultadoJson `
                            -Encoding UTF8 `
                            -Raw |
                        ConvertFrom-Json
                    )

                    $cronometroCaptura.Stop()
                    $script:MetricasCaptura.Add(
                        [pscustomobject]@{
                            Fecha = Get-Date
                            Modo = $Modo
                            Esperado = $Esperado
                            Intentos = $intento
                            Segundos = [math]::Round(
                                $cronometroCaptura.Elapsed.TotalSeconds,
                                3
                            )
                            Estado = "OK"
                            Salida = $Salida
                        }
                    )

                    return $resultadoEstructurado

                }

                $cronometroCaptura.Stop()
                $script:MetricasCaptura.Add(
                    [pscustomobject]@{
                        Fecha = Get-Date
                        Modo = $Modo
                        Esperado = $Esperado
                        Intentos = $intento
                        Segundos = [math]::Round(
                            $cronometroCaptura.Elapsed.TotalSeconds,
                            3
                        )
                        Estado = "OK"
                        Salida = $Salida
                    }
                )

                return

            }

        }

        if ($intento -lt $Reintentos) {

            Write-Host "Esperando antes de reintentar..."

            Start-Sleep -Seconds 3

        }

    }

    $detalleCaptura = (
        $salidaCaptura |
        ForEach-Object { "$_".Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Last 1
    )

    if ([string]::IsNullOrWhiteSpace($detalleCaptura)) {
        $detalleCaptura = "El capturador no devolvio detalles."
    }

    $cronometroCaptura.Stop()
    $script:MetricasCaptura.Add(
        [pscustomobject]@{
            Fecha = Get-Date
            Modo = $Modo
            Esperado = $Esperado
            Intentos = $Reintentos
            Segundos = [math]::Round(
                $cronometroCaptura.Elapsed.TotalSeconds,
                3
            )
            Estado = "ERROR"
            Salida = $Salida
        }
    )

    throw (
        "No se pudo generar la evidencia automatica. " +
        "Ultimo error del capturador: $detalleCaptura"
    )
}


# ==========================================================
# FILAS PENDIENTES
# ==========================================================

function Obtener-FilasPendientes {

    $excel = $null
    $libro = $null

    $resultado = @()

    try {

        $excel = New-Object -ComObject Excel.Application

        $excel.Visible = $false

        $libro = $excel.Workbooks.Open(
            $ExcelPath,
            0,
            $true
        )

        $hoja = $libro.Worksheets.Item($HojaNombre)

        $ultimaB = $hoja.Cells.Item(
            $hoja.Rows.Count,
            2
        ).End(-4162).Row

        $ultimaC = $hoja.Cells.Item(
            $hoja.Rows.Count,
            3
        ).End(-4162).Row

        $ultimaFila = [Math]::Max(
            $ultimaB,
            $ultimaC
        )

        for ($i = 3; $i -le $ultimaFila; $i++) {

            $nombre = [string]$hoja.Cells.Item($i,2).Value2

            $upn = [string]$hoja.Cells.Item($i,3).Value2

            $accion = [string]$hoja.Cells.Item($i,6).Value2

            if (
                [string]::IsNullOrWhiteSpace($nombre) -and
                [string]::IsNullOrWhiteSpace($upn)
            ) {
                continue
            }

            if (
                [string]::IsNullOrWhiteSpace($nombre) -or
                [string]::IsNullOrWhiteSpace($upn)
            ) {

                throw "Datos incompletos en fila $i."

            }

            if (-not [string]::IsNullOrWhiteSpace($accion)) {

                Write-Host "Fila $i omitida: $accion"

                continue

            }

            $resultado += $i

        }

    }
    finally {

        if ($null -ne $libro) {
            $libro.Close($false)
        }

        if ($null -ne $excel) {
            $excel.Quit()
        }

        [GC]::Collect()

        [GC]::WaitForPendingFinalizers()

    }

    return $resultado
}


# ==========================================================
# PROCESAR UNA FILA
# ==========================================================

function Procesar-Fila {

    param(
        [int]$NumeroFila,
        [string]$RunPath
    )

    $script:EliminacionIniciada = $false

    Write-Host ""
    Write-Host "============================================"
    Write-Host "PROCESANDO FILA $NumeroFila"
    Write-Host "============================================"

    $datos = Leer-FilaExcel -NumeroFila $NumeroFila

    if (
        [string]::IsNullOrWhiteSpace($datos.Nombre) -or
        [string]::IsNullOrWhiteSpace($datos.UPN)
    ) {

        throw "Nombre o UPN vacio en fila $NumeroFila."

    }

    if (-not [string]::IsNullOrWhiteSpace($datos.Accion)) {

        Write-Host "Fila ya procesada: $($datos.Accion)"

        return

    }

    Write-Host "Nombre: $($datos.Nombre)"
    Write-Host "UPN: $($datos.UPN)"

    $nombreSeguro = Limpiar-NombreArchivo $datos.Nombre

    $carpetaUsuario = Join-Path `
        $RunPath `
        "Fila_${NumeroFila}_${nombreSeguro}"

    New-Item `
        -ItemType Directory `
        -Path $carpetaUsuario `
        -Force |
        Out-Null

    $imagenAntes = Join-Path `
        $carpetaUsuario `
        "${nombreSeguro}_ANTES.png"

    $imagenDespues = Join-Path `
        $carpetaUsuario `
        "${nombreSeguro}_DESPUES.png"


    # ======================================================
    # BUSCAR LITERALMENTE EL VALOR DEL EXCEL EN ALL USERS
    # ======================================================

    $resultadoJson = Join-Path `
        $carpetaUsuario `
        "resultado_portal.json"

    $resultadoPortal = Capturar-Azure `
        -Modo "all" `
        -Busqueda $datos.UPN `
        -Nombre $datos.Nombre `
        -Esperado "auto" `
        -Salida $imagenAntes `
        -ResultadoJson $resultadoJson

    if ($null -eq $resultadoPortal) {
        throw "Azure no devolvio el resultado estructurado de All users."
    }

    # ======================================================
    # DETECTAR ESTADO ACTUAL
    # ======================================================

    $usuarioActivo = $null
    $usuarioEliminado = $null

    if ($resultadoPortal.encontrado -eq $true) {

        if (
            [string]::IsNullOrWhiteSpace($resultadoPortal.objectId) -or
            [string]::IsNullOrWhiteSpace($resultadoPortal.userPrincipalName) -or
            [string]::IsNullOrWhiteSpace($resultadoPortal.userType)
        ) {

            throw "La fila de Azure no contiene todos los identificadores requeridos."

        }

        $nombrePortalCoincide = (
            (Normalizar-TextoComparacion $resultadoPortal.displayName) -eq
            (Normalizar-TextoComparacion $datos.Nombre)
        )

        $upnPortalCoincide = (
            $resultadoPortal.userPrincipalName -ieq $datos.UPN
        )

        if (-not $nombrePortalCoincide -and -not $upnPortalCoincide) {

            throw "La fila visible no coincide con el nombre ni con el UPN del Excel."

        }

        if (-not $nombrePortalCoincide -and $upnPortalCoincide) {

            Write-Host (
                "ADVERTENCIA: el nombre visible difiere del Excel, " +
                "pero el UPN coincide exactamente."
            ) -ForegroundColor Yellow

        }

        $usuarioActivo = Obtener-UsuarioActivo `
            -Identificador $resultadoPortal.objectId

        if ($null -eq $usuarioActivo) {
            throw "Graph no encontro el Object ID mostrado por Azure."
        }

        if (
            $usuarioActivo.Id -ine $resultadoPortal.objectId -or
            $usuarioActivo.UserPrincipalName -ine $resultadoPortal.userPrincipalName -or
            $usuarioActivo.UserType -ine $resultadoPortal.userType -or
            (Normalizar-TextoComparacion $usuarioActivo.DisplayName) -ne
            (Normalizar-TextoComparacion $resultadoPortal.displayName)
        ) {

            throw "La identidad visible en Azure no coincide con Graph."

        }

        if ($usuarioActivo.UserType -eq "Guest") {

            $escenario = "Guest"

        }
        elseif ($usuarioActivo.UserType -eq "Member") {

            $escenario = "Member"

        }
        else {

            throw "UserType no reconocido: $($usuarioActivo.UserType)"

        }

    }
    else {

        $usuarioEliminado = Buscar-UsuarioEliminadoPorNombre `
            -Nombre $datos.Nombre

        if ($null -ne $usuarioEliminado) {

            $escenario = "Usuario Inactivo"

        }
        else {

            $escenario = "Usuario No Existe"

        }

    }

    Write-Host ""
    Write-Host "ESCENARIO: $escenario" -ForegroundColor Cyan

    Insertar-ImagenExcel `
        -NumeroFila $NumeroFila `
        -Columna "D" `
        -ImagenPath $imagenAntes


    # ======================================================
    # GUEST - ELIMINACION AUTOMATICA
    # ======================================================

    if ($escenario -eq "Guest") {

        Write-Host ""
        Write-Host "REVALIDANDO GUEST..."

        $actual = Obtener-UsuarioActivo `
            -Identificador $usuarioActivo.Id

        if ($null -eq $actual) {
            throw "El Guest dejo de existir antes de la eliminacion."
        }

        if (
            $actual.Id -ne $usuarioActivo.Id -or
            $actual.UserPrincipalName -ine $resultadoPortal.userPrincipalName -or
            (Normalizar-TextoComparacion $actual.DisplayName) -ne
            (Normalizar-TextoComparacion $resultadoPortal.displayName) -or
            $actual.UserType -ne "Guest"
        ) {

            throw "La identidad cambio. Eliminacion cancelada."

        }

        # Confirmar nuevamente tenant

        $contexto = Get-MgContext

        $tenantPermitido = (
            Get-Content -LiteralPath $TenantPath -Raw
        ).Trim()

        if ($contexto.TenantId -ne $tenantPermitido) {

            throw "Tenant no autorizado. Eliminacion cancelada."

        }

        # Marcar estado parcial ANTES de la accion destructiva

        Escribir-AccionExcel `
            -NumeroFila $NumeroFila `
            -Accion "Eliminacion en curso - verificar"

        # Esta marca debe quedar en disco antes de la accion destructiva.
        Guardar-SesionExcel

        $objectIdOriginal = $actual.Id

        Write-Host ""
        Write-Host "ELIMINANDO GUEST AUTOMATICAMENTE" -ForegroundColor Yellow
        Write-Host "Nombre: $($actual.DisplayName)"
        Write-Host "UPN: $($actual.UserPrincipalName)"
        Write-Host "Object ID: $objectIdOriginal"

        $script:EliminacionIniciada = $true

        # Eliminar el mismo objeto validado por Portal y Graph

        $objectIdUri = [Uri]::EscapeDataString($objectIdOriginal)

        Invoke-MgGraphRequest `
            -Method DELETE `
            -Uri "https://graph.microsoft.com/v1.0/users/$objectIdUri" `
            -ErrorAction Stop

        Write-Host "Solicitud de eliminacion enviada."

        # Confirmar el MISMO objeto por Object ID en Deleted users

        $eliminadoConfirmado = $null

        for ($intento = 1; $intento -le 30; $intento++) {

            try {

                $eliminadoConfirmado = Invoke-MgGraphRequest `
                    -Method GET `
                    -Uri (
                        "https://graph.microsoft.com/v1.0/" +
                        "directory/deletedItems/$objectIdUri" +
                        "?`$select=id,displayName," +
                        "userPrincipalName,deletedDateTime"
                    ) `
                    -ErrorAction Stop

                if (
                    $null -ne $eliminadoConfirmado -and
                    $eliminadoConfirmado.Id -eq $objectIdOriginal
                ) {

                    break

                }

            }
            catch {

                if (-not (Es-Error404 $_)) {
                    throw
                }

                $eliminadoConfirmado = $null

            }

            Write-Host "Esperando propagacion: $intento/30"

            Start-Sleep -Seconds 5

        }

        if (
            $null -eq $eliminadoConfirmado -or
            $eliminadoConfirmado.Id -ne $objectIdOriginal
        ) {

            throw "No se confirmo el Guest en Deleted users."

        }

        Write-Host "Eliminacion confirmada por Object ID." -ForegroundColor Green

        # Captura DESPUES automatica

        Capturar-Azure `
            -Modo "deleted" `
            -Busqueda $datos.Nombre `
            -Nombre $datos.Nombre `
            -Esperado "encontrado" `
            -Salida $imagenDespues `
            -Reintentos 4

        Insertar-ImagenExcel `
            -NumeroFila $NumeroFila `
            -Columna "E" `
            -ImagenPath $imagenDespues

        Escribir-AccionExcel `
            -NumeroFila $NumeroFila `
            -Accion "Eliminado"

        Guardar-SesionExcel

        Write-Host "GUEST COMPLETADO." -ForegroundColor Green

        return
    }


    # ======================================================
    # MEMBER
    # ======================================================

    if ($escenario -eq "Member") {

        $actual = Obtener-UsuarioActivo `
            -Identificador $usuarioActivo.Id

        if (
            $null -eq $actual -or
            $actual.Id -ne $usuarioActivo.Id -or
            $actual.UserType -ne "Member"
        ) {

            throw "El estado del Member cambio."

        }

        Capturar-Azure `
            -Modo "all" `
            -Busqueda $datos.UPN `
            -Nombre $datos.Nombre `
            -Esperado "encontrado" `
            -Salida $imagenDespues

        Insertar-ImagenExcel `
            -NumeroFila $NumeroFila `
            -Columna "E" `
            -ImagenPath $imagenDespues

        Escribir-AccionExcel `
            -NumeroFila $NumeroFila `
            -Accion "Usuario tipo Member"

        Guardar-SesionExcel

        Write-Host "MEMBER COMPLETADO." -ForegroundColor Green

        return
    }


    # ======================================================
    # USUARIO INACTIVO
    # ======================================================

    if ($escenario -eq "Usuario Inactivo") {

        Capturar-Azure `
            -Modo "deleted" `
            -Busqueda $datos.Nombre `
            -Nombre $datos.Nombre `
            -Esperado "encontrado" `
            -Salida $imagenDespues

        Insertar-ImagenExcel `
            -NumeroFila $NumeroFila `
            -Columna "E" `
            -ImagenPath $imagenDespues

        Escribir-AccionExcel `
            -NumeroFila $NumeroFila `
            -Accion "Usuario Inactivo"

        Guardar-SesionExcel

        Write-Host "USUARIO INACTIVO COMPLETADO." -ForegroundColor Green

        return
    }


    # ======================================================
    # USUARIO NO EXISTE
    # ======================================================

    if ($escenario -eq "Usuario No Existe") {

        Capturar-Azure `
            -Modo "deleted" `
            -Busqueda $datos.Nombre `
            -Nombre $datos.Nombre `
            -Esperado "no_encontrado" `
            -Salida $imagenDespues

        Insertar-ImagenExcel `
            -NumeroFila $NumeroFila `
            -Columna "E" `
            -ImagenPath $imagenDespues

        Escribir-AccionExcel `
            -NumeroFila $NumeroFila `
            -Accion "Usuario No Existe"

        Guardar-SesionExcel

        Write-Host "USUARIO NO EXISTE COMPLETADO." -ForegroundColor Green

        return
    }

}


# ==========================================================
# COMPROBACIONES INICIALES
# ==========================================================

foreach ($archivo in @(
    $ExcelPath,
    $CapturaPy,
    $TenantPath
)) {

    if (-not (Test-Path -LiteralPath $archivo -PathType Leaf)) {

        throw "No existe: $archivo"

    }

}

if (-not (Get-Command py -ErrorAction SilentlyContinue)) {
    throw "Python no esta disponible."
}

New-Item `
    -ItemType Directory `
    -Path $EjecucionesPath `
    -Force |
    Out-Null

New-Item `
    -ItemType Directory `
    -Path $BackupsPath `
    -Force |
    Out-Null


# ==========================================================
# CONECTAR GRAPH Y VALIDAR TENANT
# ==========================================================

$contexto = Conectar-Graph


# ==========================================================
# PREPARAR EJECUCION
# ==========================================================

$runId = (
    Get-Date -Format "yyyyMMdd_HHmmss"
) + "_" + (
    [guid]::NewGuid().ToString("N").Substring(0,8)
)

$runPath = Join-Path `
    $EjecucionesPath `
    $runId

New-Item `
    -ItemType Directory `
    -Path $runPath `
    -Force |
    Out-Null


# ==========================================================
# BACKUP GENERAL ANTES DEL PROCESO
# ==========================================================

$backupPath = Join-Path `
    $BackupsPath `
    "Evidencias_${runId}.xlsx"

Copy-Item `
    -LiteralPath $ExcelPath `
    -Destination $backupPath `
    -ErrorAction Stop

Write-Host ""
Write-Host "Backup:"
Write-Host $backupPath


# ==========================================================
# SELECCIONAR FILAS
# ==========================================================

if ($Todas) {

    $filas = @(
        Obtener-FilasPendientes
    )

}
else {

    $filas = @($Fila)

}

if ($filas.Count -eq 0) {

    Write-Host ""
    Write-Host "No existen filas pendientes."

    return
}

Write-Host ""
Write-Host "FILAS A PROCESAR:"
Write-Host ($filas -join ", ")


# ==========================================================
# PROCESAR
# ==========================================================

$incidenciasFilas = New-Object `
    'System.Collections.Generic.List[object]'

$metricasFilas = New-Object `
    'System.Collections.Generic.List[object]'

$env:AZURE_AUTOMATION_ROW_WARNINGS = ""

try {

Iniciar-CapturadorPersistente
Iniciar-SesionExcel

foreach ($numeroFila in $filas) {

    $cronometroFila = [Diagnostics.Stopwatch]::StartNew()
    $estadoFila = "OK"

    try {

        Procesar-Fila `
            -NumeroFila $numeroFila `
            -RunPath $runPath

    }
    catch {

        $estadoFila = "VERIFICAR"

        Write-Host ""
        Write-Host "========================================" -ForegroundColor Red
        Write-Host "ERROR EN FILA $numeroFila" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red

        $detalleError = Obtener-DetalleError `
            -ErrorRecord $_

        Write-Host $detalleError -ForegroundColor Red

        $detalleCorto = (
            $detalleError -replace '\s+', ' '
        ).Trim()

        if ($detalleCorto.Length -gt 160) {
            $detalleCorto = $detalleCorto.Substring(0,157) + "..."
        }

        if ($EliminacionIniciada) {
            $accionRevision = "Verificar eliminacion - $detalleCorto"
            Write-Host "La eliminacion pudo haberse ejecutado." -ForegroundColor Yellow
        }
        else {
            $accionRevision = "Verificar - $detalleCorto"
            Write-Host "No se inicio ninguna eliminacion en esta fila."
        }

        try {
            Escribir-AccionExcel `
                -NumeroFila $numeroFila `
                -Accion $accionRevision

            Guardar-SesionExcel
        }
        catch {
            throw (
                "No se pudo marcar la fila $numeroFila para verificar. " +
                (Obtener-DetalleError -ErrorRecord $_)
            )
        }

        $incidenciasFilas.Add(
            [pscustomobject]@{
                Fila = $numeroFila
                Accion = $accionRevision
                Error = $detalleError
            }
        )

        # Una falla de sesion, navegador, tenant o Excel afecta todo el lote.
        # En esos casos se conserva la marca de revision y se detiene por
        # seguridad; los errores propios de un usuario continuan normalmente.
        $falloGlobal = $detalleError -match (
            'SESSION_EXPIRADA|' +
            'ECONNREFUSED|' +
            'No existe contexto Chrome|' +
            'Tenant no autorizado|' +
            'Could not load file or assembly|' +
            'Excel esta abierto o bloqueado|' +
            'Excel esta bloqueado o abierto'
        )

        if ($falloGlobal) {
            throw "Fallo general de infraestructura: $detalleError"
        }

        Write-Host (
            "Fila $numeroFila marcada para verificar. " +
            "Se continuara con la siguiente fila."
        ) -ForegroundColor Yellow

        continue
    }
    finally {

        $cronometroFila.Stop()

        Write-Host (
            "Tiempo fila ${numeroFila}: " +
            $cronometroFila.Elapsed.ToString("hh\:mm\:ss\.fff")
        ) -ForegroundColor Cyan

        $metricasFilas.Add(
            [pscustomobject]@{
                Fila = $numeroFila
                Estado = $estadoFila
                Segundos = [math]::Round(
                    $cronometroFila.Elapsed.TotalSeconds,
                    3
                )
            }
        )

    }

}

}
finally {

    Cerrar-CapturadorPersistente
    Cerrar-SesionExcel

}

$metricasFilas |
    Export-Csv `
        -LiteralPath (Join-Path $runPath "TiemposPorFila.csv") `
        -NoTypeInformation `
        -Encoding UTF8

$script:MetricasCaptura |
    Export-Csv `
        -LiteralPath (Join-Path $runPath "TiemposCapturas.csv") `
        -NoTypeInformation `
        -Encoding UTF8

if ($incidenciasFilas.Count -gt 0) {

    $filasRevision = @(
        $incidenciasFilas |
        ForEach-Object { $_.Fila }
    )

    $resumenIncidencias = @(
        "EJECUCION CON FILAS PARA VERIFICAR"
        "Ejecucion: $runId"
        "Fecha: $(Get-Date)"
        ""
    )

    foreach ($incidencia in $incidenciasFilas) {
        $resumenIncidencias += "Fila $($incidencia.Fila)"
        $resumenIncidencias += "Accion: $($incidencia.Accion)"
        $resumenIncidencias += "Error: $($incidencia.Error)"
        $resumenIncidencias += ""
    }

    $archivoIncidencias = Join-Path `
        $runPath `
        "IncidenciasFilas.txt"

    Set-Content `
        -LiteralPath $archivoIncidencias `
        -Value $resumenIncidencias `
        -Encoding UTF8

    $env:AZURE_AUTOMATION_ROW_WARNINGS = (
        "Filas para verificar: " +
        ($filasRevision -join ", ") +
        ". Detalle: $archivoIncidencias"
    )

    if (
        -not [string]::IsNullOrWhiteSpace(
            [string]$env:AZURE_AUTOMATION_RESULT
        )
    ) {
        Set-Content `
            -LiteralPath $env:AZURE_AUTOMATION_RESULT `
            -Value $env:AZURE_AUTOMATION_ROW_WARNINGS `
            -Encoding UTF8
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "PROCESO COMPLETADO CON OBSERVACIONES" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host $env:AZURE_AUTOMATION_ROW_WARNINGS -ForegroundColor Yellow

}
else {

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "PROCESO V4 COMPLETADO" -ForegroundColor Green
    Write-Host "========================================"

}

Write-Host "Ejecucion:"
Write-Host $runPath
Write-Host ""
Write-Host "Excel:"
Write-Host $ExcelPath


