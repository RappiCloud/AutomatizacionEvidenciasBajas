# Automatización de usuarios en Microsoft Entra ID

Automatización para validar usuarios desde Excel, consultar Microsoft Graph,
navegar por Azure Portal con Chrome, generar evidencias visuales y eliminar
usuarios `Guest` después de validar su identidad y el tenant autorizado.

## Funciones principales

- Procesa una fila específica o todas las filas pendientes de `Evidencias.xlsx`.
- Busca el UPN exacto en **All users**.
- Distingue usuarios `Guest`, `Member`, inactivos y no existentes.
- Elimina automáticamente usuarios `Guest` después de validar Portal y Graph.
- Confirma la eliminación por Object ID en **Deleted users**.
- Inserta capturas de antes y después en Excel.
- Marca errores individuales como `Verificar` y continúa con las demás filas.
- Guarda ejecuciones, incidencias, copias de seguridad y tiempos localmente.

## Archivos del flujo

- `EjecutarAutomatizacion.ps1`: inicia y finaliza el proceso completo.
- `ProcesoUsuariosV4.ps1`: validación, Microsoft Graph, Excel y acciones.
- `CapturaAzureRapido.py`: navegación y capturas del portal.
- `IniciarPortalTemporal.ps1`: inicia Chrome con depuración remota.
- `CerrarPortalTemporal.py`: cierra la sesión de Chrome del lote.
- `TenantLaboratorio.example.txt`: formato del tenant autorizado.
- `Evidencias.xlsx`: plantilla limpia del libro operativo.

## Requisitos

- Windows 10 u 11.
- Windows PowerShell 5.1 o PowerShell 7.
- Microsoft Excel instalado.
- Google Chrome instalado.
- Python 3.10 o posterior.
- Permisos para consultar y eliminar usuarios en Microsoft Entra ID.
- Módulos de Microsoft Graph:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
```

- Dependencias de Python:

```powershell
py -m pip install -r requirements.txt
```

## Configuración

1. Copiar `TenantLaboratorio.example.txt` como `TenantLaboratorio.txt`.
2. Reemplazar el GUID de ejemplo por el Tenant ID autorizado.
3. Usar la plantilla `Evidencias.xlsx` incluida en el repositorio.
4. Confirmar que el libro tenga una hoja llamada `EVIDENCIAS` y estas columnas:
   `#`, `Nombre apellido`, `UPN`, `Antes`, `Despues`, `ACCION TOMADA`.
5. Ejecutar PowerShell desde `C:\AutomatizacionAzure`.

## Ejecución

Procesar todas las filas pendientes:

```powershell
.\EjecutarAutomatizacion.ps1 -Todas
```

Procesar una sola fila:

```powershell
.\EjecutarAutomatizacion.ps1 -Fila 5
```

## Seguridad

La automatización puede eliminar usuarios `Guest`. Antes de eliminar, compara
el usuario visible en Azure Portal con Microsoft Graph, valida el Object ID,
el UPN, el tipo de usuario y el tenant configurado. Utilice primero un entorno
de pruebas y mantenga respaldos del libro de evidencias.

La plantilla limpia `Evidencias.xlsx` se conserva en el repositorio. Las capturas y los datos agregados durante la operación deben permanecer localmente. Los perfiles del navegador, logs, respaldos y el Tenant ID real están excluidos mediante `.gitignore`.


