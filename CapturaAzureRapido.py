import argparse
import ctypes
import json
import re
import sys
import time
import unicodedata
from pathlib import Path

from PIL import ImageGrab
from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError
)


# ============================================================
# ARGUMENTOS
# ============================================================

parser = argparse.ArgumentParser(
    description="Captura automatica de Azure usando una sola pestaña."
)

parser.add_argument(
    "--modo",
    choices=["all", "deleted"]
)

parser.add_argument(
    "--busqueda"
)

parser.add_argument(
    "--nombre"
)

parser.add_argument(
    "--esperado",
    choices=["encontrado", "no_encontrado", "auto"]
)

parser.add_argument(
    "--salida"
)

parser.add_argument(
    "--resultado-json"
)

parser.add_argument(
    "--servidor",
    action="store_true",
    help="Mantiene Playwright conectado y recibe solicitudes JSON por stdin."
)

args = parser.parse_args()


# ============================================================
# CONFIGURACION
# ============================================================

CDP = "http://127.0.0.1:9223"

BASE_AUTOMATIZACION = Path(__file__).resolve().parent
ARCHIVO_TENANT = BASE_AUTOMATIZACION / "TenantLaboratorio.txt"

try:
    TENANT_ID = ARCHIVO_TENANT.read_text(encoding="utf-8-sig").strip()
except OSError as exc:
    raise RuntimeError(
        f"No se pudo leer el TenantId desde {ARCHIVO_TENANT}: {exc}"
    ) from exc

if not re.fullmatch(
    r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}",
    TENANT_ID,
):
    raise RuntimeError(
        f"TenantLaboratorio.txt no contiene un TenantId valido: {TENANT_ID!r}"
    )

PORTAL_TENANT = f"https://portal.azure.com/{TENANT_ID}/"

ALL_URL = (
    PORTAL_TENANT
    + "#view/Microsoft_AAD_UsersAndTenants/"
    "UserManagementMenuBlade/~/AllUsers"
)

SALIDA = None
RESULTADO_JSON = None


def configurar_solicitud(datos):

    global args, SALIDA, RESULTADO_JSON

    requeridos = (
        "modo",
        "busqueda",
        "nombre",
        "esperado",
        "salida",
    )

    faltantes = [
        campo
        for campo in requeridos
        if not datos.get(campo)
    ]

    if faltantes:
        raise RuntimeError(
            "Faltan argumentos requeridos: "
            + ", ".join(faltantes)
        )

    if datos["modo"] not in ("all", "deleted"):
        raise RuntimeError("Modo no valido.")

    if datos["esperado"] not in (
        "encontrado",
        "no_encontrado",
        "auto",
    ):
        raise RuntimeError("Valor esperado no valido.")

    args = argparse.Namespace(
        modo=datos["modo"],
        busqueda=str(datos["busqueda"]),
        nombre=str(datos["nombre"]),
        esperado=datos["esperado"],
        salida=str(datos["salida"]),
        resultado_json=(
            str(datos["resultado_json"])
            if datos.get("resultado_json")
            else None
        ),
        servidor=bool(datos.get("servidor", False)),
    )

    SALIDA = Path(args.salida)
    SALIDA.parent.mkdir(parents=True, exist_ok=True)

    RESULTADO_JSON = (
        Path(args.resultado_json)
        if args.resultado_json
        else None
    )

    if RESULTADO_JSON is not None:
        RESULTADO_JSON.parent.mkdir(parents=True, exist_ok=True)

NOMBRE_PESTANA = "AUTO_AZURE"

# Tiempo maximo para abrir el portal o cargar la vista de Entra ID
TIMEOUT_INICIAL = 4000

# Tiempo maximo esperando una busqueda
TIMEOUT_BUSQUEDA = 4000

# ------------------------------------------------------------
# AJUSTES DE VELOCIDAD / ESTABILIDAD
# ------------------------------------------------------------

# Evita aceptar inmediatamente el resultado anterior
ESPERA_MINIMA_RESULTADO = 0.30

# Cada cuanto revisar la pantalla
INTERVALO_POLL = 0.08

# El resultado debe estar correcto varias veces consecutivas
CONFIRMACIONES_ESTABLES = 2

# Tiempo adicional para que Azure termine de dibujar
# filas, iconos, columnas, etc.
ESPERA_RENDER_FINAL_MS = 500

# Tiempo para que Windows ponga Chrome realmente al frente
ESPERA_FOCO_MS = 100


# ============================================================
# WINDOWS
# ============================================================

user32 = ctypes.windll.user32


def seguro(texto):

    return str(texto).encode(
        "ascii",
        errors="backslashreplace"
    ).decode("ascii")


def normalizar(texto):

    texto_sin_tildes = "".join(
        caracter
        for caracter in unicodedata.normalize(
            "NFKD",
            str(texto)
        )
        if not unicodedata.combining(caracter)
    )

    return re.sub(
        r"\s+",
        " ",
        texto_sin_tildes
    ).strip().casefold()


def titulo_activo():

    hwnd = user32.GetForegroundWindow()

    titulo = ctypes.create_unicode_buffer(1024)

    user32.GetWindowTextW(
        hwnd,
        titulo,
        1024
    )

    return titulo.value


# ============================================================
# VALIDAR SESION
# ============================================================

def validar_sesion(page):

    limite = time.monotonic() + 30

    while True:

        url = page.url.lower()

        if (
            "portal.azure.com" in url
            and "/auth/login" not in url
        ):

            return

        if (
            "login.microsoftonline.com" in url
            or "portal.azure.com/auth/login" in url
        ):

            if time.monotonic() >= limite:

                raise RuntimeError(
                    "SESSION_EXPIRADA: Azure requiere autenticacion."
                )

            page.wait_for_timeout(500)

            continue

        raise RuntimeError(
            f"Pagina inesperada: {page.url}"
        )


# ============================================================
# UNA SOLA PESTANA
# ============================================================

def obtener_pagina_unica(context):

    pagina = None

    # --------------------------------------------------------
    # Buscar nuestra pestaña
    # --------------------------------------------------------

    for page in context.pages:

        try:

            nombre = page.evaluate(
                "() => window.name"
            )

            if nombre == NOMBRE_PESTANA:

                pagina = page
                break

        except Exception:
            pass

    # --------------------------------------------------------
    # Reutilizar pestaña Azure
    # --------------------------------------------------------

    if pagina is None:

        for page in context.pages:

            try:

                if "portal.azure.com" in page.url.lower():

                    pagina = page
                    break

            except Exception:
                pass

    # --------------------------------------------------------
    # Reutilizar primera pestaña disponible
    # --------------------------------------------------------

    if pagina is None and context.pages:

        pagina = context.pages[0]

    # --------------------------------------------------------
    # Crear solamente si no existe ninguna
    # --------------------------------------------------------

    if pagina is None:

        pagina = context.new_page()

        pagina.goto(
            PORTAL_TENANT,
            wait_until="domcontentloaded",
            timeout=TIMEOUT_INICIAL
        )

    validar_sesion(pagina)

    try:

        pagina.evaluate(
            f"window.name = '{NOMBRE_PESTANA}'"
        )

    except Exception:
        pass

    # --------------------------------------------------------
    # Garantizar UNA SOLA pestaña
    # --------------------------------------------------------

    for otra in list(context.pages):

        if otra == pagina:
            continue

        try:
            otra.close()
        except Exception:
            pass

    return pagina


# ============================================================
# DETECTAR SEARCHBOX DE VISTA ACTUAL
# ============================================================

def buscar_searchbox(
    page,
    iframe_name,
    timeout=500
):

    try:

        frame = page.frame_locator(
            f'iframe[name="{iframe_name}"]'
        )

        search = frame.get_by_role(
            "searchbox",
            name="Search users"
        )

        search.wait_for(
            state="visible",
            timeout=timeout
        )

        return frame, search

    except PlaywrightTimeoutError:

        return None, None

    except Exception:

        return None, None


# ============================================================
# OPCION DEL MENU DE USERS
# ============================================================

def obtener_opcion_menu_users(page, texto, timeout=4000):

    opcion = page.get_by_role(
        "link",
        name=texto,
        exact=True
    )

    try:

        opcion.wait_for(
            state="visible",
            timeout=min(timeout, 3000)
        )

        return opcion

    except PlaywrightTimeoutError:

        opcion = page.locator(
            f':text-is("{texto}"):visible'
        ).first

        opcion.wait_for(
            state="visible",
            timeout=timeout
        )

        return opcion


# ============================================================
# IR A ALL USERS
# ============================================================

def ir_all_users(page):

    frame, search = buscar_searchbox(
        page,
        "UsersList.ReactView"
    )

    if search is not None:

        print("Ya estamos en All users.")

        return frame, search

    # Si ya estamos dentro del modulo Users (por ejemplo, despues de
    # capturar Deleted users), cambiar directamente desde el menu lateral.
    if "microsoft_aad_usersandtenants/usermanagementmenublade" in page.url.lower():

        print("Cambiando directamente a All users...")

        all_users = obtener_opcion_menu_users(
            page,
            "All users",
            timeout=10000
        )

        all_users.click(
            timeout=7000
        )

        validar_sesion(page)

        frame = page.frame_locator(
            'iframe[name="UsersList.ReactView"]'
        )

        search = frame.get_by_role(
            "searchbox",
            name="Search users"
        )

        search.wait_for(
            state="visible",
            timeout=10000
        )

        print("All users cargado desde el menu lateral.")

        return frame, search

    print("Abriendo Microsoft Entra ID desde la barra de busqueda...")

    # La pagina de inicio puede cambiar sus iconos y accesos rapidos.
    # La barra global y el resultado exacto de Entra ID son mas estables.
    portal_search = page.get_by_role(
        "combobox",
        name=re.compile(
            r"Search resources, services, and docs",
            re.IGNORECASE
        )
    )

    portal_search.wait_for(
        state="visible",
        timeout=TIMEOUT_INICIAL
    )

    portal_search.fill("Microsoft Entra ID")

    opcion_entra = page.get_by_role(
        "option",
        name="Microsoft Entra ID",
        exact=True
    )

    opcion_entra.wait_for(
        state="visible",
        timeout=TIMEOUT_INICIAL
    )

    opcion_entra.click(
        timeout=TIMEOUT_INICIAL
    )

    page.wait_for_url(
        re.compile(
            r"Microsoft_AAD_IAM/ActiveDirectoryMenuBlade",
            re.IGNORECASE
        ),
        timeout=15000
    )

    validar_sesion(page)

    print("Microsoft Entra ID abierto. Abriendo Manage...")

    # El menu lateral se carga despues del contenido principal. Se espera
    # dinamicamente y solo se pulsa Manage cuando todavia esta contraido.
    manage = page.get_by_role(
        "button",
        name="Toggle Manage",
        exact=True
    )

    manage.wait_for(
        state="visible",
        timeout=10000
    )

    if manage.get_attribute("aria-expanded") != "true":

        manage.click(
            timeout=7000
        )

    print("Manage abierto. Seleccionando Users...")

    users = page.get_by_role(
        "link",
        name="Users",
        exact=True
    )

    try:

        users.wait_for(
            state="visible",
            timeout=5000
        )

    except PlaywrightTimeoutError:

        # Azure puede renderizar el texto dentro de un control dinamico
        # sin exponer inmediatamente el rol de enlace.
        users = page.locator(
            ':text-is("Users"):visible'
        ).first

        users.wait_for(
            state="visible",
            timeout=20000
        )

    users.click(
        timeout=7000
    )

    validar_sesion(page)

    print("Users abierto. Seleccionando All users...")

    all_users = obtener_opcion_menu_users(
        page,
        "All users",
        timeout=30000
    )

    all_users.click(
        timeout=7000
    )

    validar_sesion(page)

    frame = page.frame_locator(
        'iframe[name="UsersList.ReactView"]'
    )

    search = frame.get_by_role(
        "searchbox",
        name="Search users"
    )

    search.wait_for(
        state="visible",
        timeout=TIMEOUT_INICIAL
    )

    print("All users cargado.")

    return frame, search


# ============================================================
# IR A DELETED USERS
# ============================================================

def ir_deleted_users(page):

    frame, search = buscar_searchbox(
        page,
        "DeletedUsersList.ReactView"
    )

    if search is not None:

        print("Ya estamos en Deleted users.")

        return frame, search

    print("Cambiando a Deleted users...")

    # --------------------------------------------------------
    # Intentar menu lateral directamente
    # --------------------------------------------------------

    try:

        enlace = obtener_opcion_menu_users(
            page,
            "Deleted users",
            timeout=6000
        )

        enlace.click(
            timeout=3000
        )

    except Exception:

        # Si no aparece el enlace, primero ir a Users
        ir_all_users(page)

        enlace = obtener_opcion_menu_users(
            page,
            "Deleted users",
            timeout=10000
        )

        enlace.click(
            timeout=3000
        )

    frame = page.frame_locator(
        'iframe[name="DeletedUsersList.ReactView"]'
    )

    search = frame.get_by_role(
        "searchbox",
        name="Search users"
    )

    search.wait_for(
        state="visible",
        timeout=7000
    )

    validar_sesion(page)

    return frame, search


# ============================================================
# ESPERAR RESULTADO ESTABLE
# ============================================================

def esperar_resultado(
    frame,
    nombre,
    busqueda,
    esperado,
    momento_busqueda
):

    nombre_n = normalizar(nombre)
    busqueda_n = normalizar(busqueda)

    body = frame.locator("body")

    limite = (
        time.monotonic()
        + (TIMEOUT_BUSQUEDA / 1000)
    )

    confirmaciones = 0
    estado_anterior = None

    while time.monotonic() < limite:

        # ----------------------------------------------------
        # No aceptar inmediatamente el resultado anterior
        # ----------------------------------------------------

        transcurrido = (
            time.monotonic()
            - momento_busqueda
        )

        espera_minima = ESPERA_MINIMA_RESULTADO

        if esperado == "auto":
            espera_minima = max(espera_minima, 1.5)

        if transcurrido < espera_minima:

            time.sleep(
                min(
                    INTERVALO_POLL,
                    espera_minima - transcurrido
                )
            )

            continue

        try:

            texto = body.inner_text(
                timeout=1000
            )

        except Exception:

            confirmaciones = 0

            time.sleep(
                INTERVALO_POLL
            )

            continue

        texto_n = normalizar(
            texto
        )

        nombre_visible = (
            nombre_n in texto_n
        )

        busqueda_visible = (
            busqueda_n in texto_n
        )

        contador = re.search(
            r"\b(\d+)\s+users?\s+found\b",
            texto_n
        )

        cantidad = None

        if contador:

            cantidad = int(
                contador.group(1)
            )

        sin_resultados = any(
            mensaje in texto_n
            for mensaje in (
                "0 users found",
                "0 user found",
                "no users found",
                "no user found",
                "no results",
                "no results found",
                "there are no users",
                "there are no items to show",
                "no items to display",
                "we couldn't find any results",
                "we could not find any results"
            )
        )

        condicion_correcta = False
        estado_actual = None

        # ----------------------------------------------------
        # USUARIO ENCONTRADO
        # ----------------------------------------------------

        if esperado == "encontrado":

            if nombre_visible or busqueda_visible:

                condicion_correcta = True
                estado_actual = "encontrado"

        # ----------------------------------------------------
        # DETECCION AUTOMATICA DESDE EL PORTAL
        # ----------------------------------------------------

        elif esperado == "auto":

            # Azure puede devolver varias coincidencias parciales aunque el
            # filtro contenga un UPN completo. La seleccion exacta se valida
            # despues en extraer_usuario_visible; aqui basta con confirmar
            # que la fila buscada ya esta visible y estable.
            if (
                nombre_visible
                or busqueda_visible
                or (
                    cantidad is not None
                    and cantidad >= 1
                )
            ):

                condicion_correcta = True
                estado_actual = "encontrado"
            elif (
                sin_resultados
                or (
                    cantidad is not None
                    and cantidad == 0
                )
            ):

                condicion_correcta = True
                estado_actual = "no_encontrado"

        # ----------------------------------------------------
        # USUARIO NO ENCONTRADO
        # ----------------------------------------------------

        else:

            if sin_resultados:

                condicion_correcta = True
                estado_actual = "no_encontrado"

            elif (
                cantidad is not None
                and cantidad == 0
            ):

                condicion_correcta = True
                estado_actual = "no_encontrado"

        # ----------------------------------------------------
        # EXIGIR ESTABILIDAD
        # ----------------------------------------------------

        if condicion_correcta:

            if estado_actual != estado_anterior:
                confirmaciones = 0

            estado_anterior = estado_actual

            confirmaciones += 1

            print(
                "Resultado estable:",
                f"{confirmaciones}/{CONFIRMACIONES_ESTABLES}"
            )

            if (
                confirmaciones
                >= CONFIRMACIONES_ESTABLES
            ):

                return estado_actual

        else:

            # Si Azure vuelve a cambiar,
            # comenzar conteo nuevamente.
            confirmaciones = 0
            estado_anterior = None

        time.sleep(
            INTERVALO_POLL
        )

    # --------------------------------------------------------
    # TIMEOUT
    # --------------------------------------------------------

    if esperado == "encontrado":

        raise RuntimeError(
            "Azure no mostro de forma estable "
            "el usuario esperado."
        )

    if esperado == "auto":

        raise RuntimeError(
            "Azure no permitio determinar de forma estable "
            "si el usuario existe."
        )

    raise RuntimeError(
        "Azure no confirmo de forma estable "
        "cero resultados."
    )


# ============================================================
# EXTRAER IDENTIDAD DE LA FILA VISIBLE
# ============================================================

def extraer_usuario_visible(frame, nombre, busqueda):

    nombre_n = normalizar(nombre)
    busqueda_n = normalizar(busqueda)
    filas = frame.locator('[role="row"]')
    coincidencias_upn = []
    coincidencias_nombre = []

    for indice in range(1, filas.count()):

        fila = filas.nth(indice)
        celdas = fila.locator('[role="gridcell"]')

        if celdas.count() < 4:
            continue

        enlaces = fila.locator("a")

        if enlaces.count() != 1:
            continue

        enlace = enlaces.first
        href = enlace.get_attribute("href") or ""
        object_id_match = re.search(
            r"/userId/([0-9a-fA-F-]{36})(?:/|$)",
            href
        )

        upn_texto = celdas.nth(2).inner_text(timeout=3000)
        upn_match = re.search(
            r"[^\s|]+@[^\s|]+",
            upn_texto
        )

        nombre_visible = enlace.inner_text(timeout=3000)

        if upn_match is None:
            continue

        nombre_coincide = (
            normalizar(nombre_visible) == nombre_n
        )

        upn_coincide = (
            normalizar(upn_match.group(0)) == busqueda_n
        )

        if not nombre_coincide and not upn_coincide:
            continue

        tipo = celdas.nth(3).inner_text(timeout=3000).strip()

        if object_id_match is None or upn_match is None:
            raise RuntimeError(
                "La fila encontrada no expone un Object ID o UPN valido."
            )

        if tipo not in ("Guest", "Member"):
            raise RuntimeError(
                f"Tipo de usuario no reconocido en Azure: {tipo!r}"
            )

        resultado = {
            "encontrado": True,
            "displayName": nombre_visible.strip(),
            "userPrincipalName": upn_match.group(0),
            "userType": tipo,
            "objectId": object_id_match.group(1),
            "displayNameCoincide": nombre_coincide,
            "upnCoincide": upn_coincide,
        }

        if upn_coincide:
            coincidencias_upn.append(resultado)

        if nombre_coincide:
            coincidencias_nombre.append(resultado)

    # El UPN del Excel es el identificador principal. Si Azure muestra otras
    # filas por coincidencia parcial, se ignoran siempre que exista una unica
    # coincidencia exacta del UPN.
    if len(coincidencias_upn) == 1:
        return coincidencias_upn[0]

    if len(coincidencias_upn) > 1:
        raise RuntimeError(
            "Azure mostro mas de una fila con el mismo UPN exacto."
        )

    # Compatibilidad para registros cuyo nombre coincide pero cuyo UPN
    # visible tenga un formato especial en el portal.
    if len(coincidencias_nombre) == 1:
        return coincidencias_nombre[0]

    raise RuntimeError(
        "Azure no mostro una unica fila que coincida exactamente "
        "con el UPN o el nombre del Excel."
    )


# ============================================================
# EJECUTAR BUSQUEDA
# ============================================================

def ejecutar_busqueda(
    page,
    frame,
    search
):

    page.bring_to_front()

    print("")
    print(
        "Buscando:",
        seguro(args.busqueda)
    )

    # --------------------------------------------------------
    # Introducir busqueda
    # --------------------------------------------------------

    # Limpiar completamente el filtro anterior. En sesiones largas Azure
    # puede conservar internamente el ultimo valor aunque Playwright cambie
    # el texto visible.
    search.click(timeout=3000)
    search.press("Control+A")
    search.press("Backspace")
    page.wait_for_timeout(250)

    # Escribir como teclado real genera los eventos input/keyup que utiliza
    # la tabla de Azure para lanzar una busqueda nueva. Enter confirma el
    # filtro cuando el portal no reacciona solo al cambio de texto.
    search.type(
        args.busqueda,
        delay=20
    )
    search.press("Enter")

    momento_busqueda = time.monotonic()

    # --------------------------------------------------------
    # Validar contenido del buscador
    # --------------------------------------------------------

    valor = search.input_value(
        timeout=2500
    )

    if valor != args.busqueda:

        raise RuntimeError(
            "El buscador no recibio "
            "el valor correcto."
        )

    # --------------------------------------------------------
    # Esperar resultado REAL Y ESTABLE
    # --------------------------------------------------------

    estado = esperar_resultado(
        frame,
        args.nombre,
        args.busqueda,
        args.esperado,
        momento_busqueda
    )

    # --------------------------------------------------------
    # IMPORTANTE
    #
    # El resultado ya existe, pero Azure todavía puede estar
    # terminando de dibujar:
    #
    # - filas
    # - columnas
    # - iconos
    # - enlaces
    # - barra azul
    #
    # Esperamos solamente 500 ms.
    # --------------------------------------------------------

    print(
        "Resultado confirmado. "
        "Esperando render final..."
    )

    page.wait_for_timeout(
        ESPERA_RENDER_FINAL_MS
    )

    return estado


# ============================================================
# CAPTURAR
# ============================================================

def capturar(page):

    page.bring_to_front()

    # --------------------------------------------------------
    # Dar tiempo solamente al foco de Windows
    # --------------------------------------------------------

    page.wait_for_timeout(
        ESPERA_FOCO_MS
    )

    titulo = titulo_activo()

    print(
        "Ventana activa:",
        seguro(titulo)
    )

    if (
        "Azure" not in titulo
        and
        "Chrome" not in titulo
    ):

        raise RuntimeError(
            "Azure no quedo en primer plano."
        )

    # --------------------------------------------------------
    # CAPTURA COMPLETA
    # --------------------------------------------------------

    imagen = ImageGrab.grab(
        all_screens=False
    )

    imagen.save(
        SALIDA
    )

    # --------------------------------------------------------
    # VALIDAR ARCHIVO
    # --------------------------------------------------------

    if not SALIDA.exists():

        raise RuntimeError(
            "No se creo la captura."
        )

    if SALIDA.stat().st_size < 10000:

        raise RuntimeError(
            "La captura parece estar vacia."
        )


# ============================================================
# MAIN
# ============================================================

def conectar_browser(playwright):

    browser = playwright.chromium.connect_over_cdp(CDP)

    if not browser.contexts:
        raise RuntimeError("No existe contexto Chrome disponible.")

    return browser


def procesar_captura(browser):

    inicio = time.monotonic()

    print("")
    print("======================================")
    print("CAPTURA AZURE RAPIDA")
    print("UNA SOLA PESTANA")
    print("RESULTADO ESTABLE")
    print("======================================")

    context = browser.contexts[0]
    page = obtener_pagina_unica(context)
    validar_sesion(page)

    if args.modo == "all":
        frame, search = ir_all_users(page)
    else:
        frame, search = ir_deleted_users(page)

    estado_resultado = ejecutar_busqueda(
        page,
        frame,
        search
    )

    resultado_portal = None

    if RESULTADO_JSON is not None:

        if args.modo != "all" or args.esperado != "auto":
            raise RuntimeError(
                "--resultado-json requiere --modo all --esperado auto."
            )

        if estado_resultado == "encontrado":
            resultado_portal = extraer_usuario_visible(
                frame,
                args.nombre,
                args.busqueda
            )
        else:
            resultado_portal = {
                "encontrado": False,
                "displayName": None,
                "userPrincipalName": None,
                "userType": None,
                "objectId": None,
                "displayNameCoincide": False,
                "upnCoincide": False,
            }

    capturar(page)

    if RESULTADO_JSON is not None:
        RESULTADO_JSON.write_text(
            json.dumps(
                resultado_portal,
                ensure_ascii=False,
                indent=2
            ),
            encoding="utf-8"
        )

    duracion = time.monotonic() - inicio

    print("")
    print("CAPTURA GENERADA:")
    print(SALIDA)
    print(f"Tiempo captura: {duracion:.2f} segundos")


def imprimir_error(exc):

    print("")
    print("======================================")
    print("ERROR CAPTURA AZURE")
    print("======================================")
    print(seguro(str(exc)))


def ejecutar_individual():

    configurar_solicitud(vars(args))

    with sync_playwright() as playwright:
        browser = conectar_browser(playwright)
        procesar_captura(browser)


def ejecutar_servidor():

    with sync_playwright() as playwright:

        browser = conectar_browser(playwright)
        print("__CAPTURA_READY__", flush=True)

        for linea in sys.stdin:

            linea = linea.lstrip("\ufeff").strip()

            if not linea:
                continue

            solicitud = json.loads(linea)

            if solicitud.get("comando") == "cerrar":
                print(
                    "__CAPTURA_RESULT__" + json.dumps({"ok": True}),
                    flush=True
                )
                return

            try:
                configurar_solicitud(solicitud)
                procesar_captura(browser)
                respuesta = {"ok": True}
            except Exception as exc:
                imprimir_error(exc)
                respuesta = {
                    "ok": False,
                    "error": seguro(str(exc)),
                }

            print(
                "__CAPTURA_RESULT__" + json.dumps(
                    respuesta,
                    ensure_ascii=True
                ),
                flush=True
            )


if __name__ == "__main__":

    try:

        if args.servidor:
            ejecutar_servidor()
        else:
            ejecutar_individual()

    except Exception as exc:
        imprimir_error(exc)
        sys.exit(1)

