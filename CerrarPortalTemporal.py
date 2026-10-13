import sys
from playwright.sync_api import sync_playwright


CDP = "http://127.0.0.1:9223"


try:

    with sync_playwright() as p:

        browser = p.chromium.connect_over_cdp(
            CDP
        )

        print("Cerrando Google Chrome...")

        browser.close()

        print("Chrome cerrado.")

        sys.exit(0)


except Exception as exc:

    print(
        "Chrome ya estaba cerrado o no fue posible conectar."
    )

    print(str(exc))

    # No convertir el cierre en fallo del lote
    sys.exit(0)

