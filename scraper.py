import asyncio
from playwright.async_api import async_playwright
import os
from datetime import datetime
import pytz

URL = "https://visitas.servicios.gob.pe/consultas/index.php?ruc_enti=20504743307"

async def main():

    print("\nINICIO SCRAPER")

    zona = pytz.timezone("America/Lima")
    fecha = datetime.now(zona).strftime("%Y-%m-%d")

    base_dir = os.getcwd()
    carpeta = os.path.join(base_dir, "data")

    print("Directorio actual:", base_dir)
    print("Creando carpeta:", carpeta)

    os.makedirs(carpeta, exist_ok=True)

    nombre = f"reporte_visitas_{fecha}.xlsx"
    ruta = os.path.join(carpeta, nombre)

    async with async_playwright() as p:

        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        print("Abriendo portal")
        await page.goto(URL, timeout=120000)

        print("Click buscar")
        await page.click("text=Buscar")

        print("Esperando datos")
        await page.wait_for_selector("text=Mostrando", timeout=60000)

        print("Descargando Excel")

        async with page.expect_download() as download_info:
            await page.click("text=Excel")

        download = await download_info.value
        await download.save_as(ruta)

        print("\nARCHIVO GUARDADO:")
        print(ruta)

        await browser.close()

    print("\nContenido de /data:")
    print(os.listdir(carpeta))

    print("\nFIN")

asyncio.run(main())
