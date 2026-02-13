import asyncio
from playwright.async_api import async_playwright
import os
from datetime import datetime

URL = "https://visitas.servicios.gob.pe/consultas/index.php?ruc_enti=20504743307"

async def main():
    print("Iniciando")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        print("Abriendo portal...")
        await page.goto(URL, timeout=120000)

        print("Click buscar")
        await page.click("text=Buscar")

        print("Esperando datos")
        await page.wait_for_selector("text=Mostrando", timeout=60000)

        print("Descargando Excel")

        async with page.expect_download() as download_info:
            await page.click("text=Excel")

        download = await download_info.value

        fecha = datetime.now().strftime("%Y-%m-%d")
        nombre = f"reporte_visitas_{fecha}.xlsx"

        await download.save_as(nombre)

        print("Descarga completada:", nombre)

        await browser.close()

asyncio.run(main())
