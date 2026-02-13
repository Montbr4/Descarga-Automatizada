import os
import time
import shutil
from datetime import datetime
import pytz

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


URL = "https://visitas.servicios.gob.pe/consultas/index.php?ruc_enti=20504743307"

zona = pytz.timezone("America/Lima")
fecha = datetime.now(zona).strftime("%Y-%m-%d")

BASE = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE, "data", fecha)
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

FINAL_FILE = os.path.join(
    DOWNLOAD_DIR,
    f"reporte_visitas_{fecha}.xlsx"
)

print("Carpeta:", DOWNLOAD_DIR)

opts = Options()
opts.binary_location = "/usr/bin/google-chrome"
opts.add_argument("--headless=new")
opts.add_argument("--no-sandbox")
opts.add_argument("--disable-dev-shm-usage")
opts.add_argument("--window-size=1920,1080")

prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
}
opts.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=opts)

driver.execute_cdp_cmd(
    "Page.setDownloadBehavior",
    {"behavior": "allow", "downloadPath": DOWNLOAD_DIR}
)

wait = WebDriverWait(driver, 60)

try:
    print("Abriendo portal")
    driver.get(URL)

    print("Esperando datos reales en la tabla")

wait.until(lambda d: d.execute_script("""
    let table = document.querySelector('table.dataTable');
    if (!table) return false;

    let rows = table.querySelectorAll('tbody tr');
    return rows.length > 0;
"""))

rows = driver.execute_script("""
    return document.querySelectorAll(
        'table.dataTable tbody tr'
    ).length;
""")

print(f"Filas detectadas: {rows}")

    print("Esperando botón Excel")

    excel_btn = wait.until(EC.element_to_be_clickable((
        By.CSS_SELECTOR, "button.buttons-excel"
    )))

    archivos_antes = set(os.listdir(DOWNLOAD_DIR))

    print("Descargando Excel")
    driver.execute_script("arguments[0].click();", excel_btn)

    timeout = 120
    start = time.time()
    descargado = None

    while time.time() - start < timeout:
        archivos = set(os.listdir(DOWNLOAD_DIR))
        nuevos = archivos - archivos_antes

        for f in nuevos:
            if not f.endswith(".crdownload"):
                descargado = f
                break

        if descargado:
            break

        time.sleep(1)

    if not descargado:
        raise Exception("Tiempo agotado — descarga no detectada")

    origen = os.path.join(DOWNLOAD_DIR, descargado)

    if os.path.exists(FINAL_FILE):
        os.remove(FINAL_FILE)

    shutil.move(origen, FINAL_FILE)

    print("DESCARGA EXITOSA")
    print("Archivo:", FINAL_FILE)
    print("Tamaño:", os.path.getsize(FINAL_FILE), "bytes")

except Exception as e:
    print("ERROR:", e)
    driver.save_screenshot(
        os.path.join(DOWNLOAD_DIR, "debug_error.png")
    )

finally:
    driver.quit()
    print("Fin")
