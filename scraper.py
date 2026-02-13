import os
import time
import shutil
from datetime import datetime, timedelta
import pytz

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


URL = "https://visitas.servicios.gob.pe/consultas/index.php?ruc_enti=20504743307"

zona = pytz.timezone("America/Lima")
hoy = datetime.now(zona)
ayer = hoy - timedelta(days=1)

fecha_inicio = ayer.strftime("%d/%m/%Y")
fecha_fin = hoy.strftime("%d/%m/%Y")
fecha_folder = hoy.strftime("%Y-%m-%d")

BASE = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE, "data", fecha_folder)
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

FINAL_FILE = os.path.join(
    DOWNLOAD_DIR,
    f"reporte_visitas_{fecha_folder}.xlsx"
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

    print("Configurando rango de fechas")

    fecha_input = wait.until(EC.presence_of_element_located((
        By.CSS_SELECTOR, "input[name='daterange']"
    )))

    driver.execute_script("arguments[0].value='';", fecha_input)
    fecha_input.send_keys(f"{fecha_inicio} - {fecha_fin}")

    print("Ejecutando búsqueda")

    buscar_btn = driver.find_element(
        By.XPATH, "//button[contains(.,'Buscar')]"
    )
    buscar_btn.click()

    time.sleep(5)

    print("Esperando datos")

    wait.until(lambda d: d.execute_script("""
        if (!window.jQuery || !$.fn.dataTable) return false;
        let api = $( $.fn.dataTable.tables()[0] ).DataTable();
        return api.data().length >= 0;
    """))

    registros = driver.execute_script("""
        let api = $( $.fn.dataTable.tables()[0] ).DataTable();
        return api.data().length;
    """)

    print(f"Registros encontrados: {registros}")

    if registros == 0:
        raise Exception("No hay registros para exportar")
        
    excel_btn = wait.until(EC.element_to_be_clickable((
        By.CSS_SELECTOR, "button.buttons-excel"
    )))

    before = set(os.listdir(DOWNLOAD_DIR))

    print("Descargando Excel")
    driver.execute_script("arguments[0].click();", excel_btn)

    timeout = 120
    start = time.time()
    downloaded = None

    while time.time() - start < timeout:
        files = set(os.listdir(DOWNLOAD_DIR))
        diff = files - before

        for f in diff:
            if not f.endswith(".crdownload"):
                downloaded = f
                break

        if downloaded:
            break

        time.sleep(1)

    if not downloaded:
        raise Exception("Descarga no detectada")

    shutil.move(
        os.path.join(DOWNLOAD_DIR, downloaded),
        FINAL_FILE
    )

    print("DESCARGA EXITOSA")

except Exception as e:

    print("ERROR:", e)
    driver.save_screenshot(
        os.path.join(DOWNLOAD_DIR, "debug_error.png")
    )

finally:

    driver.quit()
    print("Fin del proceso")
