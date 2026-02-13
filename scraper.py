import os
import time
from datetime import datetime
import pytz

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

URL = "https://visitas.servicios.gob.pe/consultas/index.php?ruc_enti=20504743307"

zona = pytz.timezone("America/Lima")
fecha = datetime.now(zona).strftime("%Y-%m-%d")

BASE = os.getcwd()
DOWNLOAD = os.path.join(BASE, "data", fecha)
os.makedirs(DOWNLOAD, exist_ok=True)

print("INICIO SCRAPER")

options = Options()
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")

prefs = {
    "download.default_directory": DOWNLOAD,
    "download.prompt_for_download": False,
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

wait = WebDriverWait(driver, 60)

try:

    print("Abriendo portal")
    driver.get(URL)

    print("Esperando botón Buscar")
    buscar = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[contains(.,'Buscar')]")
    ))

    time.sleep(2)
    buscar.click()

    print("Esperando datos reales")

    datos_ok = False

    for intento in range(4):

        print(f"   Intento {intento+1}")

        try:

            wait.until(lambda d: d.execute_script("""
                let table = document.querySelector("table.dataTable");
                if (!table) return false;

                let txt = table.innerText;

                if (txt.includes("Cargando")) return false;
                if (txt.includes("No hay datos")) return false;

                let rows = table.querySelectorAll("tbody tr");
                return rows.length > 0;
            """))

            datos_ok = True
            print("Datos detectados")
            break

        except:
            print("Datos no listos — recargando")
            driver.refresh()
            time.sleep(5)

            buscar = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(.,'Buscar')]")
            ))
            buscar.click()

    if not datos_ok:
        raise Exception("La tabla nunca cargó datos reales")

    print("Descargando Excel")

    excel = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[contains(.,'Excel')]")
    ))

    archivos_antes = set(os.listdir(DOWNLOAD))
    excel.click()

    timeout = 60
    descargado = False

    while timeout > 0:

        archivos = set(os.listdir(DOWNLOAD))
        nuevos = archivos - archivos_antes

        for f in nuevos:
            if not f.endswith(".crdownload"):
                print("Archivo descargado:", f)
                descargado = True
                break

        if descargado:
            break

        time.sleep(1)
        timeout -= 1

    if not descargado:
        raise Exception("No se detectó descarga")

    print("SCRAPER COMPLETADO")

except Exception as e:

    print("\n ERROR:", e)
    driver.save_screenshot(os.path.join(DOWNLOAD, "debug_error.png"))
    print("Screenshot guardado")

finally:

    driver.quit()
    print("FIN")
