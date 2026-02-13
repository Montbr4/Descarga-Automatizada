import os
import time
import shutil
from datetime import datetime
import pytz

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options


URL = "https://www.transparencia.gob.pe/reportes_directos/pte_transparencia_reg_visitas.aspx?id_entidad=11476&ver=&id_tema=500"

zona = pytz.timezone("America/Lima")
fecha = datetime.now(zona).strftime("%Y-%m-%d")

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data", fecha)
os.makedirs(DATA, exist_ok=True)

FINAL_NAME = f"reporte_visitas_{fecha}.xlsx"
FINAL_PATH = os.path.join(DATA, FINAL_NAME)

print("Carpeta:", DATA)

options = Options()
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")

prefs = {
    "download.default_directory": DATA,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
}
options.add_experimental_option("prefs", prefs)

service = Service("/usr/bin/chromedriver")
driver = webdriver.Chrome(service=service, options=options)

driver.execute_cdp_cmd(
    "Page.setDownloadBehavior",
    {"behavior": "allow", "downloadPath": DATA}
)

wait = WebDriverWait(driver, 40)

try:
    print("Abriendo portal...")
    driver.get(URL)

    btn = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//a[contains(.,'consultar')]"
    )))
    btn.click()

    print("Esperando tabla...")
    wait.until(EC.presence_of_element_located((
        By.XPATH, "//table"
    )))

    try:
        buscar = driver.find_element(
            By.XPATH, "//button[contains(.,'Buscar') or contains(.,'BUSCAR')]"
        )
        buscar.click()
        time.sleep(5)
    except:
        pass

    print("Buscando botón Excel")

    excel = wait.until(EC.element_to_be_clickable((
        By.XPATH,
        "//button[contains(@class,'excel') or contains(.,'Excel')]"
    )))

    before = set(os.listdir(DATA))
    excel.click()

    print("Esperando descarga")

    timeout = 90
    start = time.time()

    downloaded = None

    while time.time() - start < timeout:
        files = set(os.listdir(DATA))
        new = files - before

        for f in new:
            if not f.endswith(".crdownload"):
                downloaded = f
                break

        if downloaded:
            break

        time.sleep(1)

    if not downloaded:
        raise Exception("No se descargó archivo")

    src = os.path.join(DATA, downloaded)

    if os.path.exists(FINAL_PATH):
        os.remove(FINAL_PATH)

    shutil.move(src, FINAL_PATH)

    print("DESCARGA EXITOSA")
    print(FINAL_PATH)

except Exception as e:
    print("ERROR:", e)
    driver.save_screenshot(os.path.join(DATA, "error.png"))

finally:
    driver.quit()
    print("Proceso terminado")
