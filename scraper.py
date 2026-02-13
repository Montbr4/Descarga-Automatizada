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

URL = "https://www.transparencia.gob.pe/reportes_directos/pte_transparencia_reg_visitas.aspx?id_entidad=11476&ver=&id_tema=500"

zona = pytz.timezone("America/Lima")
hoy = datetime.now(zona).strftime("%Y-%m-%d")

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data", hoy)
os.makedirs(DATA, exist_ok=True)

FINAL_FILE = os.path.join(DATA, f"reporte_visitas_{hoy}.xlsx")

opts = Options()
opts.binary_location = "/usr/bin/google-chrome"

opts.add_argument("--headless=new")
opts.add_argument("--no-sandbox")
opts.add_argument("--disable-dev-shm-usage")
opts.add_argument("--window-size=1920,1080")

prefs = {
    "download.default_directory": DATA,
    "download.prompt_for_download": False,
}
opts.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=opts)
driver.execute_cdp_cmd("Emulation.setTimezoneOverride", {"timezoneId": "America/Lima"})

wait = WebDriverWait(driver, 30)

try:
    print("Abriendo portal")
    driver.get(URL)

    time.sleep(5)

    print("Navegando a consulta")
    btn = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//a[contains(., 'Ir a consultar')]")
    ))
    driver.execute_script("arguments[0].click();", btn)

    time.sleep(12)

    print("Buscando botón Excel")
    excel_btn = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[contains(@class,'excel')]")
    ))

    before = set(os.listdir(DATA))
    driver.execute_script("arguments[0].click();", excel_btn)

    print("Esperando descarga")

    for _ in range(60):
        now = set(os.listdir(DATA))
        diff = now - before

        for f in diff:
            if not f.endswith(".crdownload"):
                shutil.move(os.path.join(DATA, f), FINAL_FILE)
                print("Descarga completa:", FINAL_FILE)
                raise SystemExit

        time.sleep(1)

    print("No se descargó archivo")

finally:
    driver.quit()
    print("✔ Fin")
