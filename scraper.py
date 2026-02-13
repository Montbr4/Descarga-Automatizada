import os
import time
import shutil
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

print("Iniciando")

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

wait = WebDriverWait(driver, 40)

try:

    print("Abriendo portal")
    driver.get(URL)

    print("Click Buscar")
    btn = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[contains(.,'Buscar')]")
    ))

    time.sleep(2)
    btn.click()

    print("Esperando datos reales")

    wait.until(lambda d: d.execute_script("""
        let table = document.querySelector("table.dataTable");
        if (!table) return false;

        let rows = table.querySelectorAll("tbody tr");
        if (rows.length === 0) return false;

        return !rows[0].innerText.includes("No hay datos");
    """))

    print("Datos cargados")

    excel = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[contains(.,'Excel')]")
    ))

    print("Descargando Excel")
    before = set(os.listdir(DOWNLOAD))
    excel.click()

    timeout = 60
    while timeout > 0:
        after = set(os.listdir(DOWNLOAD))
        new = after - before

        if new:
            file = list(new)[0]
            if not file.endswith(".crdownload"):
                print("Descarga completa:", file)
                break

        time.sleep(1)
        timeout -= 1

    print("Fin")

finally:
    driver.quit()
