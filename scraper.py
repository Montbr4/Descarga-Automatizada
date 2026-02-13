import os
import time
from datetime import datetime
import pytz

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


URL = "https://visitas.servicios.gob.pe/consultas/index.php?ruc_enti=20504743307"


def main():

    print("Simulando usuario humano")


    zona = pytz.timezone("America/Lima")
    hoy = datetime.now(zona).strftime("%Y-%m-%d")

    base = os.path.dirname(os.path.abspath(__file__))
    carpeta = os.path.join(base, "data", hoy)

    os.makedirs(carpeta, exist_ok=True)


    options = Options()

    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")

    options.add_argument("--disable-blink-features=AutomationControlled")

    prefs = {
        "download.default_directory": carpeta,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
    }

    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=options)

    wait = WebDriverWait(driver, 40)

    try:

        print("Abriendo portal")
        driver.get(URL)

        time.sleep(6)


        print("Click Buscar")

        buscar = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Buscar')]"))
        )

        driver.execute_script("arguments[0].click();", buscar)

        print("Esperando datos")

        wait.until(
            EC.presence_of_element_located((By.XPATH, "//table//tbody/tr"))
        )

        time.sleep(3)


        print("Descargando Excel")

        excel = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Excel')]"))
        )

        driver.execute_script("arguments[0].click();", excel)


        print("Esperando archivo")

        tiempo = 0

        while tiempo < 60:

            archivos = os.listdir(carpeta)

            if any(a.endswith(".xlsx") for a in archivos):
                print("Excel descargado correctamente")
                break

            time.sleep(1)
            tiempo += 1

        print("Proceso completado")

    finally:

        driver.quit()


if __name__ == "__main__":
    main()
