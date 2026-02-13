import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

URL = "https://visitas.servicios.gob.pe/consultas/index.php?ruc_enti=20504743307"


def main():
    print("Inicio")

    download_path = os.path.abspath("data")
    os.makedirs(download_path, exist_ok=True)

    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")

    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
    }

    chrome_options.add_experimental_option("prefs", prefs)

    print("Abriendo navegador")
    driver = webdriver.Chrome(options=chrome_options)

    try:
        driver.get(URL)

        print("Esperando carga de tabla")

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#tablaBusqueda tbody tr"))
        )

        print("Tabla detectada")

        time.sleep(3)

        print("Buscando botón Excel")

        excel_btn = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Excel')]"))
        )

        print("Descargando Excel")
        excel_btn.click()

        time.sleep(10)

        print("Descarga finalizada")

    finally:
        driver.quit()

    print("Proceso terminado")


if __name__ == "__main__":
    main()
