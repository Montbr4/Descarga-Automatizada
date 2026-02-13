import time
import os
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

URL = "https://visitas.servicios.gob.pe/consultas/index.php?ruc_enti=20504743307"

DOWNLOAD_DIR = os.getcwd()

def main():

    print("Iniciando")

    options = uc.ChromeOptions()

    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")

    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
    }

    options.add_experimental_option("prefs", prefs)

    driver = uc.Chrome(options=options)
    wait = WebDriverWait(driver, 40)

    try:

        print("Abriendo portal")
        driver.get(URL)

        print("Esperando botón Buscar")

        buscar = wait.until(lambda d: d.find_element(
            By.XPATH, "//button[contains(., 'Buscar')]"
        ))

        time.sleep(2)

        print("Ejecutando búsqueda")
        buscar.click()

        print("Esperando datos reales")

        wait.until(lambda d: d.execute_script("""
            let table = document.querySelector("table.dataTable");
            if (!table) return false;

            let rows = table.querySelectorAll("tbody tr");
            if (rows.length === 0) return false;

            return !rows[0].innerText.includes("No hay datos");
        """))

        print("Datos cargados")

        excel = wait.until(lambda d: d.find_element(
            By.XPATH, "//button[contains(., 'Excel')]"
        ))

        print("Descargando Excel")
        excel.click()

        time.sleep(10)

        print("Descarga finalizada")

    finally:
        driver.quit()
        print("Fin")

if __name__ == "__main__":
    main()
