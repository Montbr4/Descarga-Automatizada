import os
import time
import random
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


URL = "https://visitas.servicios.gob.pe/consultas/index.php?ruc_enti=20504743307"


def human_delay(a=1, b=3):
    time.sleep(random.uniform(a, b))


def main():

    print("Simulando usuario humano")

    options = Options()


    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 60)

    try:

        print("Abriendo portal")
        driver.get(URL)

        human_delay(5, 7)

        driver.execute_script("window.scrollTo(0, 500)")
        human_delay()

        print("Click Buscar")
        buscar = wait.until(
            EC.element_to_be_clickable((By.ID, "btnBuscar"))
        )
        buscar.click()

        print("Esperando datos reales")

        wait.until(lambda d: len(
            d.find_elements(By.CSS_SELECTOR, "#tablaResultados tbody tr")
        ) > 1)

        human_delay()

        print("Extrayendo tabla")

        rows = driver.find_elements(By.CSS_SELECTOR, "#tablaResultados tbody tr")

        data = []

        for r in rows:
            cols = [c.text.strip() for c in r.find_elements(By.TAG_NAME, "td")]
            if cols:
                data.append(cols)

        if not data:
            raise Exception("Bloqueo detectado — tabla vacía")

        headers = [
            "Fecha Registro","Fecha Visita","Entidad","Visitante","Documento",
            "Entidad Visitante","Funcionario","Hora Ingreso","Hora Salida",
            "Motivo","Lugar","Observación"
        ]

        df = pd.DataFrame(data, columns=headers)

        os.makedirs("data", exist_ok=True)

        filename = f"data/reporte_visitas_{time.strftime('%Y-%m-%d')}.xlsx"
        df.to_excel(filename, index=False)

        print("Excel guardado:", filename)

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
