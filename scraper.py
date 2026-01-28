import pandas as pd
from datetime import datetime
import pytz
import os
import time
import io
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

RUC = "20504743307"
URL = f"https://visitas.servicios.gob.pe/consultas/index.php?ruc_enti={RUC}"
CARPETA_DATA = "data"

zona_peru = pytz.timezone('America/Lima')
ahora_peru = datetime.now(zona_peru)
fecha_hoy = ahora_peru.strftime("%Y-%m-%d")
nombre_archivo = f"visitas_vivienda_{fecha_hoy}.csv"

print(f"Iniciando proceso: {fecha_hoy} (Hora Perú: {ahora_peru.strftime('%H:%M:%S')})")

if not os.path.exists(CARPETA_DATA):
    os.makedirs(CARPETA_DATA)

chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument(f"user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

try:
    print("Abriendo navegador")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    print(f"Entrando a: {URL}")
    driver.get(URL)

    print("Esperando a que la pagina cargue la tabla")
    time.sleep(20) 

    html_cargado = driver.page_source
    
    driver.quit()

    print("Procesando tabla...")
    tablas = pd.read_html(io.StringIO(html_cargado))

    if len(tablas) > 0:

        df = max(tablas, key=len) 
        
        df = df.dropna(how='all')
        
        if len(df) > 1:
            ruta_final = os.path.join(CARPETA_DATA, nombre_archivo)
            df.to_csv(ruta_final, index=False, encoding='utf-8-sig', sep=',')
            print(f"Guardado en: {ruta_final} ({len(df)} registros)")
        else:
            print(f"Tabla encontrada pero vacía. Fecha: {fecha_hoy}")
    else:
        print("No se encontraron tablas en el HTML renderizado.")

except Exception as e:
    print(f"ERROR CRÍTICO: {e}")
    if 'driver' in locals():
        driver.quit()
    exit(1)
