import os
import time
import shutil
from datetime import datetime
import pytz
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

URL_INICIAL = "https://www.transparencia.gob.pe/reportes_directos/pte_transparencia_reg_visitas.aspx?id_entidad=11476&ver=&id_tema=500"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CARPETA_RAIZ = os.path.join(BASE_DIR, "data")

zona_peru = pytz.timezone('America/Lima')
fecha_hoy = datetime.now(zona_peru).strftime("%Y-%m-%d")

DOWNLOAD_DIR = os.path.join(CARPETA_RAIZ, fecha_hoy)

if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)
    print(f"CARPETA LISTA: {DOWNLOAD_DIR}")

chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)

driver = None

try:
    print("1. Abriendo navegador...")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 20)
    
    print(f"2. Cargando portal: {URL_INICIAL}")
    driver.get(URL_INICIAL)
    time.sleep(5)

    print("3. Buscando botón 'Ir a consultar'")
    xpath_boton = "//a[contains(., 'Ir a consultar') or contains(@class, 'btn')]"
    
    try:
        boton_ir = wait.until(EC.presence_of_element_located((By.XPATH, xpath_boton)))
        
        ventanas_antes = driver.window_handles
        
        print("Realizando SCROLL hacia el botón")
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", boton_ir)
        time.sleep(2)
        
        print("Ejecutando Clic")
        driver.execute_script("arguments[0].click();", boton_ir)
        
        print("Esperando apertura de nueva pestaña")
        wait.until(EC.number_of_windows_to_be(len(ventanas_antes) + 1))
        
        ventanas_nuevas = driver.window_handles
        driver.switch_to.window(ventanas_nuevas[-1])
        print(f"ÉXITO Cambiado a nueva pestaña: {driver.current_url}")

    except Exception as e:
        print(f"ERROR CRÍTICO al intentar entrar a la consulta: {e}")
        raise e 

    print("4. INICIANDO ESPERA DE 5 MINUTOS (Carga de datos)")
    time.sleep(300)
    print(" Tiempo finalizado.")

    print("5. Buscando botón Excel")
    xpath_excel = "//button[contains(., 'Excel')]"
    
    try:
        boton_excel = driver.find_element(By.XPATH, xpath_excel)
        
        archivos_antes = set(os.listdir(DOWNLOAD_DIR))
        
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", boton_excel)
        time.sleep(1)
        
        driver.execute_script("arguments[0].click();", boton_excel)
        print("Clic en EXCEL realizado. Esperando archivo")
        
        tiempo = 0
        descarga_exitosa = False
        nombre_archivo = ""
        
        while tiempo < 120:
            archivos_ahora = set(os.listdir(DOWNLOAD_DIR))
            nuevos = archivos_ahora - archivos_antes
            
            for f in nuevos:
                if not f.endswith(".crdownload") and not f.endswith(".tmp"):
                    nombre_archivo = f
                    descarga_exitosa = True
                    break
            
            if descarga_exitosa:
                break
            
            time.sleep(1)
            tiempo += 1
            if tiempo % 10 == 0: print(f"Descargando ({tiempo}s)")

        if descarga_exitosa:
            ruta_final = os.path.join(DOWNLOAD_DIR, nombre_archivo)
            print(f"OPERACIÓN EXITOSA")
            print(f"Archivo: {ruta_final}")
            print(f"Tamaño: {os.path.getsize(ruta_final)} bytes")
        else:
            print("ERROR: El botón se presionó pero no llegó el archivo.")
            
    except Exception as e:
        print(f"ERROR: No se encontró el botón Excel (¿Quizás la página no cargó?). Detalles: {e}")
        driver.save_screenshot("debug_error_final.png")

except Exception as e:
    print(f"ERROR DEL SISTEMA: {e}")

finally:
    if driver:
        driver.quit()
        print("FIN")
