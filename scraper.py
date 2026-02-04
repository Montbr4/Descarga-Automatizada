import os
import time
import shutil
from datetime import datetime
import pytz
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

RUC = "20504743307"
URL = f"https://visitas.servicios.gob.pe/consultas/index.php?ruc_enti={RUC}"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CARPETA_RAIZ = os.path.join(BASE_DIR, "data")

zona_peru = pytz.timezone('America/Lima')
fecha_hoy = datetime.now(zona_peru).strftime("%Y-%m-%d")

DOWNLOAD_DIR = os.path.join(CARPETA_RAIZ, fecha_hoy)

if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)
    print(f"CARPETA CREADA: {DOWNLOAD_DIR}")
else:
    print(f"USANDO CARPETA EXISTENTE: {DOWNLOAD_DIR}")

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
    
    print(f"2. Cargando página: {URL}")
    driver.get(URL)
    
    time.sleep(5)

    print("3. Ubicando botón Excel (sin presionar)")
    xpath_excel = "//button[contains(., 'Excel')]"
    
    boton_detectado = False
    try:
        btn_check = driver.find_element(By.XPATH, xpath_excel)
        if btn_check.is_displayed():
            print("Botón Excel VISUALIZADO correctamente")
            boton_detectado = True
        else:
            print(" Botón detectado en código pero no visible.")
    except:
        print("ADVERTENCIA: No se ve el botón Excel al inicio (se buscará de nuevo tras la espera).")

    print("4. INICIANDO ESPERA DE 5 MINUTOS (Cargando datos)")
    time.sleep(300)
    print("5 Minutos completados")

    print("5. Intentando presionar el botón Excel")
    
    try:
        btn_final = driver.find_element(By.XPATH, xpath_excel)
        
        archivos_antes = set(os.listdir(DOWNLOAD_DIR))
        
        driver.execute_script("arguments[0].click();", btn_final)
        print("CLIC EJECUTADO.")
        
        print("6. Esperando descarga de archivo")
        
        tiempo = 0
        descarga_exitosa = False
        nombre_archivo_descargado = ""
        
        while tiempo < 60:
            archivos_ahora = set(os.listdir(DOWNLOAD_DIR))
            nuevos = archivos_ahora - archivos_antes
            
            for f in nuevos:
                if not f.endswith(".crdownload") and not f.endswith(".tmp"):
                    nombre_archivo_descargado = f
                    descarga_exitosa = True
                    break
            
            if descarga_exitosa:
                break
                
            time.sleep(1)
            tiempo += 1
            if tiempo % 10 == 0: print(f"   ...descargando ({tiempo}s)")
            
        if descarga_exitosa:
            ruta_final = os.path.join(DOWNLOAD_DIR, nombre_archivo_descargado)
            tamano = os.path.getsize(ruta_final)
            
            print(f"ÉXITO")
            print(f"Archivo guardado en: {ruta_final}")
            print(f"Nombre original: {nombre_archivo_descargado}")
            print(f"Tamaño: {tamano} bytes")
        else:
            print("ERROR: Pasó el tiempo y no apareció ningún archivo nuevo en la carpeta.")
            print(f"Contenido de la carpeta: {os.listdir(DOWNLOAD_DIR)}")

    except Exception as e:
        print(f"ERROR AL INTENTAR CLICKEAR: {e}")
        driver.save_screenshot("debug_error_click.png")

except Exception as e:
    print(f"ERROR CRÍTICO DEL SISTEMA: {e}")

finally:
    if driver:
        driver.quit()
        print("PROCESO TERMINADO")
