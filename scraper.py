import os
import time
import shutil
import glob
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
CARPETA_DATA = os.path.join(BASE_DIR, "data")

zona_peru = pytz.timezone('America/Lima')
ahora_peru = datetime.now(zona_peru)
fecha_hoy_str = ahora_peru.strftime("%Y-%m-%d")

DOWNLOAD_DIR = os.path.join(CARPETA_DATA, fecha_hoy_str)

if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)
    print(f"CARPETA DESTINO CREADA: {DOWNLOAD_DIR}")

NOMBRE_FINAL = f"reporte_visitas_{fecha_hoy_str}.xlsx"
RUTA_FINAL = os.path.join(DOWNLOAD_DIR, NOMBRE_FINAL)

chrome_options = Options()
chrome_options.add_argument("--headless") 
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
chrome_options.add_argument("--disable-popup-blocking")

prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)

driver = None

try:
    print("1. Iniciando navegador")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    print("Configurando zona horaria: America/Lima")
    driver.execute_cdp_cmd("Emulation.setTimezoneOverride", {"timezoneId": "America/Lima"})

    wait = WebDriverWait(driver, 30)
    
    print(f"2. Cargando portal Transparencia: {URL_INICIAL}")
    driver.get(URL_INICIAL)
    time.sleep(5)

    print("3. Buscando botón 'Ir a consultar'")
    xpath_boton = "//a[contains(., 'Ir a consultar') or contains(@class, 'btn')]"
    
    try:
        boton_ir = wait.until(EC.presence_of_element_located((By.XPATH, xpath_boton)))
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", boton_ir)
        time.sleep(2)
        
        driver.execute_script("arguments[0].setAttribute('target', '_self');", boton_ir)
        driver.execute_script("arguments[0].click();", boton_ir)
        
        print("Navegando a Visitas")
        time.sleep(10)
        
    except Exception as e:
        print(f"Error navegando: {e}")

    print("4. Verificando datos")
    
    page_source = driver.page_source
    tabla_vacia = "0 to 0 of 0 Registros" in page_source or "No hay datos disponibles" in page_source
    
    if tabla_vacia:
        print("Tabla vacía detectada.")
        print("Forzando clic en 'BUSCAR' para refrescar")
        try:
            btn_buscar = driver.find_element(By.XPATH, "//button[contains(translate(., 'BUSCAR', 'buscar'), 'buscar')]")
            driver.execute_script("arguments[0].click();", btn_buscar)
            print("   -> Botón Buscar presionado. Esperando 10s...")
            time.sleep(10)
        except Exception as e:
            print(f"No se pudo clickear Buscar: {e}")
    else:
        print("La tabla parece tener datos. Continuamos.")

    print("5. Esperando 30 segundos para carga final")
    time.sleep(30)

    print("6. Buscando botón Excel...")
    xpaths_excel = [
        "//button[contains(@class, 'buttons-excel')]", 
        "//button[contains(., 'Excel')]",
        "//span[contains(., 'Excel')]/parent::button"
    ]
    
    boton_excel = None
    for xpath in xpaths_excel:
        try:
            btn = driver.find_element(By.XPATH, xpath)
            if btn.is_displayed():
                boton_excel = btn
                break
        except:
            continue
            
    if boton_excel:
        archivos_antes = set(os.listdir(DOWNLOAD_DIR))
        
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", boton_excel)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", boton_excel)
        print("Clic en Excel realizado. Esperando descarga")
        
        tiempo = 0
        descarga_exitosa = False
        nombre_descargado = ""
        
        while tiempo < 60:
            archivos_ahora = set(os.listdir(DOWNLOAD_DIR))
            nuevos = archivos_ahora - archivos_antes
            
            for f in nuevos:
                if not f.endswith(".crdownload") and not f.endswith(".tmp"):
                    nombre_descargado = f
                    descarga_exitosa = True
                    break
            
            if descarga_exitosa:
                break
            
            time.sleep(1)
            tiempo += 1
            if tiempo % 10 == 0: print(f"Descargando ({tiempo}s)")
            
        if descarga_exitosa:
            ruta_origen = os.path.join(DOWNLOAD_DIR, nombre_descargado)
            
            if os.path.exists(RUTA_FINAL):
                os.remove(RUTA_FINAL)
            
            shutil.move(ruta_origen, RUTA_FINAL)
            
            tamano = os.path.getsize(RUTA_FINAL)
            print(f"ÉXITO")
            print(f"Archivo guardado: {RUTA_FINAL}")
            print(f"Tamaño: {tamano} bytes")
            
            if tamano < 2000:
                print("Archivo pequeño.")
                driver.save_screenshot(os.path.join(DOWNLOAD_DIR, "debug_evidencia.png"))
                
        else:
            print("ERROR: Tiempo agotado, no se descargó el archivo.")
            driver.save_screenshot(os.path.join(DOWNLOAD_DIR, "debug_error_no_descarga.png"))
            
    else:
        print("ERROR: No se encontró el botón Excel.")
        driver.save_screenshot(os.path.join(DOWNLOAD_DIR, "debug_no_boton.png"))

except Exception as e:
    print(f"ERROR CRÍTICO: {e}")

finally:
    if driver:
        driver.quit()
        print("FIN DEL PROCESO")
