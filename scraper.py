import os
import time
import glob
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

RUC = "20504743307"
URL = f"https://visitas.servicios.gob.pe/consultas/index.php?ruc_enti={RUC}"
CARPETA_DATA = "data"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, CARPETA_DATA)

zona_peru = pytz.timezone('America/Lima')
ahora_peru = datetime.now(zona_peru)
fecha_hoy_str = ahora_peru.strftime("%Y-%m-%d")
fecha_busqueda = ahora_peru.strftime("%d/%m/%Y")

nombre_archivo_final = f"reporte_excel_{fecha_hoy_str}.xls"

print(f"--- INICIANDO DESCARGA EXCEL: {fecha_hoy_str} (Data del: {fecha_busqueda}) ---")

if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

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
    "safebrowsing.enabled": True,
    "profile.default_content_setting_values.automatic_downloads": 1
}
chrome_options.add_experimental_option("prefs", prefs)

driver = None

try:
    print("1. Abriendo navegador")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 20)
    
    print(f"2. Entrando a: {URL}")
    driver.get(URL)
    time.sleep(3)

    print(f"3. Configurando fechas a: {fecha_busqueda}")
    
    try:
        inputs_fecha = driver.find_elements(By.XPATH, "//input[contains(@id, 'fec') or contains(@id, 'Fec') or contains(@name, 'fec')]")
        
        script_set_date = f"arguments[0].value = '{fecha_busqueda}';"
        
        count = 0
        for inp in inputs_fecha:
            if inp.is_displayed() or "ini" in inp.get_attribute("id") or "fin" in inp.get_attribute("id"):
                driver.execute_script(script_set_date, inp)
                count += 1
        
        if count > 0:
            print(f"Fechas inyectadas en {count} campos.")
        else:
            print("ADVERTENCIA: No se encontraron inputs de fecha obvios.")
            
    except Exception as e:
        print(f"Error al poner fechas: {e}")

    print("4. Ejecutando búsqueda previa")
    try:
        btn_buscar = driver.find_element(By.XPATH, "//button[contains(., 'Buscar') or contains(., 'BUSCAR')]")
        driver.execute_script("arguments[0].click();", btn_buscar)
        time.sleep(5)
    except:
        print("No se encontró botón Buscar o no fue necesario.")

    print("5. Buscando botón EXCEL")
    boton_excel = None
    
    xpaths_excel = [
        "//button[contains(., 'EXCEL')]",
        "//a[contains(., 'EXCEL')]",
        "//input[@value='EXCEL']",
        "//button[contains(@class, 'excel')]",
        "//img[contains(@src, 'excel')]/parent::a"
    ]
    
    for xpath in xpaths_excel:
        try:
            btns = driver.find_elements(By.XPATH, xpath)
            for btn in btns:
                if btn.is_displayed():
                    boton_excel = btn
                    print(f"Botón Excel encontrado: {xpath}")
                    break
            if boton_excel: break
        except:
            continue
            
    if boton_excel:
        archivos_antes = set(os.listdir(DOWNLOAD_DIR))
        
        driver.execute_script("arguments[0].click();", boton_excel)
        print("Esperando descarga")
        
        tiempo_espera = 0
        descarga_completa = False
        archivo_nuevo = None
        
        while tiempo_espera < 30:
            archivos_ahora = set(os.listdir(DOWNLOAD_DIR))
            nuevos = archivos_ahora - archivos_antes
            
            for f in nuevos:
                if not f.endswith(".crdownload") and (f.endswith(".xls") or f.endswith(".xlsx")):
                    archivo_nuevo = f
                    descarga_completa = True
                    break
            
            if descarga_completa:
                break
                
            time.sleep(1)
            tiempo_espera += 1
            print(f"Esperando archivo ({tiempo_espera}s)")
            
        if descarga_completa and archivo_nuevo:
            ruta_origen = os.path.join(DOWNLOAD_DIR, archivo_nuevo)
            ruta_destino = os.path.join(DOWNLOAD_DIR, nombre_archivo_final)
            
            shutil.move(ruta_origen, ruta_destino)
            print(f"ÉXITO: Archivo descargado y renombrado a: {ruta_destino}")
            
            if os.path.getsize(ruta_destino) > 0:
                print(f"Tamaño: {os.path.getsize(ruta_destino)} bytes")
            else:
                print("ADVERTENCIA: El archivo descargado está vacío (0 bytes).")
                
        else:
            print("ERROR: Tiempo de espera agotado. No se detectó archivo nuevo.")
            print(f"   Archivos en carpeta: {os.listdir(DOWNLOAD_DIR)}")
            
    else:
        print("ERROR: No se encontró el botón EXCEL en la página.")
        driver.save_screenshot("debug_no_excel.png")

except Exception as e:
    print(f"ERROR CRÍTICO: {e}")

finally:
    if driver:
        driver.quit()
        print("PROCESO FINALIZADO")
