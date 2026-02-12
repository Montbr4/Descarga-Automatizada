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
    print(f"CARPETA DESTINO: {DOWNLOAD_DIR}")

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
    print("1. Abriendo navegador")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    print("Forzando Zona Horaria: America/Lima")
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
        
        print("   -> Navegando a Visitas (Misma pestaña)...")
        time.sleep(10)
        
    except Exception as e:
        print(f"   -> Error navegando: {e}")

    print("4. Analizando si la tabla tiene datos")
    
    try:
        cuerpo_pagina = driver.page_source
        if "No hay datos disponibles" in cuerpo_pagina or "0 to 0 of 0 Registros" in cuerpo_pagina:
            print("DETECTADO: La tabla está vacía (0 registros).")
            print("ACCIÓN: Forzando clic en 'Buscar' para cargar datos")
            
            btn_buscar = driver.find_element(By.XPATH, "//button[contains(translate(., 'BUSCAR', 'buscar'), 'buscar')]")
            driver.execute_script("arguments[0].click();", btn_buscar)
            print("Botón BUSCAR presionado. Esperando recarga")
            time.sleep(5)
        else:
            print("La tabla PARECE tener datos. No tocaré el botón Buscar.")
            
    except Exception as e:
        print(f"No pude verificar la tabla, asumo que cargará sola: {e}")

    print("5. INICIANDO ESPERA DE 5 MINUTOS (Carga total)")
    time.sleep(300)
    print("   -> Tiempo finalizado.")

    screenshot_file = os.path.join(DOWNLOAD_DIR, "evidencia_antes_descarga.png")
    driver.save_screenshot(screenshot_file)
    print(f"   -> Foto de evidencia guardada: {screenshot_file}")

    print("6. Buscando botón Excel")
    xpath_excel = "//button[contains(., 'Excel')]"
    
    try:
        boton_excel = driver.find_element(By.XPATH, xpath_excel)
        
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", boton_excel)
        time.sleep(1)
        
        archivos_antes = set(os.listdir(DOWNLOAD_DIR))
        
        driver.execute_script("arguments[0].click();", boton_excel)
        print("Clic realizado. Esperando archivo")
        
        tiempo = 0
        descarga_exitosa = False
        nombre_archivo = ""
        
        while tiempo < 120:
            archivos_ahora = set(os.listdir(DOWNLOAD_DIR))
            nuevos = archivos_ahora - archivos_antes
            
            for f in nuevos:
                if not f.endswith(".crdownload") and not f.endswith(".tmp") and "evidencia" not in f:
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
            size = os.path.getsize(ruta_final)
            print(f"ÉXITO")
            print(f"Archivo guardado: {ruta_final}")
            print(f"Tamaño: {size} bytes")
            
            if size < 1000:
                print("Posible Falla")
        else:
            print("ERROR: No se descargó nada (Timeout).")
            
    except Exception as e:
        print(f"ERROR buscando Excel: {e}")

except Exception as e:
    print(f"ERROR CRÍTICO: {e}")

finally:
    if driver:
        driver.quit()
        print("FIN")
