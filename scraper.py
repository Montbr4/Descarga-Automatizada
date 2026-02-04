import os
import time
import shutil
import glob
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

RUC = "20504743307"
URL = f"https://visitas.servicios.gob.pe/consultas/index.php?ruc_enti={RUC}"
CARPETA_DATA = "data"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, CARPETA_DATA)

print("INICIANDO PROCESO")

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

    print("3. Esperando 1 minuto a que la página cargue los datos por sí misma")
    time.sleep(60) 
    print(" Tiempo de espera finalizado.")

    print("4. Buscando botón 'Excel'")
    boton_excel = None
    
    xpaths = [
        "//button[contains(., 'Excel')]", 
        "//button[contains(@class, 'buttons-excel')]",
        "//span[contains(., 'Excel')]/parent::button"
    ]
    
    for xpath in xpaths:
        try:
            btns = driver.find_elements(By.XPATH, xpath)
            for btn in btns:
                if btn.is_displayed():
                    boton_excel = btn
                    print(f"Botón encontrado: {xpath}")
                    break
            if boton_excel: break
        except:
            continue

    if boton_excel:
        archivos_antes = set(os.listdir(DOWNLOAD_DIR))
        
        driver.execute_script("arguments[0].click();", boton_excel)
        print("Clic realizado! Esperando archivo")
        
        tiempo = 0
        descarga_completa = False
        archivo_descargado = None
        
        while tiempo < 60:
            archivos_ahora = set(os.listdir(DOWNLOAD_DIR))
            nuevos = archivos_ahora - archivos_antes
            
            for f in nuevos:
                if not f.endswith(".crdownload") and not f.endswith(".tmp"):
                    archivo_descargado = f
                    descarga_completa = True
                    break
            
            if descarga_completa:
                break
                
            time.sleep(1)
            tiempo += 1
            if tiempo % 10 == 0: print(f"Esperando ({tiempo}s)")

        if descarga_completa and archivo_descargado:
            ruta_origen = os.path.join(DOWNLOAD_DIR, archivo_descargado)
            
            fecha_hoy = time.strftime("%Y-%m-%d")
            nombre_final = f"reporte_visitas_{fecha_hoy}.xls"
            ruta_destino = os.path.join(DOWNLOAD_DIR, nombre_final)
            
            if os.path.exists(ruta_destino):
                os.remove(ruta_destino)
            
            shutil.move(ruta_origen, ruta_destino)
            
            print(f"ÉXITO TOTAL: Archivo guardado en:")
            print(f"-{ruta_destino}")
            print(f"Tamaño: {os.path.getsize(ruta_destino)} bytes")
        else:
            print("ERROR: Se hizo clic pero no apareció el archivo en la carpeta.")
            print(f"Contenido carpeta: {os.listdir(DOWNLOAD_DIR)}")
            
    else:
        print("ERROR: No se encontró el botón 'Excel'.")
        driver.save_screenshot("debug_error_no_boton.png")

except Exception as e:
    print(f"ERROR CRÍTICO: {e}")

finally:
    if driver:
        driver.quit()
        print("FINALIZADO")
