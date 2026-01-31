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
CARPETA_DATA = "data"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, CARPETA_DATA)

zona_peru = pytz.timezone('America/Lima')
ahora_peru = datetime.now(zona_peru)

fecha_archivo = ahora_peru.strftime("%Y-%m-%d") 
fecha_busqueda = ahora_peru.strftime("%d/%m/%Y") 

nombre_archivo_final = f"reporte_excel_{fecha_archivo}.xls"

print(f"INICIANDO PROCESO")
print(f"Fecha: {fecha_busqueda}")
print(f"Espera programada: 1 MINUTO")

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
    print("1. Abriendo navegador")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    print(f"2. Entrando a: {URL}")
    driver.get(URL)
    time.sleep(3)

    print(f"3. Escribiendo fecha: {fecha_busqueda}")
    script_fechas = f"""
    var fecha = '{fecha_busqueda}';
    var inputs = document.querySelectorAll("input[id*='fec'], input[id*='Fec'], input[name*='fec']");
    for(var i=0; i<inputs.length; i++) {{
        inputs[i].value = fecha;
    }}
    """
    driver.execute_script(script_fechas)

    print("4. Ejecutando búsqueda")
    try:
        btn_buscar = driver.find_element(By.XPATH, "//button[contains(translate(., 'BUSCAR', 'buscar'), 'buscar')]")
        driver.execute_script("arguments[0].click();", btn_buscar)
        print(" Clic en Buscar realizado.")
    except:
        print("No se encontró botón Buscar (puede ser automático).")

    print("5. ESPERANDO 1 MINUTO A QUE CARGUE LA TABLA")
    time.sleep(60)
    print(" Tiempo de espera finalizado")

    print("6. Buscando botón EXCEL")
    boton_excel = None
    
    xpaths_excel = [
        "//button[contains(., 'Excel')]",
        "//button[contains(@class, 'buttons-excel')]",
        "//span[contains(., 'Excel')]/parent::button",
        "//button[contains(translate(., 'EXCEL', 'excel'), 'excel')]"
    ]
    
    for xpath in xpaths_excel:
        try:
            btns = driver.find_elements(By.XPATH, xpath)
            for btn in btns:
                if btn.is_displayed():
                    boton_excel = btn
                    print(f"   -> Botón encontrado con: {xpath}")
                    break
            if boton_excel: break
        except:
            continue
            
    if boton_excel:
        archivos_antes = set(os.listdir(DOWNLOAD_DIR))
        
        driver.execute_script("arguments[0].click();", boton_excel)
        print("Clic en EXCEL realizado. Iniciando descarga")
        
        tiempo_espera = 0
        descarga_completa = False
        archivo_nuevo = None

        while tiempo_espera < 60:
            archivos_ahora = set(os.listdir(DOWNLOAD_DIR))
            nuevos = archivos_ahora - archivos_antes
            
            for f in nuevos:
                if not f.endswith(".crdownload") and not f.endswith(".tmp"):
                    archivo_nuevo = f
                    descarga_completa = True
                    break
            
            if descarga_completa:
                break
            
            time.sleep(1)
            tiempo_espera += 1
            
        if descarga_completa and archivo_nuevo:
            ruta_origen = os.path.join(DOWNLOAD_DIR, archivo_nuevo)
            ruta_destino = os.path.join(DOWNLOAD_DIR, nombre_archivo_final)
            
            if os.path.exists(ruta_destino):
                os.remove(ruta_destino)
                
            shutil.move(ruta_origen, ruta_destino)
            tamano = os.path.getsize(ruta_destino)
            print(f"ÉXITO: Archivo guardado: {ruta_destino}")
            print(f"   Tamaño: {tamano} bytes")
        else:
            print("ERROR: El botón se clickeó, pero no apareció ningún archivo nuevo.")
            
    else:
        print("ERROR: Sigue sin encontrarse el botón Excel. Guardando captura de pantalla")
        driver.save_screenshot("debug_error_excel.png")

except Exception as e:
    print(f"ERROR CRÍTICO: {e}")

finally:
    if driver:
        driver.quit()
        print("PROCESO FINALIZADO")
