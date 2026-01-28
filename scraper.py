import pandas as pd
from datetime import datetime
import pytz
import os
import time
import io
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

zona_peru = pytz.timezone('America/Lima')
ahora_peru = datetime.now(zona_peru)
fecha_hoy = ahora_peru.strftime("%Y-%m-%d")
fecha_busqueda = ahora_peru.strftime("%d/%m/%Y")

nombre_archivo = f"visitas_vivienda_{fecha_hoy}.csv"

print(f"--- INICIANDO PROCESO: {fecha_hoy} (Buscando data de: {fecha_busqueda}) ---")

if not os.path.exists(CARPETA_DATA):
    os.makedirs(CARPETA_DATA)

chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

driver = None

try:
    print("1. Abriendo navegador...")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 20)
    
    print(f"2. Entrando a: {URL}")
    driver.get(URL)
    
    print(f"3. Configurando rango de fechas a: {fecha_busqueda}")
    try:
        input_inicio = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[contains(@id, 'fec') and contains(@id, 'ini')] | //input[@name='fec_inicio']")))
        input_fin = driver.find_element(By.XPATH, "//input[contains(@id, 'fec') and contains(@id, 'fin')] | //input[@name='fec_fin']")
        
        driver.execute_script("arguments[0].value = '';", input_inicio)
        input_inicio.send_keys(fecha_busqueda)
        
        driver.execute_script("arguments[0].value = '';", input_fin)
        input_fin.send_keys(fecha_busqueda)
        print("   -> Fechas ingresadas correctamente.")
        
    except Exception as e:
        print(f"   -> ADVERTENCIA: No se pudieron setear las fechas automáticamente ({e}). Se usará la fecha por defecto.")

    print("4. Buscando botón y ejecutando búsqueda...")
    boton_encontrado = None
    xpaths_posibles = [
        "//button[contains(., 'Buscar')]", 
        "//input[@value='Buscar']",
        "//*[@id='btnBuscar']"
    ]
    
    for xpath in xpaths_posibles:
        try:
            btn = driver.find_element(By.XPATH, xpath)
            if btn.is_displayed():
                boton_encontrado = btn
                break
        except:
            continue
            
    if boton_encontrado:
        driver.execute_script("arguments[0].click();", boton_encontrado)
        print("   -> Clic realizado.")
    else:
        print("ERROR: No se encontró el botón de búsqueda.")

    print("5. Esperando carga de tabla (10s)...")
    time.sleep(10)

    print("6. Extrayendo datos...")
    html_cargado = driver.page_source
    

    try:
        tablas = pd.read_html(io.StringIO(html_cargado))
    except ValueError:
        tablas = []

    df_final = None

    for i, t in enumerate(tablas):
        cols = " ".join([str(c) for c in t.columns]).lower()
        if "visitante" in cols or "entidad" in cols or "motivo" in cols:
            print(f"   -> ¡Tabla de datos detectada en índice {i}!")
            df_final = t
            break
            
    if df_final is None and len(tablas) > 0:
        print("   -> Cabeceras no reconocidas, usando la tabla con más datos.")
        df_final = max(tablas, key=len)

    if df_final is not None and not df_final.empty:
        df_final = df_final.dropna(how='all')
        
        ruta_completa = os.path.join(CARPETA_DATA, nombre_archivo)
        df_final.to_csv(ruta_completa, index=False, encoding='utf-8-sig', sep=',')
        
        if len(df_final) > 1:
            print(f"ÉXITO: Datos guardados en {ruta_completa}")
            print(f"   Filas encontradas: {len(df_final)}")
        else:
            print(f"AVISO: Archivo guardado pero parece vacío (0 visitas). Ruta: {ruta_completa}")
            
    else:
        print("RESULTADO: No se encontraron tablas o datos para la fecha seleccionada.")
        with open(os.path.join(CARPETA_DATA, nombre_archivo), 'w') as f:
            f.write("Error,SinDatos\n")

except Exception as e:
    print(f"ERROR CRÍTICO: {e}")

finally:
    if driver:
        driver.quit()
        print("--- PROCESO FINALIZADO ---")
