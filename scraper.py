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
from webdriver_manager.chrome import ChromeDriverManager

RUC = "20504743307"
URL = f"https://visitas.servicios.gob.pe/consultas/index.php?ruc_enti={RUC}"
CARPETA_DATA = "data"

zona_peru = pytz.timezone('America/Lima')
ahora_peru = datetime.now(zona_peru)
fecha_hoy = ahora_peru.strftime("%Y-%m-%d")
nombre_archivo = f"visitas_vivienda_{fecha_hoy}.csv"

print(f"--- INICIANDO PROCESO: {fecha_hoy} ---")

if not os.path.exists(CARPETA_DATA):
    os.makedirs(CARPETA_DATA)

chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=2560,1440")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

driver = None

try:
    print("1. Abriendo navegador...")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    print(f"2. Entrando a: {URL}")
    driver.get(URL)
    
    time.sleep(5)
    
    print("3. Intentando localizar el botón de búsqueda")
    boton_encontrado = None
    
    xpaths_posibles = [
        "//button[contains(., 'Buscar')]", 
        "//input[@value='Buscar']",
        "//input[@type='submit']",
        "//*[@id='btnBuscar']"
    ]
    
    for xpath in xpaths_posibles:
        try:
            elementos = driver.find_elements(By.XPATH, xpath)
            for btn in elementos:
                if btn.is_displayed():
                    boton_encontrado = btn
                    print(f"   -> Botón encontrado con: {xpath}")
                    break
            if boton_encontrado: break
        except:
            continue
            
    if boton_encontrado:
        print("4. Ejecutando CLIC en el botón")
        driver.execute_script("arguments[0].click();", boton_encontrado)
    else:
        print("No se encontró botón explícito. Esperando carga automática")

    print("5. Esperando 15 segundos a que la tabla cargue datos")
    time.sleep(15)

    print("6. Leyendo tabla...")
    html_cargado = driver.page_source
    
    tablas = pd.read_html(io.StringIO(html_cargado))
    df_final = None
    
    for i, t in enumerate(tablas):
        columnas_txt = " ".join([str(c) for c in t.columns])
        if "Visitante" in columnas_txt or "Entidad" in columnas_txt:
            print(f"   -> ¡Tabla de datos detectada en índice {i}!")
            df_final = t
            break
            
    if df_final is None and len(tablas) > 0:
        print("   -> No se detectaron cabeceras clave. Usando la tabla más grande.")
        df_final = max(tablas, key=len)

    if df_final is not None:
        df_final = df_final.dropna(how='all')
        
        ruta_completa = os.path.join(CARPETA_DATA, nombre_archivo)
        
        df_final.to_csv(ruta_completa, index=False, encoding='utf-8-sig', sep=',')
        
        if len(df_final) > 1:
            print(f"ÉXITO: Archivo guardado en {ruta_completa}")
            print(f"   Registros: {len(df_final)}")
            print(f"   Columnas: {len(df_final.columns)}")
        else:
            print("La tabla se descargó pero está vacía (0 visitas hoy).")
            driver.save_screenshot("debug_tabla_vacia.png")
            
    else:
        print("ERROR: No se encontró ninguna tabla en la página.")
        driver.save_screenshot("debug_error_no_tabla.png")

except Exception as e:
    print(f"ERROR CRÍTICO DEL SISTEMA: {e}")
    if driver:
        driver.save_screenshot("debug_crash.png")

finally:
    if driver:
        driver.quit()
