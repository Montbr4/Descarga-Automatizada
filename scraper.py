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

print(f"Iniciando descarga COMPLETA (todas las columnas): {fecha_hoy} (Hora: {ahora_peru.strftime('%H:%M:%S')})")

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
    print("Abriendo navegador")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    print(f"Entrando a: {URL}")
    driver.get(URL)

    time.sleep(5)
    
    print("Buscando botón 'Buscar'")
    boton_encontrado = None

    posibles_xpaths = [
        "//button[contains(., 'Buscar')]", 
        "//input[@value='Buscar']",
        "//input[@type='submit']"
    ]
    
    for xpath in posibles_xpaths:
        try:
            elems = driver.find_elements(By.XPATH, xpath)
            for btn in elems:
                if btn.is_displayed():
                    boton_encontrado = btn
                    break
            if boton_encontrado: break
        except:
            continue
            
    if boton_encontrado:
        print("Botón encontrado. Forzando click...")
        driver.execute_script("arguments[0].click();", boton_encontrado)
    else:
        print("Aviso: No se detectó botón de búsqueda. Esperando carga automática")

    print("Esperando 15 segundos para carga de datos")
    time.sleep(15)

    html_cargado = driver.page_source
    
    tablas = pd.read_html(io.StringIO(html_cargado))
    df_final = None
    
    for i, t in enumerate(tablas):
        cols_texto = " ".join([str(c) for c in t.columns])
        
        if "Visitante" in cols_texto or "Entidad" in cols_texto or "Hora Ingreso" in cols_texto:
            print(f"Tabla principal detectada en índice {i}")
            df_final = t 
            break
            
    if df_final is None and len(tablas) > 0:
        print("No se detectaron cabeceras clave. Usando la tabla más grande encontrada.")
        df_final = max(tablas, key=len)

    if df_final is not None:
        df_final = df_final.dropna(how='all')
        
        ruta_final = os.path.join(CARPETA_DATA, nombre_archivo)
        
        df_final.to_csv(ruta_final, index=False, encoding='utf-8-sig', sep=',')
        
        if len(df_final) > 1:
            print(f"¡ÉXITO! CSV generado en: {ruta_final}")
            print(f"Filas: {len(df_final)}")
            print(f"Columnas detectadas: {len(df_final.columns)}")
            print("Nombres de columnas descargadas:", list(df_final.columns))
        else:
            print("Tabla detectada pero vacía (0 registros hoy). Se guardó solo la cabecera.")
            driver.save_screenshot("debug_tabla_vacia.png")
            
    else:
        print("ERROR: No se encontró ninguna tabla en la página.")
        driver.save_screenshot("debug_error.png")

except Exception as e:
    print(f"ERROR CRÍTICO: {e}")
    if driver:
        driver.save_screenshot("debug_crash.png")

finally:
    if driver:
        driver.quit()
