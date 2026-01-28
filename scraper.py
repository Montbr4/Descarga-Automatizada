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

print(f"--- INICIANDO PROCESO: {fecha_hoy} (Buscando: {fecha_busqueda}) ---")

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
    
    time.sleep(3) 

    print(f"3. Configurando rango de fechas a: {fecha_busqueda}")
    
    selectores_inicio = ["fec_inicio", "txtFechaInicio", "inicio", "fechainicio"]
    selectores_fin = ["fec_fin", "txtFechaFin", "fin", "fechafin"]
    
    input_inicio = None
    input_fin = None

    for sel in selectores_inicio:
        try:
            input_inicio = driver.find_element(By.XPATH, f"//input[@id='{sel}' or @name='{sel}']")
            print(f"   -> Campo Inicio encontrado con selector: {sel}")
            break
        except:
            continue
            
    if not input_inicio:
        try:
            input_inicio = driver.find_element(By.XPATH, "(//input[contains(@class, 'fecha') or contains(@id, 'fec')])[1]")
            print("   -> Campo Inicio encontrado por posición genérica.")
        except:
            print("   -> ERROR: No se encontró campo Inicio.")

    for sel in selectores_fin:
        try:
            input_fin = driver.find_element(By.XPATH, f"//input[@id='{sel}' or @name='{sel}']")
            print(f"   -> Campo Fin encontrado con selector: {sel}")
            break
        except:
            continue
            
    if input_inicio:
        driver.execute_script(f"arguments[0].value = '{fecha_busqueda}';", input_inicio)
    
    if input_fin:
        driver.execute_script(f"arguments[0].value = '{fecha_busqueda}';", input_fin)
        
    print("   -> Fechas inyectadas vía JS.")

    print("4. Buscando botón y ejecutando búsqueda...")
    boton_encontrado = None
    
    xpaths_boton = [
        "//button[contains(translate(., 'BUSCAR', 'buscar'), 'buscar')]",
        "//input[@type='submit']",
        "//*[@id='btnBuscar']",
        "//button[@id='btn_buscar']"
    ]
    
    for xpath in xpaths_boton:
        try:
            btns = driver.find_elements(By.XPATH, xpath)
            for btn in btns:
                if btn.is_displayed():
                    boton_encontrado = btn
                    print(f"   -> Botón detectado: {xpath}")
                    break
            if boton_encontrado: break
        except:
            continue

    if boton_encontrado:
        driver.execute_script("arguments[0].click();", boton_encontrado)
        print("   -> Clic realizado.")
    else:
        print("ERROR CRÍTICO: No se encontró el botón de búsqueda.")
        with open("debug_error_html.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)

    print("5. Esperando resultados (10s)...")
    time.sleep(10)

    print("6. Extrayendo datos...")
    html_cargado = driver.page_source
    
    try:
        tablas = pd.read_html(io.StringIO(html_cargado))
    except Exception as e:
        print(f"   -> No se detectaron tablas HTML estándar: {e}")
        tablas = []

    df_final = None
    
    for i, t in enumerate(tablas):
        cols = " ".join([str(c) for c in t.columns]).lower()
        if "visitante" in cols or "entidad" in cols or "motivo" in cols:
            print(f"   -> ¡Tabla de datos detectada en índice {i}!")
            df_final = t
            break
    
    if df_final is None and len(tablas) > 0:
        print("   -> Usando la tabla más grande encontrada por defecto.")
        df_final = max(tablas, key=len)

    if df_final is not None and not df_final.empty:
        df_final = df_final.dropna(how='all')
        
        ruta_completa = os.path.join(CARPETA_DATA, nombre_archivo)
        df_final.to_csv(ruta_completa, index=False, encoding='utf-8-sig', sep=',')
        
        print(f"ÉXITO: Archivo guardado en {ruta_completa}")
        print(f"   Filas: {len(df_final)}")
        print(f"   Columnas: {list(df_final.columns)}")
    else:
        print("ADVERTENCIA: No se extrajeron datos. Puede que no haya visitas hoy o falló la búsqueda.")
        ruta_completa = os.path.join(CARPETA_DATA, nombre_archivo)
        with open(ruta_completa, 'w') as f:
            f.write("Error,SinDatos\n")

except Exception as e:
    print(f"ERROR CRÍTICO DEL PROCESO: {e}")

finally:
    if driver:
        driver.quit()
        print("--- PROCESO FINALIZADO ---")
