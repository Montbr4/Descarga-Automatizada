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

URL_INICIAL = "https://www.transparencia.gob.pe/reportes_directos/pte_transparencia_reg_visitas.aspx?id_entidad=11476&ver=&id_tema=500"
CARPETA_DATA = "data"

COLUMNAS_OBJETIVO = [
    "Fecha de Registro", "Fecha de Visita", "Entidad visitada", "Visitante", 
    "Documento del visitante", "Entidad del visitante", "Funcionario visitado", 
    "Hora Ingreso", "Hora Salida", "Motivo", "Lugar específico", "Observación"
]

zona_peru = pytz.timezone('America/Lima')
ahora_peru = datetime.now(zona_peru)
fecha_hoy = ahora_peru.strftime("%Y-%m-%d")
nombre_archivo = f"visitas_vivienda_{fecha_hoy}.csv"

print(f"Iniciando recorrido completo: {fecha_hoy} (Hora Perú: {ahora_peru.strftime('%H:%M:%S')})")

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
    print("Abriendo navegador...")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 20)

    print(f"Entrando a: {URL_INICIAL}")
    driver.get(URL_INICIAL)

    ventana_original = driver.current_window_handle

    print("Buscando botón 'Ir a consultar'...")
    boton_ir = wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Ir a consultar")))
    boton_ir.click()
    print("Click realizado en 'Ir a consultar'.")

    print("Esperando apertura de nueva pestaña")
    wait.until(EC.number_of_windows_to_be(2))

    for window_handle in driver.window_handles:
        if window_handle != ventana_original:
            driver.switch_to.window(window_handle)
            break
            
    print(f"Cambio de pestaña exitoso. URL actual: {driver.current_url}")

    print("Buscando el botón 'Buscar'")
    try:
        boton_buscar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Buscar')] | //input[@value='Buscar'] | //input[@type='submit']")))
        boton_buscar.click()
        print("Botón 'Buscar' presionado.")
    except Exception as e:
        print(f"Advertencia: No se pudo clickear 'Buscar' (quizás cargó automático): {e}")

    print("Esperando 15 segundos para carga de datos")
    time.sleep(15)

    html_cargado = driver.page_source
    print("Analizando tablas encontradas")
    
    tablas = pd.read_html(io.StringIO(html_cargado))
    df_final = None

    for i, tabla in enumerate(tablas):
        columnas_tabla = [str(c).strip() for c in tabla.columns]
        
        coincidencias = sum(1 for col in COLUMNAS_OBJETIVO if col in columnas_tabla)
        
        if coincidencias >= 3:
            print(f"¡Tabla objetivo encontrada en índice {i}! (Coinciden {coincidencias} columnas)")
            df_final = tabla
            break
    
    if df_final is not None:
        df_final = df_final.dropna(how='all')
        
        if len(df_final) > 0:
            ruta_final = os.path.join(CARPETA_DATA, nombre_archivo)
            df_final.to_csv(ruta_final, index=False, encoding='utf-8-sig', sep=',')
            print(f"Archivo guardado: {ruta_final}")
            print(f"Registros encontrados: {len(df_final)}")
            print("Muestra de datos:")
            print(df_final.head(2))
        else:
            print("La tabla correcta fue encontrada, pero está vacía (0 visitas registradas hoy).")
    else:
        print("No se encontró ninguna tabla que coincida con los encabezados del Ministerio.")
        if len(tablas) > 0:
            print(f"Columnas encontradas en la tabla 0 (ejemplo): {tablas[0].columns.tolist()}")

except Exception as e:
    print(f"ERROR CRÍTICO: {e}")
    if driver:
        print(f"URL donde ocurrió el error: {driver.current_url}")

finally:
    if driver:
        driver.quit()
