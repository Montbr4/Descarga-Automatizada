import requests
import pandas as pd
from datetime import datetime
import os

RUC = "20504743307"
URL = f"https://visitas.servicios.gob.pe/consultas/index.php?ruc_enti={RUC}"
CARPETA_DATA = "data"
fecha_hoy = datetime.now().strftime("%Y-%m-%d")
nombre_archivo = f"visitas_vivienda_{fecha_hoy}.csv"

print(f"Descarga del Día: {fecha_hoy}")

if not os.path.exists(CARPETA_DATA):
    os.makedirs(CARPETA_DATA)

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

try:
    print(f"Consultando URL...")
    response = requests.get(URL, headers=headers, timeout=45)
    response.raise_for_status()
    response.encoding = response.apparent_encoding

    tablas = pd.read_html(response.text)

    if len(tablas) > 0:
        df = tablas[0]
        df = df.dropna(how='all')
        
        if len(df) > 1:
            ruta_final = os.path.join(CARPETA_DATA, nombre_archivo)
            df.to_csv(ruta_final, index=False, encoding='utf-8-sig', sep=',')
            print(f"Guardado en {ruta_final} ({len(df)} registros)")
        else:
            print("Tabla vacía (sin visitas hoy o feriado).")
    else:
        print("No se encontraron tablas.")

except Exception as e:
    print(f"ERROR: {e}")
