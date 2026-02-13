import requests
import pandas as pd
from datetime import datetime
import pytz
import time

print("DESCARGA API CON SESIÓN REAL")

ruc = "20504743307"

url = "https://visitas.servicios.gob.pe/consultas/index.php"

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Referer": f"https://visitas.servicios.gob.pe/consultas/index.php?ruc_enti={ruc}",
    "X-Requested-With": "XMLHttpRequest",
    "Connection": "keep-alive"
}

params = {
    "ruc_enti": ruc
}

session = requests.Session()

print("Abriendo sesión…")

session.get(headers["Referer"], headers=headers, timeout=30)

time.sleep(2)

print("Consultando datos")

response = session.get(url, headers=headers, params=params, timeout=60)

if response.status_code != 200:
    raise Exception(f"Error HTTP {response.status_code}")

try:
    data = response.json()
except:
    raise Exception("Respuesta no es JSON — posible bloqueo")

if not data:
    raise Exception("API respondió vacío")

print("Registros recibidos:", len(data))

df = pd.DataFrame(data)

tz = pytz.timezone("America/Lima")
fecha = datetime.now(tz).strftime("%Y-%m-%d")

archivo = f"reporte_visitas_{fecha}.xlsx"

df.to_excel(archivo, index=False)

print("ARCHIVO GUARDADO:", archivo)
print("DESCARGA EXITOSA")
