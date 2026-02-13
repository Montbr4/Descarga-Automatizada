import requests
import pandas as pd
import os
from datetime import datetime
import pytz

print("\nDESCARGA DIRECTA API")

zona = pytz.timezone("America/Lima")
fecha = datetime.now(zona).strftime("%Y-%m-%d")

BASE = os.getcwd()
DOWNLOAD = os.path.join(BASE, "data", fecha)
os.makedirs(DOWNLOAD, exist_ok=True)

archivo = os.path.join(DOWNLOAD, f"visitas_{fecha}.xlsx")

url = "https://visitas.servicios.gob.pe/consultas/listar"

params = {
    "ruc_enti": "20504743307",
    "length": "10000",
    "start": "0"
}

headers = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "application/json"
}

print("Consultando API")

r = requests.get(url, params=params, headers=headers)

if r.status_code != 200:
    raise Exception("API no respondió")

data = r.json()

rows = data.get("data", [])

if not rows:
    raise Exception("API devolvió tabla vacía")

print(f"Registros obtenidos: {len(rows)}")

df = pd.DataFrame(rows)
df.to_excel(archivo, index=False)

print("Archivo guardado:", archivo)
print("DESCARGA COMPLETA\n")
