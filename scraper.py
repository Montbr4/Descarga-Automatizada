import requests
import pandas as pd
import os
from datetime import datetime

print("INICIO")

BASE_URL = "https://visitas.servicios.gob.pe/consultas/"
API_URL = BASE_URL + "dataBusqueda.php"

RUC = "20504743307"


session = requests.Session()

headers_page = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "text/html"
}

print("Obteniendo sesión inicial")

session.get(
    f"{BASE_URL}index.php?ruc_enti={RUC}",
    headers=headers_page
)

payload = {
    "draw": 1,
    "columns[0][data]": "fecha_registro",
    "columns[1][data]": "fecha_visita",
    "start": 0,
    "length": 1000,
    "search[value]": "",
    "ruc_enti": RUC
}

headers_api = {
    "User-Agent": "Mozilla/5.0",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": f"{BASE_URL}index.php?ruc_enti={RUC}"
}

print("Consultando API")

response = session.post(
    API_URL,
    data=payload,
    headers=headers_api
)

data = response.json()

rows = data.get("data", [])

print(f"Filas obtenidas: {len(rows)}")

if not rows:
    print("No se obtuvieron datos — posible bloqueo")
    exit()

df = pd.DataFrame(rows)

fecha = datetime.now().strftime("%Y-%m-%d")

folder = f"data/{fecha}"
os.makedirs(folder, exist_ok=True)

archivo = f"{folder}/visitas.xlsx"

df.to_excel(archivo, index=False)

print("Excel guardado en:", archivo)
print("FIN")
