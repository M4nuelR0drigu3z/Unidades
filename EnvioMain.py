import os
import sys
import logging
import requests
import pandas as pd
from dateutil import parser as dp
import pytz
from datetime import datetime, timedelta
from email.message import EmailMessage
import smtplib
from dotenv import load_dotenv
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.drawing.image import Image
import gspread
from gspread_dataframe import get_as_dataframe
from pathlib import Path
import re


# ----------------------------------------------------------------------
# CARGAR .env
env_path = Path(__file__).parent / ".env"
load_dotenv(dotenv_path=env_path)

# ----------------------------------------------------------------------
# Credenciales y configuración
PHONE_NUMBER_ID = os.getenv("WHATSAPP_PHONE_NUMBER_ID")
ACCESS_TOKEN = os.getenv("WHATSAPP_ACCESS_TOKEN")
DESTINOS = os.getenv("WHATSAPP_DESTINOS", "").split(",")

SAMSARA_API_TOKEN = os.getenv("SAMSARA_API_TOKEN")

SMTP_HOST = os.getenv("SMTP_HOST", "smtp.example.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")

TEMPLATE_NAME = os.getenv("TEMPLATE_NAME", "reporte")
LANG_CODE = os.getenv("LANG_CODE", "es_MX")
MX_TZ = "America/Mexico_City"

# ----------------------------------------------------------------------
# Logging
logging.basicConfig(
    filename="reporte_logs.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)

# ----------------------------------------------------------------------
# WhatsApp setup
BASE_URL = f"https://graph.facebook.com/v17.0/{PHONE_NUMBER_ID}"
HEADERS = {"Authorization": f"Bearer {ACCESS_TOKEN}"}


def subir_media(path: str) -> str:
    with open(path, "rb") as f:
        files = {
            "file": (
                os.path.basename(path),
                f,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ),
            "messaging_product": (None, "whatsapp"),
        }
        r = requests.post(f"{BASE_URL}/media", headers=HEADERS, files=files)
        r.raise_for_status()
        media_id = r.json()["id"]
        logging.info(f"Media subido, media_id: {media_id}")
        return media_id


def enviar_template(media_id: str, to: str, excel_path: str):
    payload = {
        "messaging_product": "whatsapp",
        "to": to,
        "type": "template",
        "template": {
            "name": TEMPLATE_NAME,
            "language": {"code": LANG_CODE},
            "components": [
                {
                    "type": "header",
                    "parameters": [
                        {
                            "type": "document",
                            "document": {
                                "id": media_id,
                                "filename": os.path.basename(excel_path),
                            },
                        }
                    ],
                }
            ],
        },
    }
    headers = {**HEADERS, "Content-Type": "application/json"}
    r = requests.post(f"{BASE_URL}/messages", json=payload, headers=headers)
    r.raise_for_status()
    msg_id = r.json()["messages"][0]["id"]
    logging.info(f"Template enviado a {to}, message ID: {msg_id}")


def obtener_datos_google_sheets(results, fecha_busqueda):
    base_dir = Path(__file__).parent
    gc = gspread.service_account(filename=base_dir / "credenciales.json")
    sh = gc.open_by_url(
        "https://docs.google.com/spreadsheets/d/1zbVe4Rk7aGaC_gyy0n2ik5VEWzn3w6XAyn91LNp2cMA/edit#gid=0"
    )
    print("[Sheets] Hojas:", [ws.title for ws in sh.worksheets()])

    # 1) Determinar hoja (mes/año)
    meses_es = {
        1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
        5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
        9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
    }
    mes = meses_es[fecha_busqueda.month]
    anio = str(fecha_busqueda.year)[-2:]  # "25"
    sheet_name = f"{mes} {anio}"
    worksheet = sh.worksheet(sheet_name)
    print(f"[Sheets] Hoja usada: {sheet_name}")

    # 2) Cargar a DataFrame
    df = get_as_dataframe(worksheet, evaluate_formulas=True).fillna("")

    # 3) Normalizar columna de fecha a tipo date (acepta ES y EN)
    col_fecha   = "FECHA DE INICIO"
    col_unidad  = "UNIDAD"
    col_roster  = "ROSTERING \nID"
    col_origen  = "Origen 0"
    col_destino = "Destino"
    col_placas  = "PLACAS"

    MES_ES_EN = {
        "ene": "jan", "feb": "feb", "mar": "mar", "abr": "apr",
        "may": "may", "jun": "jun", "jul": "jul", "ago": "aug",
        # variantes comunes
        "sep": "sep", "sept": "sep", "set": "sep",
        "oct": "oct", "nov": "nov", "dic": "dec",
    }

    def normaliza_fecha_cell(x):
        s = str(x).strip().lower()
        for es, en in MES_ES_EN.items():
            s = re.sub(rf"\b{es}\b", en, s)
        try:
            # dayfirst=True para '22-ago-25'
            return dp.parse(s, dayfirst=True).date()
        except Exception:
            return None

    # Si falta la columna, evita reventar y muestra pistas
    if col_fecha not in df.columns:
        print(f"[Sheets][WARN] No existe la columna '{col_fecha}'. Columnas disponibles: {list(df.columns)}")
        filas_fecha = df.iloc[0:0]  # vacío
    else:
        df["_fecha_norm"] = df[col_fecha].apply(normaliza_fecha_cell)
        target_date = fecha_busqueda.date()
        filas_fecha = df[df["_fecha_norm"] == target_date]

    print(f"[Sheets] Fecha objetivo: {fecha_busqueda.date()} | Coincidencias: {len(filas_fecha)}")
    if len(filas_fecha) == 0 and col_fecha in df.columns:
        print("[Sheets] Ejemplos FECHA DE INICIO:", df[col_fecha].astype(str).head(5).tolist())

    # 4) Enriquecer tus resultados
    for row in results:
        unidad = str(row.get("Unidad", "")).strip()

        if not filas_fecha.empty and col_unidad in df.columns:
            coincidencia = filas_fecha[filas_fecha[col_unidad].astype(str).str.strip() == unidad]
        else:
            coincidencia = pd.DataFrame()

        if not coincidencia.empty:
            fila = coincidencia.iloc[0]
            row["ID ROSTERING"] = fila.get(col_roster, "")
            row["ORIGEN"]       = fila.get(col_origen, "")
            row["DESTINO"]      = fila.get(col_destino, "")
            row["PLACAS"]       = fila.get(col_placas, "")

            geocerca  = row.get("Geocerca", "").strip()
            origen_0  = str(fila.get(col_origen, "")).strip()
            destino   = str(fila.get(col_destino, "")).strip()
            if geocerca:
                if geocerca == origen_0:
                    row["Estatus"] = "EN ORIGEN"
                elif geocerca == destino:
                    row["Estatus"] = "EN DESTINO"
        else:
            row["ID ROSTERING"] = ""
            row["ORIGEN"] = ""
            row["DESTINO"] = ""
            row["PLACAS"] = ""

    return results

# ----------------------------------------------------------------------
def main():
    logging.info("===> Inicio de ejecución")

    # validar token Samsara
    if not SAMSARA_API_TOKEN:
        logging.error("❌ Falta SAMSARA_API_TOKEN en variables de entorno")
        sys.exit(1)

    base_dir = Path(__file__).parent
    plantilla_path = base_dir / "PlantillaML.xlsx"

    # IDs predefinidas
    predefined_special = {
        "254792506",
        "254801835",
        "254802588",
        "254803338",
        "254859196",
        "94193861",
        "95243156",
        "95243200",
        "95243316",
        "95243513",
        "244349505",
        "245970120",
        "254794170",
        "254794716",
        "257477773",
    }

    # Obtener nuevos IDs desde Samsara…
    try:
        samsara_h = {
            "Accept": "application/json",
            "Authorization": f"Bearer {SAMSARA_API_TOKEN}",
        }
        tags = requests.get(
            "https://api.samsara.com/tags/4363967", headers=samsara_h, timeout=60
        )
        tags.raise_for_status()
        Geocercas_EC5 = {
            a["id"]
            for a in tags.json().get("data", {}).get("addresses", [])
            if a.get("id") and a.get("name")
        }

        logging.info(f"Geocercas EC5 obtenidas: {Geocercas_EC5}")
        
    except Exception:
        logging.exception("Error al obtener tags")
        sys.exit(1)

    # Obtener datos GPS…
    try:
        veh = requests.get(
            "https://api.samsara.com/fleet/vehicles/stats?types=gps",
            headers=samsara_h,
            params={"ParentTagIds": "4363967"},
            timeout=60,
        )
        veh.raise_for_status()
        vehicles = veh.json().get("data", [])
    except Exception:
        logging.exception("Error al obtener datos de vehículos")
        sys.exit(1)

    # Procesar registros
    results = []
    now_mx = datetime.now(pytz.timezone(MX_TZ))
    for u in vehicles:
        try:
            gps = u.get("gps", {})
            t = gps.get("time")
            geocerca_id = gps.get("address", {}).get("id")
            geocerca_name = gps.get("address", {}).get("name")
            if t:
                loc_time = dp.parse(t).astimezone(pytz.timezone(MX_TZ))
                if now_mx - loc_time > timedelta(hours=1):
                    continue
            # if gps.get("address", {}).get("id") in predefined_special:
            #    continue
            # else:
            #     gps.get("address", {}).get("id") in Geocercas_EC5
            if geocerca_id in predefined_special:
                continue
            geocerca_detectada = ""
            if geocerca_id in Geocercas_EC5 and geocerca_name:
                geocerca_detectada = geocerca_name.strip()
                
            speed = gps.get("speedMilesPerHour", 0)
            ecu = gps.get("isEcuSpeed", False)
            if speed == 0 and not ecu:
                continue
            status = "DETENIDO" if (speed == 0 and ecu) else "RUTA"
            location = gps.get("reverseGeo", {}).get("formattedLocation")

            lat_long = f"{gps.get('latitude')},{gps.get('longitude')}"

            results.append(
                {
                    "Unidad": u.get("name", "Sin nombre"),
                    "Ubicación": location,
                    "Estatus": status,
                    "Coordenadas": lat_long,
                    "Geocerca": geocerca_detectada,
                }
            )
        except Exception:
            logging.exception(f"Procesando unidad {u.get('name')}")

    results = obtener_datos_google_sheets(results, now_mx)
    ####DEBUG: Imprimir resultados antes de generar Excel##############
    print("Unidades procesadas:")
    for r in results:
        print(f"Unidad: {r['Unidad']}, Geocerca: '{r['Geocerca']}', Origen: '{r.get('ORIGEN','')}', Destino: '{r.get('DESTINO','')}', Estatus: {r['Estatus']}")
    ###################################

    # 6) Generar Excel desde plantilla
    wb = load_workbook(filename=plantilla_path)
    ws = wb.active

    # 6.1) Deshacer merges en filas >= 7
    start_row = 8
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= start_row:
            ws.unmerge_cells(str(mr))

    # 6.2) Fecha y hora
    ws["G3"] = now_mx.date().isoformat()
    ws["J3"] = now_mx.strftime("%H:%M:%S")

    # 6.3) Preparar estilos
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    green_fill = PatternFill(
        start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"
    )
    green_font = Font(color="006100")  # verde oscuro
    
    red_fill = PatternFill(
        start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    red_font = Font(color="9C0006")  # rojo oscuro

    blue_fill = PatternFill(
        start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
    )
    blue_font = Font(color="1F497D")  # azul oscuro
    
    reten_fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    reten_font = Font(color="7E350E")  # negro

    # 6.4) Volcar datos, fusionar Ubicación C–H, aplicar estilo y bordes
    center_al = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for i, row in enumerate(results, start=start_row):
        # ID ROSTERING
        cell_id = ws.cell(row=i, column=1, value=row["ID ROSTERING"])
        cell_id.border = border
        cell_id.alignment = center_al

        # ORIGEN
        cell_ori = ws.cell(row=i, column=2, value=row["ORIGEN"])
        cell_ori.border = border
        cell_ori.alignment = center_al

        # DESTINO
        cell_dest = ws.cell(row=i, column=3, value=row["DESTINO"])
        cell_dest.border = border
        cell_dest.alignment = center_al

        # ECO (Unidad)
        cell_eco = ws.cell(row=i, column=4, value=row["Unidad"])
        cell_eco.border = border
        cell_eco.alignment = center_al

        # PLACAS
        cell_pla = ws.cell(row=i, column=5, value=row["PLACAS"])
        cell_pla.border = border
        cell_pla.alignment = center_al

    # STATUS (Estatus)
        cell_s = ws.cell(row=i, column=6, value=row["Estatus"])
        cell_s.border = border
        cell_s.alignment = center_al
        
        if row.get("Geocerca", "").strip() == 'Reten Militar "El Desengaño" Sinaloa':
            cell_s.value = "RETEN"
            cell_s.fill = reten_fill
            cell_s.font = reten_font
        elif row["Estatus"] == "DETENIDO":
            cell_s.fill = red_fill
            cell_s.font = red_font
        elif row["Estatus"] in ("EN ORIGEN", "EN DESTINO"):
            cell_s.fill = blue_fill
            cell_s.font = blue_font
        else:
            cell_s.fill = green_fill
            cell_s.font = green_font
        

    # UBICACIÓN (merge de columnas 7 a 13)
        ws.merge_cells(start_row=i, start_column=7, end_row=i, end_column=13)
        for col in range(7, 14):  # 14 NO incluido, así hace 7–13
            cell_loc = ws.cell(
                row=i, column=col, value=row["Ubicación"] if col == 7 else None
            )
            cell_loc.border = border
            cell_loc.alignment = center_al

    # COORDENADAS
        cell_c = ws.cell(row=i, column=14, value=row["Coordenadas"])
        cell_c.border = border
        cell_c.alignment = center_al

    # 6.5) Conteo dinámico en H2
    last_row = start_row + len(results) - 1
    ws["L3"] = f"=COUNTA(D{start_row}:D{last_row})"

    # 6.6) Guardar archivo nuevo
    ws.column_dimensions["A"].width = 15  # ~110px (ajusta si es necesario)
    ws.row_dimensions[2].height = 15  # 45pt = ~60px
    ws.row_dimensions[3].height = 15
    ws.row_dimensions[4].height = 15

    # Crear la imagen y ajustar tamaño
    img_path = base_dir / "mercadolibre_logo.png"
    logo = Image(img_path)
    logo.width = 125  # Ancho total que ocupará (ajusta si el logo es más chico/largo)
    logo.height = 60  # Suma del alto de las 3 filas (45*3)
    ws.add_image(logo, "A2")

    # 6.6) Guardar archivo nuevo
    ts_str = now_mx.strftime("%Y-%m-%d_%H-%M-%S")
    nuevo_archivo = base_dir / f"Reporte de estatus de unidades {ts_str}.xlsx"
    wb.save(nuevo_archivo)
    logging.info(f"Excel generado: {nuevo_archivo}")

    # 7) Enviar por correo
    try:
        msg = EmailMessage()
        msg["From"] = SMTP_USER
        msg["To"] = "mrodriguez@bgcapitalgroup.mx"
        msg["Subject"] = "Reporte de estatus de unidades"
        msg.set_content("Hola, se adjunta el reporte.\n\nSaludos.")
        with open(nuevo_archivo, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=nuevo_archivo.name,
            )
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60) as s:
            s.starttls()
            s.login(SMTP_USER, SMTP_PASSWORD)
            s.send_message(msg)
        logging.info("Correo enviado")
    except Exception:
        logging.exception("Error enviando correo")
        sys.exit(1)

    # 8) Enviar por WhatsApp
    # for destino in DESTINOS:
    #     try:
    #         media_id = subir_media(str(nuevo_archivo))
    #         enviar_template(media_id, destino, str(nuevo_archivo))
    #     except Exception:
    #         logging.exception(f"Error enviando WhatsApp a {destino}")
    #         sys.exit(1)
    # 9) Eliminar archivo
    try:
        os.remove(str(nuevo_archivo))
        logging.info(f"Archivo eliminado: {nuevo_archivo.name}")
    except Exception:
        logging.exception(f"No se pudo eliminar: {nuevo_archivo.name}")

    logging.info("===> Ejecución finalizada correctamente")


if __name__ == "__main__":
    main()
