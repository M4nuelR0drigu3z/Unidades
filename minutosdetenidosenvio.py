import os
import sys
import logging
import requests
import pandas as pd
from dateutil import parser as dp
import pytz
from datetime import datetime, timedelta
from dotenv import load_dotenv
from pathlib import Path
import gspread
from gspread_dataframe import get_as_dataframe
import re
from Versions.detenciones_copy import enriquecer_minutos_detenido


# ----------------------------------------------------------------------
# CARGAR .env
env_path = Path(__file__).parent / ".env"
load_dotenv(dotenv_path=env_path)


# ----------------------------------------------------------------------
# Credenciales y configuracion
SAMSARA_API_TOKEN = os.getenv("SAMSARA_API_TOKEN")
GOOGLE_CHAT_WEBHOOK_URL = os.getenv("GOOGLE_CHAT_WEBHOOK_URL_PRUEBAS")
MX_TZ = "America/Mexico_City"


# ----------------------------------------------------------------------
# Logging
logging.basicConfig(
    filename="reporte_logs.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)


# ----------------------------------------------------------------------
# GOOGLE SHEETS
def obtener_datos_google_sheets(results, fecha_busqueda):
    """
    Se conserva solo para enriquecer con datos como ROSTERING y PLACAS.
    Ya no se usa para origen/destino.
    """

    base_dir = Path(__file__).parent

    gc = gspread.service_account(filename=base_dir / "credenciales.json")

    sh = gc.open_by_url(
        "https://docs.google.com/spreadsheets/d/1zbVe4Rk7aGaC_gyy0n2ik5VEWzn3w6XAyn91LNp2cMA/edit#gid=0"
    )

    print("[Sheets] Hojas:", [ws.title for ws in sh.worksheets()])

    meses_es = {
        1: "ENERO",
        2: "FEBRERO",
        3: "MARZO",
        4: "ABRIL",
        5: "MAYO",
        6: "JUNIO",
        7: "JULIO",
        8: "AGOSTO",
        9: "SEPTIEMBRE",
        10: "OCTUBRE",
        11: "NOVIEMBRE",
        12: "DICIEMBRE",
    }

    mes = meses_es[fecha_busqueda.month]
    anio = str(fecha_busqueda.year)[-2:]
    sheet_name = f"{mes} {anio}"

    try:
        worksheet = sh.worksheet(sheet_name)
    except Exception:
        logging.exception(f"No se encontro la hoja {sheet_name}")
        print(f"[Sheets][ERROR] No se encontro la hoja: {sheet_name}")
        print("[Sheets] Se continuara sin enriquecer datos.")
        return results

    print(f"[Sheets] Hoja usada: {sheet_name}")

    df = get_as_dataframe(worksheet, evaluate_formulas=True).fillna("")

    col_fecha = "FECHA DE INICIO"
    col_unidad = "UNIDAD"
    col_roster = "ROSTERING \nID"
    col_placas = "PLACAS"

    MES_ES_EN = {
        "ene": "jan",
        "feb": "feb",
        "mar": "mar",
        "abr": "apr",
        "may": "may",
        "jun": "jun",
        "jul": "jul",
        "ago": "aug",
        "sep": "sep",
        "sept": "sep",
        "set": "sep",
        "oct": "oct",
        "nov": "nov",
        "dic": "dec",
    }

    def normaliza_fecha_cell(x):
        s = str(x).strip().lower()

        if not s:
            return None

        for es, en in MES_ES_EN.items():
            s = re.sub(rf"\b{es}\b", en, s)

        try:
            return dp.parse(s, dayfirst=True).date()
        except Exception:
            return None

    if col_fecha not in df.columns:
        print(
            f"[Sheets][WARN] No existe la columna '{col_fecha}'. "
            f"Columnas disponibles: {list(df.columns)}"
        )
        filas_fecha = df.iloc[0:0]
    else:
        df["_fecha_norm"] = df[col_fecha].apply(normaliza_fecha_cell)
        target_date = fecha_busqueda.date()
        filas_fecha = df[df["_fecha_norm"] == target_date]

    print(
        f"[Sheets] Fecha objetivo: {fecha_busqueda.date()} | "
        f"Coincidencias: {len(filas_fecha)}"
    )

    if len(filas_fecha) == 0 and col_fecha in df.columns:
        print(
            "[Sheets] Ejemplos FECHA DE INICIO:",
            df[col_fecha].astype(str).head(5).tolist(),
        )

    for row in results:
        unidad = str(row.get("Unidad", "")).strip()

        if not filas_fecha.empty and col_unidad in df.columns:
            coincidencia = filas_fecha[
                filas_fecha[col_unidad].astype(str).str.strip() == unidad
            ]
        else:
            coincidencia = pd.DataFrame()

        if not coincidencia.empty:
            fila = coincidencia.iloc[0]

            row["ID ROSTERING"] = fila.get(col_roster, "")
            row["PLACAS"] = fila.get(col_placas, "")
        else:
            row["ID ROSTERING"] = ""
            row["PLACAS"] = ""

        row["ORIGEN"] = ""
        row["DESTINO"] = ""

    return results


# ----------------------------------------------------------------------
# FORMATO DEL REPORTE
def icono_estatus(row):
    estatus = str(row.get("Estatus", "")).strip()

    if estatus == "DETENIDO":
        return "⛔", "DETENIDO"
    elif estatus == "TRAFICO LENTO":
        return "🚦", "TRAFICO LENTO"
    elif estatus == "RUTA":
        return "✅", "RUTA"
    elif estatus == "RETEN":
        return "🚧", "RETEN"
    else:
        return "❓", estatus or "SIN ESTATUS"


def construir_reporte_google(results, now_mx):
    total = len(results)

    detenidos = 0
    ruta = 0
    reten = 0
    trafico_lento = 0

    filas = []

    for row in results:
        _, estatus_final = icono_estatus(row)

        if estatus_final == "DETENIDO":
            detenidos += 1
        elif estatus_final == "RUTA":
            ruta += 1
        elif estatus_final == "TRAFICO LENTO":
            trafico_lento += 1
        elif estatus_final == "RETEN":
            reten += 1

        unidad = str(row.get("Unidad", "") or "")
        ubicacion = str(row.get("Ubicación", "") or "")
        coordenadas = str(row.get("Coordenadas", "") or "")
        tiempo_detenido = str(row.get("Tiempo Detenido") or "")
        tiempo_trafico = str(row.get("Tiempo Trafico") or "")
        motor = str(row.get("Motor") or "")

        if estatus_final == "DETENIDO":
            tiempo_reporte = tiempo_detenido
        elif estatus_final == "TRAFICO LENTO":
            tiempo_reporte = tiempo_trafico
        else:
            tiempo_reporte = ""

        filas.append({
            "UNIDAD": unidad,
            "ESTATUS": estatus_final,
            "TIEMPO": tiempo_reporte,
            "MOTOR": motor,
            "UBICACION": ubicacion,
            "COORDENADAS": coordenadas,
        })

    orden_estatus = {
        "DETENIDO": 0,
        "TRAFICO LENTO": 1,
        "RETEN": 2,
        "RUTA": 3,
    }

    filas.sort(
        key=lambda x: (
            orden_estatus.get(x["ESTATUS"], 99),
            x["UNIDAD"]
        )
    )

    lineas = []

    lineas.append("🚛 *REPORTE DE ESTATUS DE UNIDADES*")
    lineas.append("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    lineas.append(f"📅 *Fecha:* {now_mx.strftime('%Y-%m-%d')}")
    lineas.append(f"🕒 *Hora:* {now_mx.strftime('%H:%M:%S')}")
    lineas.append(f"📦 *Total unidades:* {total}")
    lineas.append("")
    lineas.append("📊 *Resumen*")
    lineas.append(f"✅ En ruta: {ruta}")
    lineas.append(f"⛔ Detenidos: {detenidos}")
    lineas.append(f"🚦 Tráfico lento: {trafico_lento}")
    lineas.append(f"🚧 Retén: {reten}")
    lineas.append("")
    lineas.append("*Detalle:*")

    if not filas:
        lineas.append("```")
        lineas.append("No se encontraron unidades para reportar.")
        lineas.append("```")
        return "\n".join(lineas)

    ancho_unidad = 15
    ancho_estatus = 14
    ancho_tiempo = 12
    ancho_motor = 8
    ancho_ubicacion = 65
    ancho_coordenadas = 24

    def cortar(texto, ancho):
        texto = str(texto or "").replace("\n", " ").replace("\r", " ").strip()
        if len(texto) > ancho:
            return texto[: ancho - 3] + "..."
        return texto

    def celda(texto, ancho):
        return cortar(texto, ancho).ljust(ancho)

    lineas.append("```")

    encabezado = (
        celda("UNIDAD", ancho_unidad)
        + " | "
        + celda("ESTATUS", ancho_estatus)
        + " | "
        + celda("TIEMPO", ancho_tiempo)
        + " | "
        + celda("MOTOR", ancho_motor)
        + " | "
        + celda("UBICACION", ancho_ubicacion)
        + " | "
        + celda("COORDENADAS", ancho_coordenadas)
    )

    separador = (
        "-" * ancho_unidad
        + "-+-"
        + "-" * ancho_estatus
        + "-+-"
        + "-" * ancho_tiempo
        + "-+-"
        + "-" * ancho_motor
        + "-+-"
        + "-" * ancho_ubicacion
        + "-+-"
        + "-" * ancho_coordenadas
    )

    lineas.append(encabezado)
    lineas.append(separador)

    for fila in filas:
        linea = (
            celda(fila["UNIDAD"], ancho_unidad)
            + " | "
            + celda(fila["ESTATUS"], ancho_estatus)
            + " | "
            + celda(fila["TIEMPO"], ancho_tiempo)
            + " | "
            + celda(fila["MOTOR"], ancho_motor)
            + " | "
            + celda(fila["UBICACION"], ancho_ubicacion)
            + " | "
            + celda(fila["COORDENADAS"], ancho_coordenadas)
        )
        lineas.append(linea)

    lineas.append("```")
    lineas.append("")
    lineas.append("✅ *Reporte generado automáticamente*")

    return "\n".join(lineas)


# ----------------------------------------------------------------------
# ENVIO A GOOGLE CHAT
def enviar_google_chat(texto):
    if not GOOGLE_CHAT_WEBHOOK_URL:
        logging.error("Falta GOOGLE_CHAT_WEBHOOK_URL en variables de entorno")
        print("❌ Falta GOOGLE_CHAT_WEBHOOK_URL en el archivo .env")
        sys.exit(1)

    payload = {
        "text": texto
    }

    response = requests.post(
        GOOGLE_CHAT_WEBHOOK_URL,
        json=payload,
        timeout=30,
    )

    print("Google Chat status:", response.status_code)
    print("Google Chat respuesta:", response.text)

    response.raise_for_status()

    logging.info("Mensaje enviado a Google Chat correctamente")


# ----------------------------------------------------------------------
# MAIN
def main():
    logging.info("===> Inicio de ejecucion")
    print("===> Inicio de ejecucion")

    if not SAMSARA_API_TOKEN:
        logging.error("Falta SAMSARA_API_TOKEN en variables de entorno")
        print("❌ Falta SAMSARA_API_TOKEN en el archivo .env")
        sys.exit(1)

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

    samsara_h = {
        "Accept": "application/json",
        "Authorization": f"Bearer {SAMSARA_API_TOKEN}",
    }

    # ------------------------------------------------------------------
    # Obtener geocercas desde Samsara
    try:
        print("[Samsara] Obteniendo geocercas...")

        tags = requests.get(
            "https://api.samsara.com/tags/4363967",
            headers=samsara_h,
            timeout=60,
        )

        print("[Samsara] Tags status:", tags.status_code)

        tags.raise_for_status()

        Geocercas_EC5 = {
            str(a["id"]): a["name"].strip()
            for a in tags.json().get("data", {}).get("addresses", [])
            if a.get("id") and a.get("name")
        }

        logging.info(f"Geocercas EC5 obtenidas: {Geocercas_EC5}")
        print(f"[Samsara] Geocercas obtenidas: {len(Geocercas_EC5)}")

        print("[Samsara] Lista de geocercas EC5:")
        for gid, nombre in Geocercas_EC5.items():
            print(f"  {gid} | {nombre}")

    except Exception as e:
        logging.exception("Error al obtener tags")
        print("❌ Error al obtener tags de Samsara:", e)
        sys.exit(1)

    # ------------------------------------------------------------------
    # Obtener datos GPS actuales
    try:
        print("[Samsara] Obteniendo datos GPS...")

        veh = requests.get(
            "https://api.samsara.com/fleet/vehicles/stats?types=gps,engineStates,ecuSpeedMph",
            headers=samsara_h,
            params={"ParentTagIds": "4363967"},
            timeout=60,
        )

        print("[Samsara] Vehiculos status:", veh.status_code)

        veh.raise_for_status()

        vehicles = veh.json().get("data", [])

        print(f"[Samsara] Vehiculos recibidos: {len(vehicles)}")
        logging.info(f"Vehiculos recibidos: {len(vehicles)}")

    except Exception as e:
        logging.exception("Error al obtener datos de vehiculos")
        print("❌ Error al obtener datos de vehiculos:", e)
        sys.exit(1)

    # ------------------------------------------------------------------
    # Procesar registros
    results = []
    now_mx = datetime.now(pytz.timezone(MX_TZ))

    omitidas_patio = 0
    omitidas_especiales = 0
    omitidas_sin_ecu = 0
    omitidas_gps_viejo = 0

    for u in vehicles:
        try:
            gps = u.get("gps", {})
            t = gps.get("time")

            geocerca_id = gps.get("address", {}).get("id")
            geocerca_name = gps.get("address", {}).get("name")

            if t:
                loc_time = dp.parse(t).astimezone(pytz.timezone(MX_TZ))

                if now_mx - loc_time > timedelta(hours=1):
                    omitidas_gps_viejo += 1
                    print(
                        f"[Filtro GPS viejo] Unidad {u.get('name')} omitida | "
                        f"GPS time={loc_time}"
                    )
                    continue

            geocerca_id = str(geocerca_id) if geocerca_id else ""

            if geocerca_id in predefined_special:
                omitidas_especiales += 1
                print(
                    f"[Filtro Geocerca Especial] Unidad {u.get('name')} omitida | "
                    f"geocerca_id={geocerca_id}"
                )
                continue

            geocerca_detectada = ""

            if geocerca_id in Geocercas_EC5:
                geocerca_detectada = Geocercas_EC5[geocerca_id]
            elif geocerca_name:
                geocerca_detectada = geocerca_name.strip()

            # Si está dentro de patio/geocerca, no se reporta.
            if geocerca_detectada:
                omitidas_patio += 1
                print(
                    f"[Filtro Patio] Unidad {u.get('name')} omitida por geocerca: "
                    f"{geocerca_detectada}"
                )
                continue

            speed = gps.get("speedMilesPerHour", 0)
            ecu = gps.get("isEcuSpeed", False)

            if speed == 0 and not ecu:
                omitidas_sin_ecu += 1
                print(
                    f"[Filtro Sin ECU] Unidad {u.get('name')} omitida | "
                    f"speed={speed} | ecu={ecu}"
                )
                continue

            status = "DETENIDO" if speed == 0 and ecu else "RUTA"

            location = gps.get("reverseGeo", {}).get("formattedLocation", "")

            lat = gps.get("latitude", "")
            lon = gps.get("longitude", "")
            lat_long = f"{lat},{lon}"

            print(
                f"[DEBUG ESTATUS] Unidad={u.get('name')} | "
                f"id={u.get('id')} | "
                f"speed={speed} | "
                f"ecu={ecu} | "
                f"geocerca_id={geocerca_id} | "
                f"status_inicial={status}"
            )

            results.append(
                {
                    "Unidad": u.get("name", "Sin nombre"),
                    "SamsaraVehicleId": u.get("id"),
                    "GpsActual": gps,
                    "Ubicación": location,
                    "Estatus": status,
                    "Coordenadas": lat_long,
                    "Geocerca": "",
                    "Minutos Detenido": None,
                    "Tiempo Detenido": None,
                    "Detenido Desde": None,
                    "Ventana Detenido": "",
                    "Minutos Trafico": None,
                    "Tiempo Trafico": None,
                    "Trafico Desde": None,
                    "Motor": "",
                    "EcuSpeedActual": None,
                }
            )

        except Exception:
            logging.exception(f"Procesando unidad {u.get('name')}")

    print(f"[Proceso] Unidades para reporte antes de analizar detenidos: {len(results)}")
    print(f"[Proceso] Omitidas por patio/geocerca: {omitidas_patio}")
    print(f"[Proceso] Omitidas por geocerca especial: {omitidas_especiales}")
    print(f"[Proceso] Omitidas por GPS viejo: {omitidas_gps_viejo}")
    print(f"[Proceso] Omitidas por speed 0 sin ECU: {omitidas_sin_ecu}")

    logging.info(f"Unidades para reporte antes de analizar detenidos: {len(results)}")

    # ------------------------------------------------------------------
    # Enriquecer con Google Sheets
    try:
        results = obtener_datos_google_sheets(results, now_mx)
    except Exception as e:
        logging.exception("Error obteniendo datos de Google Sheets")
        print("⚠️ Error obteniendo datos de Google Sheets:", e)
        print("⚠️ Se continuara con datos basicos de Samsara.")

    # ------------------------------------------------------------------
    # Enriquecer detenidos con historial GPS, engineStates y ecuSpeedMph
    try:
        results = enriquecer_minutos_detenido(
            results=results,
            token=SAMSARA_API_TOKEN,
            now_mx=now_mx,
        )
    except Exception as e:
        logging.exception("Error enriqueciendo minutos detenido")
        print("⚠️ Error enriqueciendo minutos detenido:", e)
        print("⚠️ Se continuara sin informacion de minutos detenido.")

    print("")
    print("[DEBUG DESPUES DETENCIONES]")
    for r in results:
        print(
            f"Unidad={r.get('Unidad')} | "
            f"Estatus={r.get('Estatus')} | "
            f"TiempoDet={r.get('Tiempo Detenido')} | "
            f"TiempoTrafico={r.get('Tiempo Trafico')} | "
            f"Motor={r.get('Motor')} | "
            f"ECU={r.get('EcuSpeedActual')} | "
            f"Ventana={r.get('Ventana Detenido')} | "
            f"Geocerca={r.get('Geocerca')}"
        )

    # ------------------------------------------------------------------
    # Construir y enviar mensaje a Google Chat
    try:
        mensaje_chat = construir_reporte_google(results, now_mx)

        print("")
        print("Mensaje para Google Chat:")
        print(mensaje_chat)

        enviar_google_chat(mensaje_chat)

    except Exception as e:
        logging.exception("Error construyendo o enviando mensaje a Google Chat")
        print("❌ Error construyendo o enviando mensaje a Google Chat:", e)
        sys.exit(1)

    logging.info("===> Ejecucion finalizada correctamente")
    print("===> Ejecucion finalizada correctamente")


# ----------------------------------------------------------------------
if __name__ == "__main__":
    main()