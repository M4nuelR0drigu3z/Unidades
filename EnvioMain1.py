import os
import sys
import logging
import re
from pathlib import Path
from datetime import datetime, timedelta

import requests
import pandas as pd
import pytz
import gspread

from dateutil import parser as dp
from dotenv import load_dotenv
from gspread_dataframe import get_as_dataframe

from detenciones import enriquecer_minutos_detenido


# ----------------------------------------------------------------------
# CARGAR .env
# ----------------------------------------------------------------------
env_path = Path(__file__).parent / ".env"
load_dotenv(dotenv_path=env_path)


# ----------------------------------------------------------------------
# CREDENCIALES Y CONFIGURACIÓN
# ----------------------------------------------------------------------
SAMSARA_API_TOKEN = os.getenv("SAMSARA_API_TOKEN")
GOOGLE_CHAT_WEBHOOK_URL = os.getenv("GOOGLE_CHAT_WEBHOOK_URL_PRUEBAS")

MX_TZ = "America/Mexico_City"
PARENT_TAG_ID = "4363967"

SAMSARA_BASE_URL = "https://api.samsara.com"


# ----------------------------------------------------------------------
# LOGGING
# ----------------------------------------------------------------------
logging.basicConfig(
    filename="reporte_logs.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)


# ----------------------------------------------------------------------
# OBTENER VEHÍCULOS DE SAMSARA CON PAGINACIÓN
# ----------------------------------------------------------------------
def obtener_vehiculos_stats(headers, parent_tag_id):
    """
    Obtiene todos los vehículos asociados al parent tag.

    Maneja la paginación utilizando endCursor y hasNextPage.
    """

    url = f"{SAMSARA_BASE_URL}/fleet/vehicles/stats"

    vehicles = []
    after = None
    numero_pagina = 1

    while True:
        params = {
            "types": "gps,engineStates,ecuSpeedMph",
            "parentTagIds": parent_tag_id,
        }

        if after:
            params["after"] = after

        print(
            f"[Samsara] Consultando página {numero_pagina} "
            f"de vehículos..."
        )

        response = requests.get(
            url,
            headers=headers,
            params=params,
            timeout=60,
        )

        print(
            f"[Samsara] Vehículos página {numero_pagina} "
            f"status: {response.status_code}"
        )

        response.raise_for_status()

        respuesta_json = response.json()
        datos_pagina = respuesta_json.get("data", [])

        vehicles.extend(datos_pagina)

        print(
            f"[Samsara] Página {numero_pagina}: "
            f"{len(datos_pagina)} vehículos"
        )

        pagination = respuesta_json.get("pagination") or {}
        has_next_page = pagination.get("hasNextPage", False)
        after = pagination.get("endCursor")

        if not has_next_page:
            break

        if not after:
            print(
                "[Samsara][WARN] hasNextPage=True, pero no existe "
                "endCursor. Se detiene la paginación."
            )
            break

        numero_pagina += 1

    return vehicles


# ----------------------------------------------------------------------
# GOOGLE SHEETS
# ----------------------------------------------------------------------
def obtener_datos_google_sheets(results, fecha_busqueda):
    """
    Enriquece los resultados con ROSTERING y PLACAS.

    Ya no se usa Google Sheets para obtener origen y destino.
    """

    base_dir = Path(__file__).parent

    gc = gspread.service_account(
        filename=base_dir / "credenciales.json"
    )

    sh = gc.open_by_url(
        "https://docs.google.com/spreadsheets/"
        "d/1zbVe4Rk7aGaC_gyy0n2ik5VEWzn3w6XAyn91LNp2cMA/"
        "edit#gid=0"
    )

    print(
        "[Sheets] Hojas:",
        [ws.title for ws in sh.worksheets()]
    )

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
        logging.exception(
            f"No se encontró la hoja {sheet_name}"
        )

        print(
            f"[Sheets][ERROR] No se encontró la hoja: "
            f"{sheet_name}"
        )
        print(
            "[Sheets] Se continuará sin enriquecer datos."
        )

        return results

    print(f"[Sheets] Hoja usada: {sheet_name}")

    df = get_as_dataframe(
        worksheet,
        evaluate_formulas=True
    ).fillna("")

    col_fecha = "FECHA DE INICIO"
    col_unidad = "UNIDAD"
    col_roster = "ROSTERING \nID"
    col_placas = "PLACAS"

    meses_es_en = {
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

    def normaliza_fecha_cell(valor):
        texto = str(valor).strip().lower()

        if not texto:
            return None

        for abreviatura_es, abreviatura_en in meses_es_en.items():
            texto = re.sub(
                rf"\b{abreviatura_es}\b",
                abreviatura_en,
                texto
            )

        try:
            return dp.parse(
                texto,
                dayfirst=True
            ).date()

        except Exception:
            return None

    if col_fecha not in df.columns:
        print(
            f"[Sheets][WARN] No existe la columna "
            f"'{col_fecha}'. Columnas disponibles: "
            f"{list(df.columns)}"
        )

        filas_fecha = df.iloc[0:0]

    else:
        df["_fecha_norm"] = df[col_fecha].apply(
            normaliza_fecha_cell
        )

        target_date = fecha_busqueda.date()

        filas_fecha = df[
            df["_fecha_norm"] == target_date
        ]

    print(
        f"[Sheets] Fecha objetivo: "
        f"{fecha_busqueda.date()} | "
        f"Coincidencias: {len(filas_fecha)}"
    )

    if len(filas_fecha) == 0 and col_fecha in df.columns:
        print(
            "[Sheets] Ejemplos FECHA DE INICIO:",
            df[col_fecha].astype(str).head(5).tolist(),
        )

    for row in results:
        unidad = str(
            row.get("Unidad", "")
        ).strip()

        if (
            not filas_fecha.empty
            and col_unidad in df.columns
        ):
            coincidencia = filas_fecha[
                filas_fecha[col_unidad]
                .astype(str)
                .str.strip() == unidad
            ]

        else:
            coincidencia = pd.DataFrame()

        if not coincidencia.empty:
            fila = coincidencia.iloc[0]

            row["ID ROSTERING"] = fila.get(
                col_roster,
                ""
            )

            row["PLACAS"] = fila.get(
                col_placas,
                ""
            )

        else:
            row["ID ROSTERING"] = ""
            row["PLACAS"] = ""

        row["ORIGEN"] = ""
        row["DESTINO"] = ""

    return results


# ----------------------------------------------------------------------
# FORMATO DEL REPORTE
# ----------------------------------------------------------------------
def icono_estatus(row):
    estatus = str(
        row.get("Estatus", "")
    ).strip()

    if estatus == "DETENIDO":
        return "⛔", "DETENIDO"

    if estatus == "TRAFICO LENTO":
        return "🚦", "TRAFICO LENTO"

    if estatus == "RUTA":
        return "✅", "RUTA"

    if estatus == "RETEN":
        return "🚧", "RETEN"

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

        unidad = str(
            row.get("Unidad", "") or ""
        )

        ubicacion = str(
            row.get("Ubicación", "") or ""
        )

        coordenadas = str(
            row.get("Coordenadas", "") or ""
        )

        tiempo_detenido = str(
            row.get("Tiempo Detenido") or ""
        )

        tiempo_trafico = str(
            row.get("Tiempo Trafico") or ""
        )

        motor = str(
            row.get("Motor") or ""
        )

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
        key=lambda item: (
            orden_estatus.get(
                item["ESTATUS"],
                99
            ),
            item["UNIDAD"]
        )
    )

    lineas = []

    lineas.append(
        "🚛 *REPORTE DE ESTATUS DE UNIDADES*"
    )
    lineas.append(
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    )
    lineas.append(
        f"📅 *Fecha:* {now_mx.strftime('%Y-%m-%d')}"
    )
    lineas.append(
        f"🕒 *Hora:* {now_mx.strftime('%H:%M:%S')}"
    )
    lineas.append(
        f"📦 *Total unidades:* {total}"
    )
    lineas.append("")
    lineas.append("📊 *Resumen*")
    lineas.append(f"✅ En ruta: {ruta}")
    lineas.append(f"⛔ Detenidos: {detenidos}")
    lineas.append(
        f"🚦 Tráfico lento: {trafico_lento}"
    )
    lineas.append(f"🚧 Retén: {reten}")
    lineas.append("")
    lineas.append("*Detalle:*")

    if not filas:
        lineas.append("```")
        lineas.append(
            "No se encontraron unidades para reportar."
        )
        lineas.append("```")

        return "\n".join(lineas)

    ancho_unidad = 15
    ancho_estatus = 14
    ancho_tiempo = 12
    ancho_motor = 8
    ancho_ubicacion = 65
    ancho_coordenadas = 24

    def cortar(texto, ancho):
        texto = (
            str(texto or "")
            .replace("\n", " ")
            .replace("\r", " ")
            .strip()
        )

        if len(texto) > ancho:
            return texto[: ancho - 3] + "..."

        return texto

    def celda(texto, ancho):
        return cortar(
            texto,
            ancho
        ).ljust(ancho)

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
            celda(
                fila["UNIDAD"],
                ancho_unidad
            )
            + " | "
            + celda(
                fila["ESTATUS"],
                ancho_estatus
            )
            + " | "
            + celda(
                fila["TIEMPO"],
                ancho_tiempo
            )
            + " | "
            + celda(
                fila["MOTOR"],
                ancho_motor
            )
            + " | "
            + celda(
                fila["UBICACION"],
                ancho_ubicacion
            )
            + " | "
            + celda(
                fila["COORDENADAS"],
                ancho_coordenadas
            )
        )

        lineas.append(linea)

    lineas.append("```")
    lineas.append("")
    lineas.append(
        "✅ *Reporte generado automáticamente*"
    )

    return "\n".join(lineas)


# ----------------------------------------------------------------------
# ENVÍO A GOOGLE CHAT
# ----------------------------------------------------------------------
def enviar_google_chat(texto):
    if not GOOGLE_CHAT_WEBHOOK_URL:
        logging.error(
            "Falta GOOGLE_CHAT_WEBHOOK_URL "
            "en variables de entorno"
        )

        print(
            "❌ Falta GOOGLE_CHAT_WEBHOOK_URL "
            "en el archivo .env"
        )

        sys.exit(1)

    payload = {
        "text": texto
    }

    response = requests.post(
        GOOGLE_CHAT_WEBHOOK_URL,
        json=payload,
        timeout=30,
    )

    print(
        "Google Chat status:",
        response.status_code
    )

    print(
        "Google Chat respuesta:",
        response.text
    )

    response.raise_for_status()

    logging.info(
        "Mensaje enviado a Google Chat correctamente"
    )


# ----------------------------------------------------------------------
# EXPORTAR DEPURACIÓN
# ----------------------------------------------------------------------
def exportar_unidades_excluidas(unidades_excluidas):
    if not unidades_excluidas:
        print(
            "[Depuración] No hay unidades excluidas "
            "para exportar."
        )
        return

    ruta_archivo = (
        Path(__file__).parent
        / "unidades_excluidas.csv"
    )

    df_excluidas = pd.DataFrame(
        unidades_excluidas
    )

    df_excluidas = df_excluidas.sort_values(
        by=["Motivo", "Unidad"],
        ascending=[True, True]
    )

    df_excluidas.to_csv(
        ruta_archivo,
        index=False,
        encoding="utf-8-sig"
    )

    print(
        f"[Depuración] Archivo generado: "
        f"{ruta_archivo}"
    )


def exportar_unidades_con_error(unidades_con_error):
    if not unidades_con_error:
        return

    ruta_archivo = (
        Path(__file__).parent
        / "unidades_con_error.csv"
    )

    df_errores = pd.DataFrame(
        unidades_con_error
    )

    df_errores.to_csv(
        ruta_archivo,
        index=False,
        encoding="utf-8-sig"
    )

    print(
        f"[Depuración] Archivo de errores generado: "
        f"{ruta_archivo}"
    )


def exportar_comparacion_samsara(unidades_comparacion):
    """Exporta la fotografia completa usada por EnvioMain1."""
    ruta_archivo = (
        Path(__file__).parent
        / "comparacion_telemetria_samsara.csv"
    )

    columnas = [
        "ReporteGeneradoMexico",
        "ParentTagId",
        "Unidad",
        "SamsaraVehicleId",
        "Decision",
        "Motivo",
        "Detalle",
        "EstatusInicial",
        "EstatusFinal",
        "MotorFinal",
        "TiempoDetenidoFinal",
        "GpsTimeUTC",
        "GpsTimeMexico",
        "AntiguedadMinutos",
        "Latitude",
        "Longitude",
        "Coordenadas",
        "ReverseGeo",
        "GeocercaId",
        "Geocerca",
        "SpeedMilesPerHour",
        "SpeedUsadaPorFiltro",
        "IsEcuSpeed",
    ]

    df_comparacion = pd.DataFrame(
        unidades_comparacion,
        columns=columnas
    )

    if not df_comparacion.empty:
        df_comparacion = df_comparacion.sort_values(
            by=["Decision", "Unidad"],
            ascending=[True, True]
        )

    df_comparacion.to_csv(
        ruta_archivo,
        index=False,
        encoding="utf-8-sig"
    )

    print(
        f"[Comparacion] Archivo generado: "
        f"{ruta_archivo} | "
        f"Filas={len(df_comparacion)}"
    )

    logging.info(
        f"Archivo de comparacion generado: "
        f"{ruta_archivo} | "
        f"Filas={len(df_comparacion)}"
    )


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
def main():
    logging.info("===> Inicio de ejecución")
    print("===> Inicio de ejecución")

    if not SAMSARA_API_TOKEN:
        logging.error(
            "Falta SAMSARA_API_TOKEN "
            "en variables de entorno"
        )

        print(
            "❌ Falta SAMSARA_API_TOKEN "
            "en el archivo .env"
        )

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
        "Authorization": (
            f"Bearer {SAMSARA_API_TOKEN}"
        ),
    }

    # ------------------------------------------------------------------
    # OBTENER GEOCERCAS DEL TAG
    # ------------------------------------------------------------------
    try:
        print(
            "[Samsara] Obteniendo geocercas..."
        )

        tags_response = requests.get(
            f"{SAMSARA_BASE_URL}/tags/{PARENT_TAG_ID}",
            headers=samsara_h,
            timeout=60,
        )

        print(
            "[Samsara] Tag status:",
            tags_response.status_code
        )

        tags_response.raise_for_status()

        tag_data = (
            tags_response.json().get("data") or {}
        )

        Geocercas_EC5 = {
            str(address["id"]): str(
                address["name"]
            ).strip()
            for address in tag_data.get(
                "addresses",
                []
            )
            if address.get("id")
            and address.get("name")
        }

        logging.info(
            f"Geocercas EC5 obtenidas: "
            f"{Geocercas_EC5}"
        )

        print(
            f"[Samsara] Geocercas obtenidas: "
            f"{len(Geocercas_EC5)}"
        )

        print(
            "[Samsara] Lista de geocercas EC5:"
        )

        for geocerca_id, nombre in Geocercas_EC5.items():
            print(
                f"  {geocerca_id} | {nombre}"
            )

    except Exception as error:
        logging.exception(
            "Error al obtener tags"
        )

        print(
            "❌ Error al obtener tags de Samsara:",
            error
        )

        sys.exit(1)

    # ------------------------------------------------------------------
    # OBTENER DATOS GPS ACTUALES
    # ------------------------------------------------------------------
    try:
        print(
            "[Samsara] Obteniendo datos GPS..."
        )

        vehicles = obtener_vehiculos_stats(
            headers=samsara_h,
            parent_tag_id=PARENT_TAG_ID,
        )

        print(
            f"[Samsara] Vehículos totales recibidos: "
            f"{len(vehicles)}"
        )

        logging.info(
            f"Vehículos recibidos: {len(vehicles)}"
        )

    except Exception as error:
        logging.exception(
            "Error al obtener datos de vehículos"
        )

        print(
            "❌ Error al obtener datos de vehículos:",
            error
        )

        sys.exit(1)

    # ------------------------------------------------------------------
    # PROCESAR REGISTROS
    # ------------------------------------------------------------------
    results = []

    unidades_excluidas = []
    unidades_con_error = []
    unidades_comparacion = []

    now_mx = datetime.now(
        pytz.timezone(MX_TZ)
    )

    omitidas_patio = 0
    omitidas_especiales = 0
    omitidas_sin_ecu = 0
    omitidas_gps_viejo = 0

    for u in vehicles:
        unidad_id = str(
            u.get("id", "")
        )

        unidad_nombre = str(
            u.get("name", "Sin nombre")
        )

        registro_comparacion = {
            "ReporteGeneradoMexico": now_mx.strftime(
                "%Y-%m-%d %H:%M:%S %Z"
            ),
            "ParentTagId": PARENT_TAG_ID,
            "Unidad": unidad_nombre,
            "SamsaraVehicleId": unidad_id,
            "Decision": "",
            "Motivo": "",
            "Detalle": "",
            "EstatusInicial": "",
            "EstatusFinal": "",
            "MotorFinal": "",
            "TiempoDetenidoFinal": "",
            "GpsTimeUTC": "",
            "GpsTimeMexico": "",
            "AntiguedadMinutos": None,
            "Latitude": "",
            "Longitude": "",
            "Coordenadas": "",
            "ReverseGeo": "",
            "GeocercaId": "",
            "Geocerca": "",
            "SpeedMilesPerHour": None,
            "SpeedUsadaPorFiltro": None,
            "IsEcuSpeed": None,
        }

        def registrar_comparacion(
            decision,
            motivo="",
            detalle="",
            estatus=""
        ):
            registro = registro_comparacion.copy()
            registro.update({
                "Decision": decision,
                "Motivo": motivo,
                "Detalle": detalle,
                "EstatusInicial": estatus,
            })
            unidades_comparacion.append(registro)

        try:
            gps = u.get("gps") or {}
            gps_time = gps.get("time")

            address = gps.get("address") or {}

            geocerca_id = str(
                address.get("id") or ""
            )

            geocerca_name = str(
                address.get("name") or ""
            ).strip()

            speed = gps.get(
                "speedMilesPerHour"
            )

            ecu = gps.get(
                "isEcuSpeed",
                False
            )

            reverse_geo = (
                gps.get("reverseGeo") or {}
            )

            location = reverse_geo.get(
                "formattedLocation",
                ""
            )

            lat = gps.get("latitude", "")
            lon = gps.get("longitude", "")
            lat_long = f"{lat},{lon}"

            loc_time = None
            antiguedad_minutos = None

            if gps_time:
                loc_time = dp.parse(
                    gps_time
                ).astimezone(
                    pytz.timezone(MX_TZ)
                )

                antiguedad = now_mx - loc_time
                antiguedad_minutos = int(
                    antiguedad.total_seconds() / 60
                )

            registro_comparacion.update({
                "GpsTimeUTC": gps_time or "",
                "GpsTimeMexico": (
                    loc_time.strftime(
                        "%Y-%m-%d %H:%M:%S %Z"
                    )
                    if loc_time
                    else ""
                ),
                "AntiguedadMinutos": antiguedad_minutos,
                "Latitude": lat,
                "Longitude": lon,
                "Coordenadas": lat_long,
                "ReverseGeo": location,
                "GeocercaId": geocerca_id,
                "Geocerca": geocerca_name,
                "SpeedMilesPerHour": speed,
                "SpeedUsadaPorFiltro": speed,
                "IsEcuSpeed": ecu,
            })

            # ----------------------------------------------------------
            # FILTRO 1: GPS VIEJO
            # ----------------------------------------------------------
            if gps_time:
                if antiguedad > timedelta(hours=1):
                    omitidas_gps_viejo += 1

                    unidades_excluidas.append({
                        "Unidad": unidad_nombre,
                        "SamsaraVehicleId": unidad_id,
                        "Motivo": "GPS VIEJO",
                        "Detalle": (
                            f"Último GPS: "
                            f"{loc_time.strftime('%Y-%m-%d %H:%M:%S')} | "
                            f"Antigüedad minutos: "
                            f"{antiguedad_minutos}"
                        ),
                        "GeocercaId": geocerca_id,
                        "Geocerca": geocerca_name,
                        "Speed": speed,
                        "IsEcuSpeed": ecu,
                        "GpsTime": gps_time,
                    })

                    registrar_comparacion(
                        decision="EXCLUIDA",
                        motivo="GPS VIEJO",
                        detalle=(
                            f"Ultimo GPS: "
                            f"{loc_time.strftime('%Y-%m-%d %H:%M:%S')} | "
                            f"Antiguedad minutos: "
                            f"{antiguedad_minutos}"
                        )
                    )

                    print(
                        f"[EXCLUIDA GPS VIEJO] "
                        f"Unidad={unidad_nombre} | "
                        f"ID={unidad_id} | "
                        f"GPS={loc_time} | "
                        f"Antigüedad={antiguedad_minutos} min"
                    )

                    continue

            # ----------------------------------------------------------
            # FILTRO 2: GEOCERCA ESPECIAL
            # ----------------------------------------------------------
            if geocerca_id in predefined_special:
                omitidas_especiales += 1

                unidades_excluidas.append({
                    "Unidad": unidad_nombre,
                    "SamsaraVehicleId": unidad_id,
                    "Motivo": "GEOCERCA ESPECIAL",
                    "Detalle": (
                        "La geocerca está dentro de "
                        "predefined_special"
                    ),
                    "GeocercaId": geocerca_id,
                    "Geocerca": geocerca_name,
                    "Speed": speed,
                    "IsEcuSpeed": ecu,
                    "GpsTime": gps_time,
                })

                registrar_comparacion(
                    decision="EXCLUIDA",
                    motivo="GEOCERCA ESPECIAL",
                    detalle=(
                        "La geocerca esta dentro de "
                        "predefined_special"
                    )
                )

                print(
                    f"[EXCLUIDA GEOCERCA ESPECIAL] "
                    f"Unidad={unidad_nombre} | "
                    f"ID={unidad_id} | "
                    f"GeocercaId={geocerca_id} | "
                    f"Geocerca={geocerca_name}"
                )

                continue

            # ----------------------------------------------------------
            # FILTRO 3: PATIO O GEOCERCA EC5
            #
            # Únicamente se excluye si el ID está dentro de las
            # geocercas vinculadas al tag configurado.
            # ----------------------------------------------------------
            geocerca_detectada = ""

            if geocerca_id in Geocercas_EC5:
                geocerca_detectada = (
                    Geocercas_EC5[geocerca_id]
                )

            if geocerca_detectada:
                omitidas_patio += 1
                registro_comparacion[
                    "Geocerca"
                ] = geocerca_detectada

                unidades_excluidas.append({
                    "Unidad": unidad_nombre,
                    "SamsaraVehicleId": unidad_id,
                    "Motivo": "PATIO/GEOCERCA EC5",
                    "Detalle": (
                        f"La geocerca pertenece al tag "
                        f"{PARENT_TAG_ID}"
                    ),
                    "GeocercaId": geocerca_id,
                    "Geocerca": geocerca_detectada,
                    "Speed": speed,
                    "IsEcuSpeed": ecu,
                    "GpsTime": gps_time,
                })

                registrar_comparacion(
                    decision="EXCLUIDA",
                    motivo="PATIO/GEOCERCA EC5",
                    detalle=(
                        f"La geocerca pertenece al tag "
                        f"{PARENT_TAG_ID}"
                    )
                )

                print(
                    f"[EXCLUIDA PATIO] "
                    f"Unidad={unidad_nombre} | "
                    f"ID={unidad_id} | "
                    f"GeocercaId={geocerca_id} | "
                    f"Geocerca={geocerca_detectada}"
                )

                continue

            # ----------------------------------------------------------
            # FILTRO 4: SPEED 0 SIN ECU
            # ----------------------------------------------------------
            if speed is None:
                speed = 0

            registro_comparacion[
                "SpeedUsadaPorFiltro"
            ] = speed

            if speed == 0 and not ecu:
                omitidas_sin_ecu += 1

                unidades_excluidas.append({
                    "Unidad": unidad_nombre,
                    "SamsaraVehicleId": unidad_id,
                    "Motivo": "SPEED 0 SIN ECU",
                    "Detalle": (
                        "speedMilesPerHour es 0 y "
                        "isEcuSpeed es False"
                    ),
                    "GeocercaId": geocerca_id,
                    "Geocerca": geocerca_name,
                    "Speed": speed,
                    "IsEcuSpeed": ecu,
                    "GpsTime": gps_time,
                })

                registrar_comparacion(
                    decision="EXCLUIDA",
                    motivo="SPEED 0 SIN ECU",
                    detalle=(
                        "speedMilesPerHour es 0 y "
                        "isEcuSpeed es False"
                    )
                )

                print(
                    f"[EXCLUIDA SIN ECU] "
                    f"Unidad={unidad_nombre} | "
                    f"ID={unidad_id} | "
                    f"Speed={speed} | "
                    f"ECU={ecu}"
                )

                continue

            # ----------------------------------------------------------
            # UNIDAD INCLUIDA
            # ----------------------------------------------------------
            if speed == 0 and ecu:
                status = "DETENIDO"
            else:
                status = "RUTA"

            print(
                f"[INCLUIDA] "
                f"Unidad={unidad_nombre} | "
                f"ID={unidad_id} | "
                f"Speed={speed} | "
                f"ECU={ecu} | "
                f"GeocercaId={geocerca_id} | "
                f"Geocerca={geocerca_name} | "
                f"Estatus={status}"
            )

            results.append({
                "Unidad": unidad_nombre,
                "SamsaraVehicleId": unidad_id,
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
            })

            registrar_comparacion(
                decision="INCLUIDA",
                motivo="CUMPLE FILTROS",
                detalle=(
                    "La unidad se incluyo en el reporte"
                ),
                estatus=status
            )

        except Exception as error:
            logging.exception(
                f"Error procesando unidad "
                f"{unidad_nombre}"
            )

            unidades_con_error.append({
                "Unidad": unidad_nombre,
                "SamsaraVehicleId": unidad_id,
                "Error": str(error),
            })

            registrar_comparacion(
                decision="ERROR",
                motivo="ERROR DE PROCESAMIENTO",
                detalle=str(error)
            )

            print(
                f"[ERROR UNIDAD] "
                f"Unidad={unidad_nombre} | "
                f"ID={unidad_id} | "
                f"Error={error}"
            )

    # ------------------------------------------------------------------
    # RESUMEN DE DEPURACIÓN
    # ------------------------------------------------------------------
    total_procesado = (
        len(results)
        + len(unidades_excluidas)
        + len(unidades_con_error)
    )

    diferencia = (
        len(vehicles) - total_procesado
    )

    print("")
    print("=" * 100)
    print("RESUMEN DE DEPURACIÓN")
    print("=" * 100)

    print(
        f"Vehículos recibidos desde Samsara: "
        f"{len(vehicles)}"
    )

    print(
        f"Vehículos incluidos en el reporte: "
        f"{len(results)}"
    )

    print(
        f"Vehículos excluidos: "
        f"{len(unidades_excluidas)}"
    )

    print(
        f"Vehículos con error: "
        f"{len(unidades_con_error)}"
    )

    print(
        f"Total procesado: {total_procesado}"
    )

    print(
        f"Diferencia: {diferencia}"
    )

    print(
        f"Filas para comparacion externa: "
        f"{len(unidades_comparacion)}"
    )

    print("")
    print("DESGLOSE DE EXCLUSIONES")
    print(
        f"GPS viejo: {omitidas_gps_viejo}"
    )
    print(
        f"Geocerca especial: "
        f"{omitidas_especiales}"
    )
    print(
        f"Patio/geocerca EC5: "
        f"{omitidas_patio}"
    )
    print(
        f"Speed 0 sin ECU: "
        f"{omitidas_sin_ecu}"
    )

    print("")
    print("LISTA COMPLETA DE UNIDADES EXCLUIDAS")

    if not unidades_excluidas:
        print(
            "No se excluyó ninguna unidad."
        )

    for item in unidades_excluidas:
        print(
            f"Unidad={item['Unidad']} | "
            f"ID={item['SamsaraVehicleId']} | "
            f"Motivo={item['Motivo']} | "
            f"Detalle={item['Detalle']} | "
            f"Geocerca={item['Geocerca']} | "
            f"GeocercaId={item['GeocercaId']} | "
            f"Speed={item['Speed']} | "
            f"ECU={item['IsEcuSpeed']}"
        )

    if unidades_con_error:
        print("")
        print("LISTA DE UNIDADES CON ERROR")

        for item in unidades_con_error:
            print(
                f"Unidad={item['Unidad']} | "
                f"ID={item['SamsaraVehicleId']} | "
                f"Error={item['Error']}"
            )

    print("=" * 100)

    logging.info(
        f"Unidades recibidas: {len(vehicles)}"
    )

    logging.info(
        f"Unidades incluidas antes de detenciones: "
        f"{len(results)}"
    )

    logging.info(
        f"Unidades excluidas: "
        f"{len(unidades_excluidas)}"
    )

    logging.info(
        f"Unidades con error: "
        f"{len(unidades_con_error)}"
    )

    logging.info(
        f"Filas de comparacion: "
        f"{len(unidades_comparacion)} | "
        f"Diferencia contra recibidas: "
        f"{len(vehicles) - len(unidades_comparacion)}"
    )

    exportar_unidades_excluidas(
        unidades_excluidas
    )

    exportar_unidades_con_error(
        unidades_con_error
    )

    # ------------------------------------------------------------------
    # ENRIQUECER CON GOOGLE SHEETS
    # ------------------------------------------------------------------
    try:
        results = obtener_datos_google_sheets(
            results,
            now_mx
        )

    except Exception as error:
        logging.exception(
            "Error obteniendo datos de Google Sheets"
        )

        print(
            "⚠️ Error obteniendo datos de "
            "Google Sheets:",
            error
        )

        print(
            "⚠️ Se continuará con datos básicos "
            "de Samsara."
        )

    # ------------------------------------------------------------------
    # ENRIQUECER DETENIDOS
    # ------------------------------------------------------------------
    try:
        results = enriquecer_minutos_detenido(
            results=results,
            token=SAMSARA_API_TOKEN,
            now_mx=now_mx,
        )

    except Exception as error:
        logging.exception(
            "Error enriqueciendo minutos detenido"
        )

        print(
            "⚠️ Error enriqueciendo minutos detenido:",
            error
        )

        print(
            "⚠️ Se continuará sin información de "
            "minutos detenido."
        )

    print("")
    print("[DEBUG DESPUÉS DETENCIONES]")

    for row in results:
        print(
            f"Unidad={row.get('Unidad')} | "
            f"Estatus={row.get('Estatus')} | "
            f"TiempoDet={row.get('Tiempo Detenido')} | "
            f"TiempoTrafico={row.get('Tiempo Trafico')} | "
            f"Motor={row.get('Motor')} | "
            f"ECU={row.get('EcuSpeedActual')} | "
            f"Ventana={row.get('Ventana Detenido')} | "
            f"Geocerca={row.get('Geocerca')}"
        )

    resultados_por_id = {
        str(row.get("SamsaraVehicleId", "")): row
        for row in results
    }

    for registro in unidades_comparacion:
        resultado = resultados_por_id.get(
            str(registro.get("SamsaraVehicleId", ""))
        )

        if resultado:
            registro["EstatusFinal"] = resultado.get(
                "Estatus",
                ""
            )
            registro["MotorFinal"] = resultado.get(
                "Motor",
                ""
            )
            registro[
                "TiempoDetenidoFinal"
            ] = resultado.get(
                "Tiempo Detenido",
                ""
            )

    exportar_comparacion_samsara(
        unidades_comparacion
    )

    # ------------------------------------------------------------------
    # CONSTRUIR Y ENVIAR MENSAJE
    # ------------------------------------------------------------------
    try:
        mensaje_chat = construir_reporte_google(
            results,
            now_mx
        )

        print("")
        print("Mensaje para Google Chat:")
        print(mensaje_chat)

        enviar_google_chat(
            mensaje_chat
        )

    except Exception as error:
        logging.exception(
            "Error construyendo o enviando "
            "mensaje a Google Chat"
        )

        print(
            "❌ Error construyendo o enviando "
            "mensaje a Google Chat:",
            error
        )

        sys.exit(1)

    logging.info(
        "===> Ejecución finalizada correctamente"
    )

    print(
        "===> Ejecución finalizada correctamente"
    )


# ----------------------------------------------------------------------
# EJECUCIÓN
# ----------------------------------------------------------------------
if __name__ == "__main__":
    main()
