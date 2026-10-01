"""Envio configurable de reportes de telemetria Samsara.

La seleccion, contenido y canales se definen en config/reportes.json.
"""
from __future__ import annotations

import argparse
from copy import deepcopy
from datetime import datetime, timedelta
from email.message import EmailMessage
from io import BytesIO
import json
import logging
import math
import os
from pathlib import Path
import re
import smtplib
import sys
import unicodedata
from typing import Any, Iterable
from urllib.parse import quote

from dateutil import parser as dp
from dotenv import load_dotenv
import gspread
from gspread_dataframe import get_as_dataframe
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import pytz
import requests

from detenciones import enriquecer_minutos_detenido, haversine_meters
from LogitrackEstatus import (
    DEFAULT_API_URL as LOGITRACK_DEFAULT_API_URL,
    DEFAULT_TOKEN_URL as LOGITRACK_DEFAULT_TOKEN_URL,
    fetch_last_status,
    request_access_token,
)


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG_PATH = BASE_DIR / "config" / "reportes.json"
DEFAULT_TAG_CATALOG_PATH = BASE_DIR / "config" / "catalogo_etiquetas_samsara.json"
SAMSARA_BASE_URL = "https://api.samsara.com"
DEFAULT_TIMEZONE = "America/Mexico_City"
MAX_GOOGLE_CHAT_CHARS = 30_000
WEBHOOK_PRUEBAS_ENV = "GOOGLE_CHAT_WEBHOOK_URL_PRUEBAS"
load_dotenv(dotenv_path=BASE_DIR / ".env")

logging.basicConfig(
    filename=str(BASE_DIR / "reporte_logs.log"),
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)


class ConfiguracionError(ValueError):
    """La configuracion no permite construir un reporte seguro."""


def normalizar_texto(valor: Any) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or "").strip())
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.casefold().split())


def slug(valor: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", normalizar_texto(valor)).strip("_") or "reporte"


KMH_POR_MPH = 1.609344
MOTIVOS_SAMSARA_SIN_GPS_VIGENTE = {"GPS VIEJO", "GPS SIN FECHA"}


def normalizar_nombre_unidad(
    valor: Any, sufijos: Iterable[str] = ("TDR",)
) -> str:
    """Construye una llave comun para nombres de unidad de ambos proveedores.

    Ignora mayusculas, espacios, guiones, ceros a la izquierda y el identificador
    de empresa al inicio o al final: 2234 - TDR, 2234tdr y TDR-02234 dan 2234.
    """
    llave = re.sub(r"[^a-z0-9]+", "", normalizar_texto(valor))
    tokens = [
        token for token in (re.sub(r"[^a-z0-9]+", "", normalizar_texto(s)) for s in sufijos)
        if token
    ]
    recortada = True
    while recortada:
        recortada = False
        for token in tokens:
            if llave != token and llave.endswith(token):
                llave, recortada = llave[: -len(token)], True
            if llave != token and llave.startswith(token):
                llave, recortada = llave[len(token):], True
    if not llave:
        return ""
    return llave.lstrip("0") or "0"


def parsear_fecha_logitrack(valor: Any, now_mx: datetime) -> datetime | None:
    if not valor:
        return None
    try:
        fecha = dp.parse(str(valor))
    except (TypeError, ValueError, OverflowError):
        return None
    if fecha.tzinfo is None:
        zona = pytz.timezone(getattr(now_mx.tzinfo, "zone", DEFAULT_TIMEZONE))
        fecha = zona.localize(fecha)
    return fecha.astimezone(now_mx.tzinfo)


def convertir_float(valor: Any) -> float | None:
    if valor is None or str(valor).strip() == "":
        return None
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def motor_logitrack(status: dict[str, Any]) -> str:
    encendido = status.get("engine_ign")
    if encendido is None:
        encendido = status.get("io_ign")
    texto = str(encendido).strip()
    return "ON" if texto == "1" else "OFF" if texto == "0" else ""


def indexar_estatus_logitrack(
    statuses: Iterable[dict[str, Any]],
    now_mx: datetime,
    settings: dict[str, Any],
) -> dict[str, dict[str, Any]]:
    """Indexa por unidad y conserva el reporte mas reciente ante duplicados."""
    sufijos = settings.get("sufijos_unidad") or ["TDR"]
    indice: dict[str, dict[str, Any]] = {}
    for status in statuses:
        llave = normalizar_nombre_unidad(status.get("unit_name"), sufijos)
        if not llave:
            continue
        fecha = parsear_fecha_logitrack(
            status.get("datetime") or status.get("event_time"), now_mx
        )
        anterior = indice.get(llave)
        fecha_anterior = (anterior or {}).get("_fecha_logitrack")
        if anterior is None or (fecha and (fecha_anterior is None or fecha > fecha_anterior)):
            item = dict(status)
            item["_fecha_logitrack"] = fecha
            indice[llave] = item
    return indice


def leer_status_logitrack(
    status: dict[str, Any], now_mx: datetime, settings: dict[str, Any]
) -> dict[str, Any]:
    """Resume una lectura Logitrack: antiguedad, vigencia, velocidad (km/h) y estado."""
    max_antiguedad = int(settings.get("max_antiguedad_minutos", 60))
    stop_logitrack = float(
        settings.get(
            "velocidad_detenido_logitrack_kmh",
            settings.get("velocidad_detenido_logitrack", 5),
        )
    )
    fecha = status.get("_fecha_logitrack")
    antiguedad = (now_mx - fecha).total_seconds() / 60 if fecha is not None else None
    fresco = antiguedad is not None and -5 <= antiguedad <= max_antiguedad
    velocidad = convertir_float(status.get("speed"))
    detenido = None if velocidad is None else velocidad <= stop_logitrack
    if not fresco:
        estado = "DESACTUALIZADO"
    elif detenido is None:
        estado = "SIN VELOCIDAD"
    else:
        estado = "DETENIDO" if detenido else "RUTA"
    return {
        "fecha": fecha,
        "fresco": fresco,
        "velocidad": velocidad,
        "detenido": detenido,
        "campos": {
            "Logitrack Encontrado": "SI",
            "Estatus Logitrack": estado,
            "Fecha Logitrack": fecha,
            "Antiguedad Logitrack Min": round(antiguedad, 1) if antiguedad is not None else None,
            "Velocidad Logitrack": velocidad,
            "Motor Logitrack": motor_logitrack(status),
            "Ubicación Logitrack": str(status.get("address") or ""),
            "Logitrack Fresco": "SI" if fresco else "NO",
        },
    }


def enriquecer_con_logitrack(
    results: list[dict[str, Any]],
    statuses: Iterable[dict[str, Any]],
    now_mx: datetime,
    settings: dict[str, Any],
) -> list[dict[str, Any]]:
    """Cruza la lectura actual de ambas telemetrias y elige la mas real.

    Si coinciden, la lectura queda respaldada por ambas; si difieren, gana la mas
    reciente. La distancia tolerada crece con la velocidad y el desfase entre
    lecturas, porque una unidad en ruta avanza entre un reporte y otro.
    """
    indice = indexar_estatus_logitrack(statuses, now_mx, settings)
    sufijos = settings.get("sufijos_unidad") or ["TDR"]
    distancia_base = float(settings.get("distancia_detenido_metros", 150))
    factor = float(settings.get("factor_tolerancia_movimiento", 1.3))
    desfase_minimo = float(settings.get("desfase_minimo_minutos", 1))
    velocidad_maxima = float(settings.get("velocidad_maxima_plausible_kmh", 120))
    stop_samsara = float(settings.get("velocidad_detenido_samsara_mph", 3))
    encontrados = frescos = dobles = desacuerdos = lejanas = 0
    distancias_detenidas: list[float] = []

    for row in results:
        estatus_samsara_actual = str(row.get("Estatus") or "")
        row.update({
            "Estatus Samsara Actual": estatus_samsara_actual,
            "Logitrack Encontrado": "NO",
            "Estatus Logitrack": "SIN COINCIDENCIA",
            "Fecha Logitrack": None,
            "Antiguedad Logitrack Min": None,
            "Velocidad Logitrack": None,
            "Motor Logitrack": "",
            "Ubicación Logitrack": "",
            "Distancia GPS Metros": None,
            "Distancia Esperada Metros": None,
            "Desfase Lecturas Min": None,
            "Ubicación Coincide": "",
            "Logitrack Fresco": "NO",
            "Doble Comprobación Actual": "NO",
            "Fuente Confirmación": "SAMSARA ACTUAL",
            "Motivo Decisión": "Sin coincidencia en Logitrack",
        })
        llave = normalizar_nombre_unidad(row.get("Unidad"), sufijos)
        status = indice.get(llave)
        gps_samsara = row.get("GpsActual") or {}
        velocidad_samsara = convertir_float(gps_samsara.get("speedMilesPerHour")) or 0.0
        samsara_detenido = velocidad_samsara <= stop_samsara

        # Una lectura Samsara cercana a cero debe revisar historial aunque no
        # venga marcada como velocidad ECU; Logitrack aporta la segunda senal.
        if samsara_detenido:
            row["Estatus"] = "DETENIDO"

        if not status:
            continue
        encontrados += 1
        lectura = leer_status_logitrack(status, now_mx, settings)
        row.update(lectura["campos"])
        fecha_logitrack = lectura["fecha"]
        velocidad_logitrack = lectura["velocidad"]
        logitrack_detenido = lectura["detenido"]

        # Desfase positivo: la lectura Samsara es mas reciente que la Logitrack.
        fecha_samsara = row.get("Fecha GPS")
        if not isinstance(fecha_samsara, datetime):
            fecha_samsara = now_mx
        desfase = (
            (fecha_samsara - fecha_logitrack).total_seconds() / 60
            if fecha_logitrack is not None else None
        )
        distancia = None
        try:
            distancia = haversine_meters(
                gps_samsara.get("latitude"),
                gps_samsara.get("longitude"),
                status.get("lat"),
                status.get("lon"),
            )
        except (TypeError, ValueError):
            pass
        # Solo la velocidad de una unidad en movimiento explica distancia entre
        # lecturas; la deriva GPS de una unidad detenida no debe ampliar la tolerancia.
        segundos_desfase = abs(desfase or 0) * 60
        velocidad_movimiento_kmh = max(
            0 if samsara_detenido else velocidad_samsara * KMH_POR_MPH,
            0 if logitrack_detenido in (True, None) else velocidad_logitrack,
        )
        esperada = distancia_base + velocidad_movimiento_kmh / 3.6 * segundos_desfase * factor
        maxima = max(esperada, distancia_base + velocidad_maxima / 3.6 * segundos_desfase)
        coincide = None if distancia is None else distancia <= esperada
        imposible = distancia is not None and distancia > maxima
        row.update({
            "Distancia GPS Metros": round(distancia, 1) if distancia is not None else None,
            "Distancia Esperada Metros": round(esperada),
            "Desfase Lecturas Min": round(desfase, 1) if desfase is not None else None,
            "Ubicación Coincide": "" if coincide is None else "SI" if coincide else "NO",
        })

        if not lectura["fresco"]:
            row["Motivo Decisión"] = "La lectura Logitrack excede la antiguedad permitida"
            continue
        frescos += 1
        if imposible:
            lejanas += 1
            row["Estatus"] = "REVISAR"
            row["Fuente Confirmación"] = "SAMSARA VS LOGITRACK"
            row["Motivo Decisión"] = (
                f"Ubicación no coincide: {distancia:.0f} m entre equipos en "
                f"{abs(desfase or 0):.1f} min (máximo posible {maxima:.0f} m)"
            )
            continue
        if logitrack_detenido is None:
            row["Motivo Decisión"] = "Logitrack no reporta velocidad; se usa Samsara"
            continue

        if samsara_detenido == logitrack_detenido:
            row["Fuente Confirmación"] = "SAMSARA + LOGITRACK"
            if not samsara_detenido:
                row["Estatus"] = "RUTA"
                row["Motivo Decisión"] = "Ambas telemetrías reportan movimiento"
            elif coincide:
                dobles += 1
                distancias_detenidas.append(distancia)
                row["Doble Comprobación Actual"] = "SI"
                row["Motivo Decisión"] = "Ambas lecturas actuales indican detención y coinciden en ubicación"
            elif coincide is False:
                row["Fuente Confirmación"] = "SAMSARA ACTUAL"
                row["Motivo Decisión"] = (
                    f"Ambas indican detención a {distancia:.0f} m; la unidad cambió de "
                    f"posición entre lecturas ({abs(desfase or 0):.1f} min)"
                )
            else:
                row["Fuente Confirmación"] = "SAMSARA ACTUAL"
                row["Motivo Decisión"] = "Ambas indican detención, sin coordenadas para comparar"
            continue

        desacuerdos += 1
        if desfase is not None and desfase <= -desfase_minimo:
            minutos = abs(desfase)
            row["Fuente Confirmación"] = "LOGITRACK"
            if logitrack_detenido:
                row["Estatus"] = "DETENIDO LOGITRACK"
                row["Motivo Decisión"] = (
                    f"Logitrack ({minutos:.1f} min más reciente) reporta detención; "
                    "Samsara aún reporta movimiento"
                )
            else:
                row["Estatus"] = "RUTA"
                row["Motivo Decisión"] = (
                    f"Logitrack ({minutos:.1f} min más reciente) reporta movimiento"
                )
        elif samsara_detenido:
            row["Motivo Decisión"] = (
                "Samsara (lectura más reciente) reporta detención; Logitrack aún reporta movimiento"
            )
        else:
            row["Estatus"] = "RUTA"
            row["Motivo Decisión"] = (
                "Samsara (lectura más reciente) reporta movimiento; Logitrack aún reporta detención"
            )

    promedio = (
        f"{sum(distancias_detenidas) / len(distancias_detenidas):.0f} m"
        if distancias_detenidas else "N/D"
    )
    print(
        f"[Logitrack] Vinculadas={encontrados}/{len(results)} Frescas={frescos} "
        f"DobleSenalDetenida={dobles} Desacuerdos={desacuerdos} "
        f"UbicacionNoCoincide={lejanas} DistanciaPromedioDetenidas={promedio}"
    )
    sin_match = [str(r.get("Unidad")) for r in results if r.get("Logitrack Encontrado") == "NO"]
    if sin_match:
        print(f"[Logitrack] Sin coincidencia: {', '.join(sin_match)}")
    retraso = mediana_antiguedad_logitrack(results)
    if retraso is not None and retraso > RETRASO_LOGITRACK_AVISO_MINUTOS:
        print(f"[Logitrack][AVISO] Lecturas con retraso: mediana {retraso:.0f} min")
    return results


def punto_en_geocerca(lat: Any, lon: Any, geofence: dict[str, Any]) -> bool:
    """Evalua si una coordenada cae en una geocerca Samsara (circulo o poligono)."""
    lat, lon = convertir_float(lat), convertir_float(lon)
    if lat is None or lon is None:
        return False
    circulo = geofence.get("circle")
    if circulo:
        return haversine_meters(
            lat, lon, circulo.get("latitude"), circulo.get("longitude")
        ) <= float(circulo.get("radiusMeters") or 0)
    vertices = [
        (float(v["latitude"]), float(v["longitude"]))
        for v in (geofence.get("polygon") or {}).get("vertices") or []
    ]
    dentro = False
    for (lat1, lon1), (lat2, lon2) in zip(vertices, vertices[-1:] + vertices[:-1]):
        if (lat1 > lat) != (lat2 > lat):
            cruce = lon1 + (lat - lat1) * (lon2 - lon1) / (lat2 - lat1)
            if lon < cruce:
                dentro = not dentro
    return dentro


def distancia_a_geocerca(lat: Any, lon: Any, geofence: dict[str, Any]) -> float | None:
    """Metros desde la coordenada al borde de la geocerca; 0 si esta dentro."""
    lat, lon = convertir_float(lat), convertir_float(lon)
    if lat is None or lon is None:
        return None
    circulo = geofence.get("circle")
    if circulo:
        centro = haversine_meters(lat, lon, circulo.get("latitude"), circulo.get("longitude"))
        return max(0.0, centro - float(circulo.get("radiusMeters") or 0))
    vertices = (geofence.get("polygon") or {}).get("vertices") or []
    if not vertices:
        return None
    if punto_en_geocerca(lat, lon, geofence):
        return 0.0
    # Proyeccion local en metros; suficiente para distancias de pocos kilometros.
    escala_lon = 111_320 * math.cos(math.radians(lat))
    puntos = [
        ((float(v["longitude"]) - lon) * escala_lon, (float(v["latitude"]) - lat) * 110_540)
        for v in vertices
    ]
    minima = None
    for (x1, y1), (x2, y2) in zip(puntos, puntos[1:] + puntos[:1]):
        dx, dy = x2 - x1, y2 - y1
        largo = dx * dx + dy * dy
        t = 0.0 if largo == 0 else max(0.0, min(1.0, -(x1 * dx + y1 * dy) / largo))
        distancia = math.hypot(x1 + t * dx, y1 + t * dy)
        minima = distancia if minima is None else min(minima, distancia)
    return minima


def geocerca_mas_cercana(
    lat: Any, lon: Any, geocercas: dict[str, dict[str, Any]]
) -> tuple[str, float] | None:
    cercana = None
    for geocerca in geocercas.values():
        distancia = distancia_a_geocerca(lat, lon, geocerca.get("geofence") or {})
        if distancia is not None and (cercana is None or distancia < cercana[1]):
            cercana = (geocerca.get("nombre") or "", distancia)
    return cercana


def obtener_geometrias_geocercas(
    sesion: requests.Session,
    etiqueta_ids: Iterable[str],
    especiales_ids: Iterable[str] = (),
    direcciones: list[dict[str, Any]] | None = None,
) -> dict[str, dict[str, Any]]:
    """Nombre, geometria y motivo de las geocercas excluidas (por etiqueta) y especiales (por ID)."""
    etiquetas = {str(x) for x in etiqueta_ids if x}
    especiales = {str(x) for x in especiales_ids if x}
    if not etiquetas and not especiales:
        return {}
    if direcciones is None:
        direcciones = obtener_paginas(sesion, "/addresses", {"limit": 512})
    geometrias = {}
    for d in direcciones:
        address_id = str(d.get("id") or "")
        if address_id in especiales:
            motivo = "GEOCERCA ESPECIAL"
        elif etiquetas & {str(t.get("id")) for t in d.get("tags") or []}:
            motivo = "PATIO/GEOCERCA EXCLUIDA"
        else:
            continue
        geometrias[address_id] = {
            "nombre": str(d.get("name") or ""), "geofence": d.get("geofence") or {}, "motivo": motivo,
        }
    return geometrias


def excluir_por_coordenadas_en_geocerca(
    results: list[dict[str, Any]], geocercas: dict[str, dict[str, Any]]
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Excluye unidades cuyas coordenadas caen dentro de una geocerca excluida o especial.

    Samsara no siempre llena gps.address aunque la unidad este dentro de la
    geocerca, por eso se valida tambien contra el poligono.
    """
    conservadas, omitidas = [], []
    for row in results:
        geocerca = next(
            (g for g in geocercas.values()
             if punto_en_geocerca(row.get("Latitud"), row.get("Longitud"), g.get("geofence") or {})),
            None,
        )
        if geocerca is None:
            conservadas.append(row)
            continue
        omitidas.append({
            "Unidad": row.get("Unidad"), "SamsaraVehicleId": row.get("SamsaraVehicleId"),
            "Motivo": geocerca.get("motivo") or "PATIO/GEOCERCA EXCLUIDA",
            "Detalle": f"Coordenadas dentro de {geocerca['nombre']}; Samsara no reportó la geocerca",
            "GpsTimeMexico": row.get("Fecha GPS"), "Geocerca": geocerca["nombre"], "GeocercaId": "",
            "Latitud": row.get("Latitud"), "Longitud": row.get("Longitud"),
            "Coordenadas": row.get("Coordenadas"), "VelocidadMph": row.get("Velocidad Mph"),
            "IsEcuSpeed": row.get("IsEcuSpeed"),
        })
    if omitidas:
        detalle = ", ".join(f"{x['Unidad']} ({x['Geocerca']})" for x in omitidas)
        print(f"[Geocercas] Dentro de geocerca por coordenadas: {len(omitidas)} -> {detalle}")
    return conservadas, omitidas


ESTADOS_DETENIDOS_DEFAULT = (
    "DETENIDO", "DETENIDO CONFIRMADO", "DETENIDO SAMSARA", "DETENIDO LOGITRACK", "REVISAR",
)


def omitir_detenidas_cerca_de_geocerca(
    results: list[dict[str, Any]],
    geocercas: dict[str, dict[str, Any]],
    distancia_max: float,
    estados: Iterable[str] | None = None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Omite unidades detenidas junto a un patio: esperan afuera y no llevan viaje."""
    estados_omitir = {normalizar_texto(x) for x in (estados or ESTADOS_DETENIDOS_DEFAULT)}
    conservadas, omitidas = [], []
    for row in results:
        estatus = str(row.get("Estatus") or "")
        cercana = None
        if normalizar_texto(estatus) in estados_omitir:
            cercana = geocerca_mas_cercana(row.get("Latitud"), row.get("Longitud"), geocercas)
        if cercana is None or cercana[1] > distancia_max:
            conservadas.append(row)
            continue
        nombre, distancia = cercana
        omitidas.append({
            "Unidad": row.get("Unidad"), "SamsaraVehicleId": row.get("SamsaraVehicleId"),
            "Motivo": "CERCA DE GEOCERCA",
            "Detalle": f"{estatus} a {distancia:.0f} m de {nombre} (máximo {distancia_max:.0f} m)",
            "GpsTimeMexico": row.get("Fecha GPS"), "Geocerca": nombre, "GeocercaId": "",
            "Latitud": row.get("Latitud"), "Longitud": row.get("Longitud"),
            "Coordenadas": row.get("Coordenadas"), "VelocidadMph": row.get("Velocidad Mph"),
            "IsEcuSpeed": row.get("IsEcuSpeed"),
        })
    if omitidas:
        detalle = ", ".join(f"{x['Unidad']} ({x['Geocerca']})" for x in omitidas)
        print(f"[Geocercas] Detenidas cerca de geocerca omitidas: {len(omitidas)} -> {detalle}")
    return conservadas, omitidas


def rescatar_con_logitrack(
    excluidas: list[dict[str, Any]],
    statuses: Iterable[dict[str, Any]],
    now_mx: datetime,
    settings: dict[str, Any],
    geocercas: dict[str, dict[str, Any]] | None = None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Usa Logitrack vigente para unidades omitidas porque su GPS Samsara no esta vigente.

    `geocercas` son las geocercas excluidas con su geometria; una unidad que
    Logitrack ubica dentro de alguna sigue omitida. None significa que no se
    pudo validar y, por seguridad, no se rescata ninguna unidad.
    """
    if not settings.get("usar_logitrack_si_samsara_viejo", True) or geocercas is None:
        return [], excluidas
    geocercas = geocercas or {}
    indice = indexar_estatus_logitrack(statuses, now_mx, settings)
    sufijos = settings.get("sufijos_unidad") or ["TDR"]
    rescatadas, restantes = [], []
    for omitida in excluidas:
        status = None
        if omitida.get("Motivo") in MOTIVOS_SAMSARA_SIN_GPS_VIGENTE:
            status = indice.get(normalizar_nombre_unidad(omitida.get("Unidad"), sufijos))
        lectura = leer_status_logitrack(status, now_mx, settings) if status else None
        if not lectura or not lectura["fresco"] or lectura["detenido"] is None:
            restantes.append(omitida)
            continue
        lat, lon = status.get("lat"), status.get("lon")
        geocerca = next(
            (g["nombre"] for g in geocercas.values() if punto_en_geocerca(lat, lon, g["geofence"])),
            None,
        )
        if geocerca is not None:
            restantes.append({
                **omitida,
                "Detalle": f"{omitida.get('Detalle') or ''}; Logitrack la ubica en geocerca {geocerca}".lstrip("; "),
            })
            continue
        row = {
            "Unidad": omitida.get("Unidad"),
            "SamsaraVehicleId": omitida.get("SamsaraVehicleId"),
            "Fecha GPS": omitida.get("GpsTimeMexico"),
            "Ubicación": str(status.get("address") or ""),
            "Estatus": "DETENIDO LOGITRACK" if lectura["detenido"] else "RUTA",
            "Velocidad Mph": omitida.get("VelocidadMph"),
            "IsEcuSpeed": omitida.get("IsEcuSpeed"),
            "Latitud": lat,
            "Longitud": lon,
            "Coordenadas": f"{lat},{lon}" if lat not in (None, "") and lon not in (None, "") else "",
            "Geocerca": omitida.get("Geocerca", ""),
            "Estatus Samsara Actual": omitida.get("Motivo"),
            "Doble Comprobación Actual": "NO",
            "Fuente Confirmación": "LOGITRACK",
            "Motivo Decisión": (
                f"Samsara sin GPS vigente ({omitida.get('Detalle') or omitida.get('Motivo')}); "
                "se usa Logitrack"
            ),
        }
        row.update(lectura["campos"])
        rescatadas.append(row)
    if rescatadas:
        print(f"[Logitrack] Unidades con GPS Samsara viejo cubiertas por Logitrack: {len(rescatadas)}")
    return rescatadas, restantes


def finalizar_comprobacion_telemetria(
    results: list[dict[str, Any]], settings: dict[str, Any]
) -> list[dict[str, Any]]:
    """Convierte el analisis historico en un estatus final auditable."""
    minimo = int(settings.get("detencion_minima_minutos", 5))
    for row in results:
        estado = str(row.get("Estatus") or "").strip()
        if estado != "DETENIDO":
            if estado in {"RUTA", "TRAFICO LENTO"} and row.get("Ventana Detenido"):
                row["Fuente Confirmación"] = "HISTORICO SAMSARA"
                row["Motivo Decisión"] = "El historial Samsara muestra movimiento reciente"
            continue

        minutos = row.get("Minutos Detenido")
        historia_confirma = isinstance(minutos, (int, float)) and minutos >= minimo
        doble_actual = row.get("Doble Comprobación Actual") == "SI"
        if historia_confirma and doble_actual:
            row["Estatus"] = "DETENIDO CONFIRMADO"
            row["Fuente Confirmación"] = "HISTORICO SAMSARA + LOGITRACK"
            row["Motivo Decisión"] = "Historial estacionario y doble lectura actual coincidente"
        elif historia_confirma:
            row["Estatus"] = "DETENIDO SAMSARA"
            row["Fuente Confirmación"] = "HISTORICO SAMSARA"
            if row.get("Logitrack Fresco") == "SI":
                row["Motivo Decisión"] = "Historial detenido, pero Logitrack no coincide"
            else:
                row["Motivo Decisión"] = "Historial detenido sin una lectura Logitrack vigente"
        else:
            row["Estatus"] = "REVISAR"
            row["Fuente Confirmación"] = "INFORMACION INSUFICIENTE"
            row["Motivo Decisión"] = f"No se confirmaron al menos {minimo} minutos de detención"
    return results


RETRASO_LOGITRACK_AVISO_MINUTOS = 10


def mediana_antiguedad_logitrack(results: Iterable[dict[str, Any]]) -> float | None:
    edades = sorted(
        row["Antiguedad Logitrack Min"] for row in results
        if isinstance(row.get("Antiguedad Logitrack Min"), (int, float))
    )
    return edades[len(edades) // 2] if edades else None


def resumen_telemetria(results: list[dict[str, Any]]) -> list[str]:
    """Lineas de calidad del cruce Samsara vs Logitrack para el mensaje de Chat."""
    if not any("Logitrack Encontrado" in row for row in results):
        return []
    vinculadas = sum(1 for row in results if row.get("Logitrack Encontrado") == "SI")
    distancias = sorted(
        row["Distancia GPS Metros"] for row in results
        if row.get("Doble Comprobación Actual") == "SI"
        and isinstance(row.get("Distancia GPS Metros"), (int, float))
    )
    lejanas = sum(
        1 for row in results
        if str(row.get("Motivo Decisión") or "").startswith("Ubicación no coincide")
    )
    lineas = [
        "", "📡 *Samsara vs Logitrack*",
        f"🔗 Vinculadas: {vinculadas}/{len(results)}",
        f"🤝 Detenidas en ambas: {len(distancias)}",
    ]
    retraso = mediana_antiguedad_logitrack(results)
    if retraso is not None and retraso > RETRASO_LOGITRACK_AVISO_MINUTOS:
        lineas.append(
            f"⚠️ Logitrack con retraso: mediana {retraso:.0f} min; "
            "la doble comprobación depende solo de lecturas vigentes"
        )
    if distancias:
        lineas.append(
            f"📏 Distancia entre equipos (detenidas): promedio "
            f"{sum(distancias) / len(distancias):.0f} m, mediana "
            f"{distancias[len(distancias) // 2]:.0f} m"
        )
    lineas.append(f"📍 Ubicación no coincide: {lejanas}")
    sin_match = [str(row.get("Unidad")) for row in results if row.get("Logitrack Encontrado") == "NO"]
    if sin_match:
        lineas.append(f"❓ Sin coincidencia: {', '.join(sin_match[:20])}")
    return lineas


def cargar_configuracion(ruta: Path) -> dict[str, Any]:
    if not ruta.exists():
        raise ConfiguracionError(f"No existe el archivo de configuracion: {ruta}")
    try:
        config = json.loads(ruta.read_text(encoding="utf-8"))
    except json.JSONDecodeError as error:
        raise ConfiguracionError(f"JSON invalido en {ruta}: {error}") from error

    reportes = config.get("reportes")
    if not isinstance(reportes, list) or not reportes:
        raise ConfiguracionError("La configuracion debe contener 'reportes'.")
    nombres: set[str] = set()
    perfiles = config.get("perfiles_filtros") or {}
    for indice, reporte in enumerate(reportes, start=1):
        nombre = str(reporte.get("nombre") or "").strip()
        if not nombre:
            raise ConfiguracionError(f"El reporte #{indice} no tiene nombre.")
        clave = normalizar_texto(nombre)
        if clave in nombres:
            raise ConfiguracionError(f"Nombre de reporte duplicado: {nombre}")
        nombres.add(clave)
        if not (reporte.get("etiquetas") or reporte.get("etiqueta_ids")):
            raise ConfiguracionError(f"El reporte '{nombre}' necesita etiquetas o etiqueta_ids.")
        tipo = reporte.get("tipo_filtro_etiqueta", "tagIds")
        if tipo not in {"tagIds", "parentTagIds"}:
            raise ConfiguracionError(f"'{nombre}': use tagIds o parentTagIds.")
        perfil = reporte.get("perfil_filtros")
        if perfil and perfil not in perfiles:
            raise ConfiguracionError(f"'{nombre}': perfil_filtros inexistente: {perfil}")
        for entrega in reporte.get("entregas") or []:
            canal = str(entrega.get("canal") or "").strip().lower()
            if canal not in {"google_chat", "correo"}:
                raise ConfiguracionError(f"'{nombre}': canal no soportado: {canal or '(vacio)'}")
    return config


def crear_sesion_samsara(token: str) -> requests.Session:
    sesion = requests.Session()
    sesion.headers.update({"Accept": "application/json", "Authorization": f"Bearer {token}"})
    return sesion


def obtener_paginas(
    sesion: requests.Session,
    ruta: str,
    params: dict[str, Any] | None = None,
    timeout: int = 60,
) -> list[dict[str, Any]]:
    """Consume todas las paginas de un endpoint Samsara basado en cursor."""
    url = ruta if ruta.startswith("http") else f"{SAMSARA_BASE_URL}{ruta}"
    base_params = dict(params or {})
    resultado: list[dict[str, Any]] = []
    after: str | None = None
    while True:
        pagina_params = dict(base_params)
        if after:
            pagina_params["after"] = after
        response = sesion.get(url, params=pagina_params, timeout=timeout)
        logging.info("GET %s status=%s", response.url, response.status_code)
        response.raise_for_status()
        payload = response.json()
        resultado.extend(payload.get("data") or [])
        pagination = payload.get("pagination") or {}
        if not pagination.get("hasNextPage"):
            return resultado
        after = pagination.get("endCursor")
        if not after:
            raise RuntimeError(f"{ruta} indico otra pagina pero no envio endCursor")


def listar_etiquetas_samsara(sesion: requests.Session) -> list[dict[str, Any]]:
    return obtener_paginas(sesion, "/tags", {"limit": 512})


def obtener_catalogo_etiquetas_samsara(
    sesion: requests.Session, limit: int = 10
) -> list[dict[str, Any]]:
    """Descarga tags por lotes pequenos y conserva solo datos de jerarquia."""
    resultado: list[dict[str, Any]] = []
    after: str | None = None
    while True:
        params: dict[str, Any] = {"limit": limit}
        if after:
            params["after"] = after
        response = sesion.get(f"{SAMSARA_BASE_URL}/tags", params=params, timeout=120)
        response.raise_for_status()
        payload = response.json()
        for tag in payload.get("data") or []:
            resultado.append(
                {
                    "id": str(tag.get("id") or ""),
                    "nombre": str(tag.get("name") or "").strip(),
                    "parentTagId": str(tag.get("parentTagId") or ""),
                }
            )
        pagination = payload.get("pagination") or {}
        print(f"[Tags] Descargadas: {len(resultado)}", flush=True)
        if not pagination.get("hasNextPage"):
            return resultado
        after = pagination.get("endCursor")
        if not after:
            raise RuntimeError("Samsara no devolvio endCursor para la siguiente pagina de tags")


def construir_jerarquia_etiquetas(tags: list[dict[str, Any]]) -> dict[str, Any]:
    nodos = {
        str(tag["id"]): {
            "id": str(tag["id"]),
            "nombre": str(tag.get("nombre") or ""),
            "parentTagId": str(tag.get("parentTagId") or ""),
            "hijos": [],
        }
        for tag in tags
        if tag.get("id")
    }
    raices = []
    for nodo in nodos.values():
        padre = nodos.get(nodo["parentTagId"])
        if padre:
            padre["hijos"].append(nodo)
        else:
            raices.append(nodo)

    def ordenar(nodo: dict[str, Any]) -> None:
        nodo["hijos"].sort(key=lambda x: normalizar_texto(x["nombre"]))
        for hijo in nodo["hijos"]:
            ordenar(hijo)

    raices.sort(key=lambda x: normalizar_texto(x["nombre"]))
    for raiz in raices:
        ordenar(raiz)
    padres = sorted(
        (
            {"id": nodo["id"], "nombre": nodo["nombre"], "parentTagId": nodo["parentTagId"]}
            for nodo in nodos.values()
            if nodo["hijos"]
        ),
        key=lambda x: normalizar_texto(x["nombre"]),
    )
    return {
        "generado": datetime.now(pytz.timezone(DEFAULT_TIMEZONE)).isoformat(),
        "totalEtiquetas": len(nodos),
        "totalEtiquetasPadre": len(padres),
        "etiquetasPadre": padres,
        "jerarquia": raices,
        "etiquetas": sorted(tags, key=lambda x: normalizar_texto(x.get("nombre"))),
    }


def guardar_catalogo_etiquetas(
    sesion: requests.Session, ruta: Path, limit: int = 10
) -> dict[str, Any]:
    catalogo = construir_jerarquia_etiquetas(obtener_catalogo_etiquetas_samsara(sesion, limit))
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_text(json.dumps(catalogo, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return catalogo


def cargar_cache_catalogo(ruta: Path) -> dict[str, dict[str, str]]:
    if not ruta.exists():
        return {}
    data = json.loads(ruta.read_text(encoding="utf-8"))
    return {
        normalizar_texto(tag.get("nombre")): {
            "id": str(tag.get("id") or ""),
            "name": str(tag.get("nombre") or ""),
        }
        for tag in data.get("etiquetas") or []
        if tag.get("id") and tag.get("nombre")
    }


def resolver_etiquetas(
    nombres_solicitados: Iterable[str], etiquetas_samsara: Iterable[dict[str, Any]]
) -> dict[str, str]:
    """Resuelve nombres exactos ignorando mayusculas, espacios y acentos."""
    indice: dict[str, list[dict[str, Any]]] = {}
    for etiqueta in etiquetas_samsara:
        nombre = str(etiqueta.get("name") or "").strip()
        etiqueta_id = str(etiqueta.get("id") or "").strip()
        if nombre and etiqueta_id:
            indice.setdefault(normalizar_texto(nombre), []).append(etiqueta)
    encontrados: dict[str, str] = {}
    faltantes, ambiguas = [], []
    for solicitado in nombres_solicitados:
        coincidencias = indice.get(normalizar_texto(solicitado), [])
        if not coincidencias:
            faltantes.append(str(solicitado))
        elif len(coincidencias) > 1:
            ambiguas.append(str(solicitado))
        else:
            encontrados[str(solicitado)] = str(coincidencias[0]["id"])
    if faltantes or ambiguas:
        partes = []
        if faltantes:
            partes.append("no encontradas: " + ", ".join(faltantes))
        if ambiguas:
            partes.append("duplicadas en Samsara: " + ", ".join(ambiguas))
        raise ConfiguracionError("Etiquetas " + "; ".join(partes))
    return encontrados


def obtener_ids_etiquetas_reporte(
    reporte: dict[str, Any], etiquetas_samsara: list[dict[str, Any]]
) -> tuple[list[str], dict[str, str]]:
    resueltas = resolver_etiquetas(reporte.get("etiquetas") or [], etiquetas_samsara)
    ids = [str(x).strip() for x in reporte.get("etiqueta_ids") or [] if str(x).strip()]
    ids.extend(resueltas.values())
    return list(dict.fromkeys(ids)), resueltas


def resolver_etiquetas_samsara(
    sesion: requests.Session,
    nombres_solicitados: Iterable[str],
    cache: dict[str, dict[str, str]] | None = None,
) -> dict[str, str]:
    """Busca cada tag por el external ID automatico ``samsara.name``."""
    cache = cache if cache is not None else {}
    resultado: dict[str, str] = {}
    for nombre in nombres_solicitados:
        clave = normalizar_texto(nombre)
        if clave not in cache:
            external_id = quote(f"samsara.name:{str(nombre).strip()}", safe=":")
            response = sesion.get(f"{SAMSARA_BASE_URL}/tags/{external_id}", timeout=60)
            if response.status_code == 404:
                raise ConfiguracionError(f"Etiqueta no encontrada en Samsara: {nombre}")
            response.raise_for_status()
            data = response.json().get("data") or {}
            etiqueta_id = str(data.get("id") or "").strip()
            nombre_real = str(data.get("name") or "").strip()
            if not etiqueta_id or normalizar_texto(nombre_real) != clave:
                raise ConfiguracionError(
                    f"Samsara no devolvio una coincidencia exacta para: {nombre}"
                )
            cache[clave] = {"id": etiqueta_id, "name": nombre_real}
        resultado[str(nombre)] = cache[clave]["id"]
    return resultado


def obtener_ids_etiquetas_reporte_remoto(
    sesion: requests.Session,
    reporte: dict[str, Any],
    cache: dict[str, dict[str, str]],
) -> tuple[list[str], dict[str, str]]:
    resueltas = resolver_etiquetas_samsara(sesion, reporte.get("etiquetas") or [], cache)
    ids = [str(x).strip() for x in reporte.get("etiqueta_ids") or [] if str(x).strip()]
    ids.extend(resueltas.values())
    return list(dict.fromkeys(ids)), resueltas


def obtener_geocercas_de_etiquetas(
    sesion: requests.Session, etiqueta_ids: Iterable[str]
) -> dict[str, str]:
    geocercas: dict[str, str] = {}
    for etiqueta_id in dict.fromkeys(str(x) for x in etiqueta_ids if x):
        response = sesion.get(f"{SAMSARA_BASE_URL}/tags/{etiqueta_id}", timeout=60)
        response.raise_for_status()
        data = response.json().get("data") or {}
        for address in data.get("addresses") or []:
            address_id = str(address.get("id") or "").strip()
            if address_id:
                geocercas[address_id] = str(address.get("name") or "").strip()
    return geocercas


def obtener_vehiculos(
    sesion: requests.Session, etiqueta_ids: list[str], tipo_filtro: str
) -> list[dict[str, Any]]:
    return obtener_paginas(
        sesion,
        "/fleet/vehicles/stats",
        {"types": "gps,engineStates,ecuSpeedMph", tipo_filtro: ",".join(etiqueta_ids)},
    )


def enriquecer_operadores_samsara(
    sesion: requests.Session,
    results: list[dict[str, Any]],
    now_mx: datetime,
    ventana_horas: int = 24,
) -> list[dict[str, Any]]:
    """Asigna el operador no pasajero mas reciente reportado por Samsara."""
    vehiculo_ids = [str(row.get("SamsaraVehicleId") or "") for row in results]
    vehiculo_ids = list(dict.fromkeys(x for x in vehiculo_ids if x))
    for row in results:
        row.setdefault("Operador", "")
        row.setdefault("ID Operador", "")
    if not vehiculo_ids:
        return results

    fin = now_mx.astimezone(pytz.UTC)
    inicio = fin - timedelta(hours=max(1, int(ventana_horas)))
    asignaciones = obtener_paginas(
        sesion,
        "/fleet/driver-vehicle-assignments",
        {
            "filterBy": "vehicles",
            "vehicleIds": ",".join(vehiculo_ids),
            "startTime": inicio.isoformat(),
            "endTime": fin.isoformat(),
        },
    )
    recientes: dict[str, tuple[datetime, dict[str, Any]]] = {}
    for asignacion in asignaciones:
        if asignacion.get("isPassenger"):
            continue
        vehiculo_id = str((asignacion.get("vehicle") or {}).get("id") or "")
        conductor = asignacion.get("driver") or {}
        inicio_asignacion = asignacion.get("startTime")
        if not vehiculo_id or not conductor.get("name") or not inicio_asignacion:
            continue
        fecha = dp.parse(inicio_asignacion)
        if vehiculo_id not in recientes or fecha > recientes[vehiculo_id][0]:
            recientes[vehiculo_id] = (fecha, conductor)

    encontrados = 0
    for row in results:
        asignacion = recientes.get(str(row.get("SamsaraVehicleId") or ""))
        if not asignacion:
            continue
        conductor = asignacion[1]
        row["Operador"] = str(conductor.get("name") or "")
        row["ID Operador"] = str(conductor.get("id") or "")
        encontrados += 1
    print(f"[Samsara] Operadores vinculados: {encontrados}/{len(results)}")
    return results


def obtener_datos_google_sheets(
    results: list[dict[str, Any]], fecha_busqueda: datetime, settings: dict[str, Any]
) -> list[dict[str, Any]]:
    """Enriquece placas, identificador y operador desde la planeacion diaria."""
    for row in results:
        row.setdefault("ID ROSTERING", "")
        row.setdefault("Operador", "")
        row.setdefault("PLACAS", "")
        row.setdefault("ORIGEN", "")
        row.setdefault("DESTINO", "")
    if not results:
        return results

    credentials = BASE_DIR / settings.get("credenciales", "Credenciales.json")
    sheet_url = settings.get(
        "url",
        "https://docs.google.com/spreadsheets/d/1zbVe4Rk7aGaC_gyy0n2ik5VEWzn3w6XAyn91LNp2cMA/edit#gid=0",
    )
    gc = gspread.service_account(filename=credentials)
    workbook = gc.open_by_url(sheet_url)
    meses = {
        1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
        5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
        9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
    }
    candidatos = [
        f"{meses[fecha_busqueda.month]} {str(fecha_busqueda.year)[-2:]}",
        f"{meses[fecha_busqueda.month].title()} {fecha_busqueda.year}",
    ]
    hojas = {normalizar_texto(ws.title): ws.title for ws in workbook.worksheets()}
    nombre_hoja = next(
        (hojas[normalizar_texto(x)] for x in candidatos if normalizar_texto(x) in hojas),
        None,
    )
    if not nombre_hoja:
        logging.warning("No se encontro hoja mensual. Candidatas: %s", candidatos)
        return results

    df = get_as_dataframe(workbook.worksheet(nombre_hoja), evaluate_formulas=True).fillna("")
    indice_columnas = {normalizar_texto(columna): columna for columna in df.columns}

    def encontrar_columna(*candidatos: str) -> str | None:
        return next(
            (indice_columnas[normalizar_texto(x)] for x in candidatos if normalizar_texto(x) in indice_columnas),
            None,
        )

    columna_fecha = encontrar_columna("FECHA DE INICIO", "FECHA")
    columna_unidad = encontrar_columna("UNIDAD")
    columna_operador = encontrar_columna("OPERADOR", "OP")
    columna_roster = encontrar_columna("ROSTERING ID", "ROSTERING \nID", "ID OP 1", "ROSTER")
    columna_placas = encontrar_columna("PLACAS")
    if not columna_fecha or not columna_unidad:
        logging.warning("Faltan columnas de fecha o unidad en la hoja %s", nombre_hoja)
        return results

    meses_es_en = {
        "ene": "jan", "feb": "feb", "mar": "mar", "abr": "apr",
        "may": "may", "jun": "jun", "jul": "jul", "ago": "aug",
        "sep": "sep", "sept": "sep", "set": "sep", "oct": "oct",
        "nov": "nov", "dic": "dec",
    }

    def normalizar_fecha(valor: Any):
        texto = str(valor).strip().lower()
        for es, en in meses_es_en.items():
            texto = re.sub(rf"\b{es}\b", en, texto)
        try:
            return dp.parse(texto, dayfirst=True).date()
        except Exception:
            return None

    df["_fecha_norm"] = df[columna_fecha].apply(normalizar_fecha)
    filas = df[df["_fecha_norm"] == fecha_busqueda.date()]
    operadores_encontrados = 0
    for row in results:
        unidad = str(row.get("Unidad") or "").strip()
        coincidencias = filas[filas[columna_unidad].astype(str).str.strip() == unidad]
        if coincidencias.empty:
            continue
        if columna_operador:
            con_operador = coincidencias[
                coincidencias[columna_operador].astype(str).str.strip() != ""
            ]
            fila = (con_operador if not con_operador.empty else coincidencias).iloc[-1]
            row["Operador"] = fila.get(columna_operador, "")
            if str(row["Operador"]).strip():
                operadores_encontrados += 1
        else:
            fila = coincidencias.iloc[-1]
        row["ID ROSTERING"] = fila.get(columna_roster, "") if columna_roster else ""
        row["PLACAS"] = fila.get(columna_placas, "") if columna_placas else ""
    logging.info("Operadores vinculados desde Sheets: %s/%s", operadores_encontrados, len(results))
    print(f"[Sheets] Operadores vinculados: {operadores_encontrados}/{len(results)}")
    return results


def procesar_vehiculos(
    vehicles: Iterable[dict[str, Any]],
    now_mx: datetime,
    filtros: dict[str, Any],
    geocercas_excluidas: dict[str, str],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    results, excluidas = [], []
    gps_max_minutos = filtros.get("gps_max_minutos", 60)
    incluir_todas = filtros.get("incluir_todas_las_unidades", False)
    especiales = {str(x) for x in filtros.get("geocercas_especiales_ids") or []}
    for u in vehicles:
        unidad_id = str(u.get("id") or "")
        unidad = str(u.get("name") or "Sin nombre")
        try:
            gps = u.get("gps") or {}
            gps_time = gps.get("time")
            loc_time = dp.parse(gps_time).astimezone(now_mx.tzinfo) if gps_time else None
            antiguedad = int((now_mx - loc_time).total_seconds() / 60) if loc_time else None
            address = gps.get("address") or {}
            geocerca_id = str(address.get("id") or "")
            geocerca = str(address.get("name") or "").strip()
            speed = float(gps.get("speedMilesPerHour") or 0)
            ecu = bool(gps.get("isEcuSpeed", False))
            lat, lon = gps.get("latitude", ""), gps.get("longitude", "")
            coordenadas = f"{lat},{lon}" if lat != "" and lon != "" else ""
            motivo, detalle = "", ""
            if gps_max_minutos is not None and (antiguedad is None or antiguedad > int(gps_max_minutos)):
                motivo = "GPS SIN FECHA" if loc_time is None else "GPS VIEJO"
                detalle = "GPS sin marca de tiempo" if loc_time is None else f"Antiguedad: {antiguedad} minutos"
            elif geocerca_id in especiales:
                motivo, detalle = "GEOCERCA ESPECIAL", "El ID esta configurado como geocerca especial"
            elif geocerca_id in geocercas_excluidas:
                motivo, detalle = "PATIO/GEOCERCA EXCLUIDA", "La geocerca pertenece a una etiqueta excluida"
                geocerca = geocercas_excluidas[geocerca_id] or geocerca
            elif filtros.get("excluir_speed_cero_sin_ecu", True) and speed == 0 and not ecu:
                motivo, detalle = "SPEED 0 SIN ECU", "speedMilesPerHour es 0 e isEcuSpeed es False"
            if motivo and not incluir_todas:
                excluidas.append({
                    "Unidad": unidad, "SamsaraVehicleId": unidad_id, "Motivo": motivo,
                    "Detalle": detalle, "GpsTimeMexico": loc_time,
                    "AntiguedadMinutos": antiguedad, "GeocercaId": geocerca_id,
                    "Geocerca": geocerca, "Latitud": lat, "Longitud": lon,
                    "Coordenadas": coordenadas, "VelocidadMph": speed, "IsEcuSpeed": ecu,
                })
                continue
            location = (gps.get("reverseGeo") or {}).get("formattedLocation", "")
            results.append({
                "Unidad": unidad, "SamsaraVehicleId": unidad_id, "GpsActual": gps,
                "Fecha GPS": loc_time, "Ubicación": location,
                "Estatus": "DETENIDO" if speed == 0 and ecu else "RUTA",
                "Velocidad Mph": speed, "IsEcuSpeed": ecu,
                "Latitud": lat, "Longitud": lon, "Coordenadas": coordenadas,
                "Geocerca": geocerca,
                "Minutos Detenido": None, "Tiempo Detenido": None,
                "Detenido Desde": None, "Ventana Detenido": "",
                "Minutos Trafico": None, "Tiempo Trafico": None,
                "Trafico Desde": None, "Motor": "", "EcuSpeedActual": None,
            })
        except Exception as error:
            logging.exception("Error procesando unidad %s", unidad)
            if incluir_todas:
                results.append({
                    "Unidad": unidad, "SamsaraVehicleId": unidad_id,
                    "Operador": "", "Ubicación": "", "Coordenadas": "",
                })
            else:
                excluidas.append({"Unidad": unidad, "SamsaraVehicleId": unidad_id, "Motivo": "ERROR", "Detalle": str(error)})
    return results, excluidas


def filtrar_contenido(results: list[dict[str, Any]], contenido: dict[str, Any]):
    estados = {normalizar_texto(x) for x in contenido.get("estados") or []}
    unidades = {normalizar_texto(x) for x in contenido.get("unidades") or []}
    filtrados = [
        row for row in results
        if (not estados or normalizar_texto(row.get("Estatus")) in estados)
        and (not unidades or normalizar_texto(row.get("Unidad")) in unidades)
    ]
    orden_estados = {
        normalizar_texto(estado): indice
        for indice, estado in enumerate(contenido.get("orden_estados") or [])
    }
    if orden_estados:
        def clave_unidad(row: dict[str, Any]):
            unidad = str(row.get("Unidad") or "")
            return (0, int(unidad)) if unidad.isdigit() else (1, normalizar_texto(unidad))

        filtrados.sort(key=lambda row: (
            orden_estados.get(normalizar_texto(row.get("Estatus")), len(orden_estados)),
            clave_unidad(row),
        ))
    return filtrados


def estatus_chat(estatus: Any) -> str:
    """Estatus simplificado para Chat; el detalle de la fuente queda en Excel y consola."""
    texto = str(estatus or "").strip()
    return "DETENIDO" if texto.startswith("DETENIDO") else texto


def construir_reporte_google(results: list[dict[str, Any]], now_mx: datetime, titulo: str):
    conteos: dict[str, int] = {}
    for row in results:
        estado = estatus_chat(row.get("Estatus")) or "SIN ESTATUS"
        conteos[estado] = conteos.get(estado, 0) + 1
    lineas = [
        f"🚛 *{titulo}*", "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
        f"📅 *Fecha:* {now_mx:%Y-%m-%d}", f"🕒 *Hora:* {now_mx:%H:%M:%S}",
        f"📦 *Total unidades:* {len(results)}", "", "📊 *Resumen*",
    ]
    iconos = {
        "DETENIDO": "⛔",
        "REVISAR": "🔎",
        "TRAFICO LENTO": "🚦",
        "RETEN": "🚧",
        "RUTA": "✅",
    }
    for estado in iconos:
        lineas.append(f"{iconos[estado]} {estado.title()}: {conteos.get(estado, 0)}")
    lineas.extend(["", "*Detalle:*", "```"])
    if not results:
        lineas.append("No se encontraron unidades para reportar.")
    else:
        lineas.append(f"{'UNIDAD':<12} | {'ESTATUS':<13} | {'TIEMPO':<12} | {'COORDENADAS':<23} | UBICACION")
        lineas.append("-" * 124)
        # Dentro de DETENIDO se conserva primero lo confirmado por ambas telemetrias.
        orden = {
            "DETENIDO CONFIRMADO": 0,
            "DETENIDO SAMSARA": 1,
            "DETENIDO LOGITRACK": 2,
            "DETENIDO": 3,
            "REVISAR": 4,
            "TRAFICO LENTO": 5,
            "RETEN": 6,
            "RUTA": 7,
        }
        for row in sorted(results, key=lambda x: (orden.get(x.get("Estatus"), 99), str(x.get("Unidad")))):
            estado = estatus_chat(row.get("Estatus"))
            tiempo = (
                row.get("Tiempo Detenido")
                if estado == "DETENIDO"
                else row.get("Tiempo Trafico")
            )
            ubicacion = str(row.get("Ubicación") or "").replace("\n", " ")[:62]
            coordenadas = str(row.get("Coordenadas") or "")[:23]
            lineas.append(f"{str(row.get('Unidad') or '')[:12]:<12} | {estado[:13]:<13} | {str(tiempo or '')[:12]:<12} | {coordenadas:<23} | {ubicacion}")
    lineas.extend(["```", "", "✅ *Reporte generado automáticamente*"])
    return "\n".join(lineas)


def imprimir_seguro(texto: str) -> None:
    """Imprime una vista previa aun si la consola de Windows no admite emojis."""
    try:
        print(texto)
    except UnicodeEncodeError:
        encoding = sys.stdout.encoding or "ascii"
        print(texto.encode(encoding, errors="replace").decode(encoding))


def dividir_mensaje(texto: str, max_chars: int = MAX_GOOGLE_CHAT_CHARS) -> list[str]:
    if len(texto) <= max_chars:
        return [texto]
    partes, actual, longitud = [], [], 0
    for linea in texto.splitlines():
        costo = len(linea) + 1
        if actual and longitud + costo > max_chars:
            partes.append("\n".join(actual)); actual, longitud = [], 0
        actual.append(linea); longitud += costo
    if actual:
        partes.append("\n".join(actual))
    return partes


def enviar_google_chat(texto: str, webhook_url: str, sesion=requests) -> None:
    for parte in dividir_mensaje(texto):
        response = sesion.post(webhook_url, json={"text": parte}, timeout=30)
        response.raise_for_status()


DEFAULT_COLUMNS = [
    "Unidad", "Estatus", "Tiempo Detenido", "Tiempo Trafico", "Motor",
    "Fecha GPS", "Ubicación", "Coordenadas", "Geocerca", "Velocidad Mph",
    "IsEcuSpeed", "PLACAS", "ID ROSTERING",
]


def normalizar_valor_excel(valor: Any) -> Any:
    if isinstance(valor, datetime):
        return valor.replace(tzinfo=None)
    return valor if isinstance(valor, (str, int, float, bool)) or valor is None else str(valor)


def aplicar_estilo_hoja(ws, fila_encabezado: int) -> None:
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = f"A{fila_encabezado + 1}"
    for cell in ws[fila_encabezado]:
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for columna in range(1, ws.max_column + 1):
        max_len = max((len(str(ws.cell(fila, columna).value or "")) for fila in range(1, ws.max_row + 1)), default=0)
        ws.column_dimensions[get_column_letter(columna)].width = min(max(max_len + 2, 12), 55)
    if ws.max_row > fila_encabezado:
        ws.auto_filter.ref = f"A{fila_encabezado}:{get_column_letter(ws.max_column)}{ws.max_row}"
        for row in ws.iter_rows(min_row=fila_encabezado + 1):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                if cell.row % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor="D9EAF7")


def crear_excel_reporte(
    nombre: str, results: list[dict[str, Any]], excluidas: list[dict[str, Any]],
    now_mx: datetime, contenido: dict[str, Any],
) -> bytes:
    wb = Workbook()
    ws = wb.active; ws.title = "Unidades"
    columnas = contenido.get("columnas") or DEFAULT_COLUMNS
    ws.append([nombre]); ws.append(["Generado", now_mx.replace(tzinfo=None)])
    if contenido.get("resumen_estados_excel", False):
        conteos = {
            estado: sum(1 for row in results if normalizar_texto(row.get("Estatus")) == normalizar_texto(estado))
            for estado in ("RUTA", "DETENIDO")
        }
        ws.append(["Total", len(results), "En ruta", conteos["RUTA"], "Detenidos", conteos["DETENIDO"]])
        for referencia in ("A3", "C3", "E3"):
            ws[referencia].font = Font(bold=True, color="17365D")
    else:
        ws.append(["Total", len(results)])
    ws.append([]); ws.append(columnas)
    for row in results:
        ws.append([normalizar_valor_excel(row.get(columna)) for columna in columnas])
    aplicar_estilo_hoja(ws, 5); ws["B2"].number_format = "yyyy-mm-dd hh:mm:ss"

    if contenido.get("incluir_resumen_excel", True):
        resumen = wb.create_sheet("Resumen", 0)
        resumen.append(["Resumen del reporte", nombre]); resumen.append(["Generado", now_mx.replace(tzinfo=None)])
        resumen.append(["Unidades incluidas", len(results)]); resumen.append(["Unidades omitidas", len(excluidas)])
        estados: dict[str, int] = {}
        for row in results:
            estado = str(row.get("Estatus") or "SIN ESTATUS"); estados[estado] = estados.get(estado, 0) + 1
        for estado, cantidad in sorted(estados.items()):
            resumen.append([estado, cantidad])
        aplicar_estilo_hoja(resumen, 1); resumen["B2"].number_format = "yyyy-mm-dd hh:mm:ss"

    if contenido.get("incluir_omitidas_excel", True):
        omitidas_ws = wb.create_sheet("Omitidas")
        columnas_omitidas = ["Unidad", "SamsaraVehicleId", "Motivo", "Detalle", "GpsTimeMexico", "AntiguedadMinutos", "Geocerca", "GeocercaId", "Latitud", "Longitud", "Coordenadas", "VelocidadMph", "IsEcuSpeed"]
        omitidas_ws.append(columnas_omitidas)
        for row in excluidas:
            omitidas_ws.append([normalizar_valor_excel(row.get(c)) for c in columnas_omitidas])
        aplicar_estilo_hoja(omitidas_ws, 1)
    salida = BytesIO(); wb.save(salida)
    return salida.getvalue()


def obtener_destinatarios(entrega: dict[str, Any]) -> list[str]:
    destinos = [str(x).strip() for x in entrega.get("destinatarios") or [] if str(x).strip()]
    env_name = str(entrega.get("destinatarios_env") or "").strip()
    if env_name:
        destinos.extend(x.strip() for x in os.getenv(env_name, "").split(",") if x.strip())
    return list(dict.fromkeys(destinos))


def enviar_correo_excel(destinatarios: list[str], asunto: str, cuerpo: str, nombre_archivo: str, excel: bytes) -> None:
    host, port = os.getenv("SMTP_HOST", "smtp.gmail.com"), int(os.getenv("SMTP_PORT", "587"))
    user, password = os.getenv("SMTP_USER", ""), os.getenv("SMTP_PASSWORD", "")
    from_address = os.getenv("EMAIL_FROM_ADDRESS", "") or user
    from_name = os.getenv("EMAIL_FROM_NAME", "Reportes de telemetria")
    if not user or not password or not from_address:
        raise ConfiguracionError("Faltan SMTP_USER/SMTP_PASSWORD/EMAIL_FROM_ADDRESS")
    if not destinatarios:
        raise ConfiguracionError("La entrega por correo no tiene destinatarios")
    msg = EmailMessage(); msg["From"] = f"{from_name} <{from_address}>"
    msg["To"] = ", ".join(destinatarios); msg["Subject"] = asunto; msg.set_content(cuerpo)
    msg.add_attachment(excel, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=nombre_archivo)
    smtp_cls = smtplib.SMTP_SSL if port == 465 else smtplib.SMTP
    with smtp_cls(host, port, timeout=60) as smtp:
        if port != 465:
            smtp.starttls()
        smtp.login(user, password); smtp.send_message(msg)


def formatear_plantilla(texto: str, reporte: dict[str, Any], now_mx: datetime) -> str:
    return texto.format(nombre=reporte["nombre"], equipo=reporte.get("equipo", reporte["nombre"]), fecha=now_mx.strftime("%Y-%m-%d"), hora=now_mx.strftime("%H:%M:%S"))


def seleccionar_entregas(reporte: dict[str, Any], canal_forzado: str | None = None) -> list[dict[str, Any]]:
    entregas = reporte.get("entregas") or []
    if canal_forzado:
        seleccionadas = [
            entrega for entrega in entregas
            if str(entrega.get("canal", "")).lower() == canal_forzado.lower()
        ]
        if not seleccionadas:
            raise ConfiguracionError(
                f"El reporte {reporte['nombre']} no tiene configurado el canal {canal_forzado}"
            )
        return seleccionadas
    return [entrega for entrega in entregas if entrega.get("activo", True)]


def entregar_reporte(
    reporte, results, excluidas, now_mx, dry_run, canal_forzado=None, chat_prueba=False
):
    contenido = reporte.get("contenido") or {}
    entregas = seleccionar_entregas(reporte, "google_chat" if chat_prueba else canal_forzado)
    if not entregas:
        print(f"[{reporte['nombre']}] Sin entregas activas; no se envio nada."); return
    excel_cache = None
    for entrega in entregas:
        canal = entrega["canal"].lower()
        if canal == "google_chat":
            titulo = f"🧪 PRUEBA - {reporte['nombre']}" if chat_prueba else reporte["nombre"]
            mensaje = construir_reporte_google(results, now_mx, titulo)
            if dry_run:
                imprimir_seguro(
                    f"\n[DRY-RUN][Google Chat][{reporte['nombre']}]\n{mensaje}"
                )
                continue
            env_name = (
                WEBHOOK_PRUEBAS_ENV if chat_prueba
                else entrega.get("webhook_env", "GOOGLE_CHAT_WEBHOOK_URL")
            )
            webhook = os.getenv(env_name, "")
            if not webhook:
                raise ConfiguracionError(f"Falta la variable {env_name}")
            enviar_google_chat(mensaje, webhook)
            print(f"[{reporte['nombre']}] Enviado a Google Chat ({env_name}).")
        elif canal == "correo":
            if excel_cache is None:
                excel_cache = crear_excel_reporte(reporte["nombre"], results, excluidas, now_mx, contenido)
            filename = f"{slug(reporte['nombre'])}_{now_mx:%Y%m%d_%H%M%S}.xlsx"
            if dry_run:
                destino = BASE_DIR / "outputs" / "envio_previews" / filename
                destino.parent.mkdir(parents=True, exist_ok=True); destino.write_bytes(excel_cache)
                load_workbook(BytesIO(excel_cache), read_only=True).close()
                print(f"[DRY-RUN][Correo] Excel generado: {destino}"); continue
            destinatarios = obtener_destinatarios(entrega)
            asunto = formatear_plantilla(entrega.get("asunto", "{nombre} - {fecha} {hora}"), reporte, now_mx)
            cuerpo = formatear_plantilla(entrega.get("cuerpo", "Se adjunta el reporte {nombre}."), reporte, now_mx)
            enviar_correo_excel(destinatarios, asunto, cuerpo, filename, excel_cache)
            print(f"[{reporte['nombre']}] Enviado por correo a {len(destinatarios)} destinatario(s).")


def combinar_filtros(base, reporte, perfiles=None):
    resultado = deepcopy(base)
    perfil = reporte.get("perfil_filtros")
    if perfil:
        resultado.update(deepcopy((perfiles or {}).get(perfil) or {}))
    resultado.update(reporte.get("filtros") or {})
    return resultado


def seleccionar_reportes(config, nombres):
    activos = [r for r in config["reportes"] if r.get("activo", True)]
    if not nombres:
        return activos
    solicitados = {normalizar_texto(x) for x in nombres}
    seleccionados = [
        r for r in config["reportes"]
        if normalizar_texto(r["nombre"]) in solicitados
    ]
    faltantes = solicitados - {normalizar_texto(r["nombre"]) for r in seleccionados}
    if faltantes:
        raise ConfiguracionError("Reportes no encontrados: " + ", ".join(sorted(faltantes)))
    return seleccionados


def ejecutar(args: argparse.Namespace) -> int:
    if not 1 <= args.tag_page_size <= 512:
        raise ConfiguracionError("--tag-page-size debe estar entre 1 y 512")
    if getattr(args, "chat_prueba", False) and args.canal == "correo":
        raise ConfiguracionError("--chat-prueba solo envia por Google Chat; no use --canal correo")
    config = cargar_configuracion(args.config)
    reportes = seleccionar_reportes(config, args.solo)
    token = os.getenv("SAMSARA_API_TOKEN") or os.getenv("SAMSARA_TOKEN", "").removeprefix("Bearer ")
    if not token:
        raise ConfiguracionError("Falta SAMSARA_API_TOKEN o SAMSARA_TOKEN")
    sesion = crear_sesion_samsara(token)
    catalogo_path = Path(config.get("catalogo_etiquetas", DEFAULT_TAG_CATALOG_PATH))
    if not catalogo_path.is_absolute():
        catalogo_path = BASE_DIR / catalogo_path
    if args.sincronizar_etiquetas:
        catalogo = guardar_catalogo_etiquetas(sesion, catalogo_path, args.tag_page_size)
        print(
            f"Catalogo generado: {catalogo_path} | "
            f"Etiquetas={catalogo['totalEtiquetas']} Padres={catalogo['totalEtiquetasPadre']}"
        )
        return 0
    if args.listar_etiquetas:
        tags = obtener_catalogo_etiquetas_samsara(sesion, args.tag_page_size)
        for tag in sorted(tags, key=lambda x: normalizar_texto(x.get("nombre"))):
            print(f"{tag.get('id')}\t{tag.get('parentTagId')}\t{tag.get('nombre')}")
        return 0
    tz = pytz.timezone(config.get("zona_horaria", DEFAULT_TIMEZONE)); now_mx = datetime.now(tz)
    filtros_base = config.get("filtros_base") or {}
    cache_geocercas: dict[tuple[str, ...], dict[str, str]] = {}
    direcciones_samsara: list[dict[str, Any]] | None = None
    cache_etiquetas = cargar_cache_catalogo(catalogo_path)
    errores = []
    for reporte in reportes:
        try:
            ids, resueltas = obtener_ids_etiquetas_reporte_remoto(
                sesion, reporte, cache_etiquetas
            )
            print(f"\n[{reporte['nombre']}] Etiquetas: {resueltas or ids}")
            vehicles = obtener_vehiculos(sesion, ids, reporte.get("tipo_filtro_etiqueta", "tagIds"))
            filtros = combinar_filtros(
                filtros_base, reporte, config.get("perfiles_filtros") or {}
            )
            excluir_geocercas = filtros.get("excluir_unidades_en_geocerca", False)
            ids_geocercas = []
            if excluir_geocercas:
                geocercas_por_nombre = resolver_etiquetas_samsara(
                    sesion, filtros.get("etiquetas_geocercas_excluidas") or [], cache_etiquetas
                )
                ids_geocercas.extend(geocercas_por_nombre.values())
                ids_geocercas.extend(
                    str(x) for x in filtros.get("etiqueta_ids_geocercas_excluidas") or []
                )
            ids_geocercas = tuple(dict.fromkeys(ids_geocercas))
            if ids_geocercas not in cache_geocercas:
                cache_geocercas[ids_geocercas] = obtener_geocercas_de_etiquetas(sesion, ids_geocercas)
            results, excluidas = procesar_vehiculos(
                vehicles, now_mx, filtros, cache_geocercas[ids_geocercas]
            )
            logitrack_settings = reporte.get("telemetria_logitrack") or {}
            usar_logitrack = logitrack_settings.get("activo", False)
            distancia_cercania = filtros.get("distancia_cercania_geocerca_metros")
            # None = no se pudo descargar la geometria; las reglas que la usan no se aplican.
            geometrias: dict[str, dict[str, Any]] | None = {}
            especiales_ids = tuple(str(x) for x in filtros.get("geocercas_especiales_ids") or [])
            if (ids_geocercas or especiales_ids) and not filtros.get("incluir_todas_las_unidades"):
                try:
                    if direcciones_samsara is None:
                        direcciones_samsara = obtener_paginas(sesion, "/addresses", {"limit": 512})
                    geometrias = obtener_geometrias_geocercas(
                        sesion, ids_geocercas, especiales_ids, direcciones_samsara
                    )
                except Exception:
                    logging.exception("No se pudieron descargar las geometrias de geocercas")
                    print("[Geocercas][ERROR] No se pudo descargar la geometria de las geocercas.")
                    geometrias = None
            if geometrias:
                results, dentro = excluir_por_coordenadas_en_geocerca(results, geometrias)
                excluidas.extend(dentro)
            if usar_logitrack:
                statuses_logitrack: list[dict[str, Any]] = []
                try:
                    api_url = os.getenv(
                        "LOGITRACK_API_URL", LOGITRACK_DEFAULT_API_URL
                    ).strip().rstrip("/")
                    token_url = os.getenv(
                        "LOGITRACK_TOKEN_URL", LOGITRACK_DEFAULT_TOKEN_URL
                    ).strip()
                    timeout_logitrack = int(os.getenv("LOGITRACK_TIMEOUT_SECONDS", "60"))
                    with requests.Session() as sesion_logitrack:
                        token_logitrack = request_access_token(
                            sesion_logitrack, token_url, timeout_logitrack
                        )
                        statuses_logitrack = fetch_last_status(
                            sesion_logitrack,
                            api_url,
                            token_logitrack,
                            timeout_logitrack,
                        )
                except Exception:
                    logging.exception("No se pudo consultar Logitrack")
                    print(
                        f"[Logitrack][ERROR] No se pudo consultar; "
                        f"{reporte['nombre']} continuara solo con Samsara."
                    )
                results = enriquecer_con_logitrack(
                    results, statuses_logitrack, now_mx, logitrack_settings
                )
                rescatadas, excluidas = rescatar_con_logitrack(
                    excluidas, statuses_logitrack, now_mx, logitrack_settings, geometrias
                )
                results.extend(rescatadas)
            sheets_settings = config.get("google_sheets") or {}
            if sheets_settings.get("activo", True) and reporte.get("enriquecer_google_sheets", True):
                try:
                    results = obtener_datos_google_sheets(results, now_mx, sheets_settings)
                except Exception:
                    logging.exception("No se pudo enriquecer desde Google Sheets")
            if reporte.get("enriquecer_operadores_samsara", False):
                try:
                    results = enriquecer_operadores_samsara(
                        sesion,
                        results,
                        now_mx,
                        reporte.get("operadores_ventana_horas", 24),
                    )
                except Exception:
                    logging.exception("No se pudo enriquecer operadores desde Samsara")
            if reporte.get("analizar_detenciones", True):
                try:
                    results = enriquecer_minutos_detenido(results, token, now_mx)
                except Exception:
                    logging.exception("No se pudo enriquecer detenciones")
            if usar_logitrack:
                results = finalizar_comprobacion_telemetria(
                    results, logitrack_settings
                )
            if usar_logitrack:
                imprimir_seguro("\n".join(resumen_telemetria(results)).strip())
            if distancia_cercania and geometrias:
                results, cercanas = omitir_detenidas_cerca_de_geocerca(
                    results, geometrias, float(distancia_cercania),
                    filtros.get("estados_omitir_cerca_geocerca"),
                )
                excluidas.extend(cercanas)
            results = filtrar_contenido(results, reporte.get("contenido") or {})
            print(f"[{reporte['nombre']}] Recibidas={len(vehicles)} Incluidas={len(results)} Omitidas={len(excluidas)}")
            entregar_reporte(
                reporte, results, excluidas, now_mx, args.dry_run, args.canal,
                chat_prueba=args.chat_prueba,
            )
        except Exception as error:
            logging.exception("Fallo el reporte %s", reporte["nombre"])
            errores.append(f"{reporte['nombre']}: {error}"); print(f"[ERROR] [{reporte['nombre']}] {error}")
    if errores:
        print("\nReportes con error:"); [print(f"- {error}") for error in errores]; return 1
    return 0


def crear_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG_PATH)
    parser.add_argument("--solo", action="append", help="Ejecuta solo este reporte; se puede repetir.")
    parser.add_argument(
        "--canal",
        choices=("google_chat", "correo"),
        help="Fuerza un canal configurado, aunque su entrega este inactiva en el JSON.",
    )
    parser.add_argument(
        "--chat-prueba",
        action="store_true",
        help=f"Envia solo por Google Chat al webhook de pruebas ({WEBHOOK_PRUEBAS_ENV}).",
    )
    parser.add_argument("--dry-run", action="store_true", help="Previsualiza sin enviar.")
    parser.add_argument("--listar-etiquetas", action="store_true", help="Lista ID y nombre de etiquetas.")
    parser.add_argument("--sincronizar-etiquetas", action="store_true", help="Genera el catalogo local de padres e hijos.")
    parser.add_argument("--tag-page-size", type=int, default=10, help="Tamano de pagina para descargar tags (1-512).")
    return parser


def main() -> None:
    try:
        sys.exit(ejecutar(crear_parser().parse_args()))
    except Exception as error:
        logging.exception("Error fatal"); print(f"[ERROR] {error}"); sys.exit(1)


if __name__ == "__main__":
    main()
