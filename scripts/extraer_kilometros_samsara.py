"""Extrae kilometros recorridos por unidad usando lecturas acumuladas de Samsara.

La fuente primaria es ``obdOdometerMeters``. Si la lectura ECU no cubre el
periodo, se usa ``gpsDistanceMeters`` como respaldo, conforme a la guia de
Samsara sobre kilometraje y distancia.
"""

from __future__ import annotations

import argparse
import json
import os
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import requests
from dotenv import load_dotenv


REPO_ROOT = Path(__file__).resolve().parents[1]
API_BASE = "https://api.samsara.com"
MX_TZ = ZoneInfo("America/Mexico_City")
TIPOS = "obdOdometerMeters,gpsDistanceMeters"
FUENTES = (
    ("obdOdometerMeters", "ECU"),
    ("gpsDistanceMeters", "GPS"),
)


def iso_utc(fecha: datetime) -> str:
    return fecha.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")


def parsear_fecha(valor: str | None) -> datetime | None:
    if not valor:
        return None
    return datetime.fromisoformat(valor.replace("Z", "+00:00"))


def obtener_paginas(
    sesion: requests.Session,
    ruta: str,
    params: dict[str, Any] | None = None,
    timeout: int = 120,
) -> list[dict[str, Any]]:
    resultado: list[dict[str, Any]] = []
    after: str | None = None
    while True:
        pagina_params = dict(params or {})
        if after:
            pagina_params["after"] = after
        respuesta = sesion.get(
            f"{API_BASE}{ruta}", params=pagina_params, timeout=timeout
        )
        respuesta.raise_for_status()
        payload = respuesta.json()
        resultado.extend(payload.get("data") or [])
        paginacion = payload.get("pagination") or {}
        if not paginacion.get("hasNextPage"):
            return resultado
        after = paginacion.get("endCursor")
        if not after:
            raise RuntimeError(f"{ruta} indico otra pagina sin endCursor")
        time.sleep(0.05)


def catalogo_vehiculos(sesion: requests.Session) -> list[dict[str, Any]]:
    return obtener_paginas(sesion, "/fleet/vehicles", {"limit": 512})


def snapshot(
    sesion: requests.Session, fecha: datetime, vehicle_ids: list[str] | None = None
) -> dict[str, dict[str, Any]]:
    params: dict[str, Any] = {"types": TIPOS, "time": iso_utc(fecha)}
    if vehicle_ids:
        params["vehicleIds"] = ",".join(vehicle_ids)
    datos = obtener_paginas(sesion, "/fleet/vehicles/stats", params)
    return {str(fila["id"]): fila for fila in datos}


def primera_lectura_posterior(
    sesion: requests.Session,
    vehicle_id: str,
    fuente: str,
    inicio: datetime,
    fin: datetime,
) -> dict[str, Any] | None:
    """Busca el primer punto para unidades creadas dentro del periodo.

    Se consultan ventanas crecientes para no descargar meses de lecturas de
    alta frecuencia cuando normalmente el primer dato aparece el primer dia.
    """
    cursor_inicio = inicio
    for dias in (1, 7, 30, 90, 366):
        cursor_fin = min(cursor_inicio + timedelta(days=dias), fin)
        datos = obtener_paginas(
            sesion,
            "/fleet/vehicles/stats/history",
            {
                "types": fuente,
                "vehicleIds": vehicle_id,
                "startTime": iso_utc(cursor_inicio),
                "endTime": iso_utc(cursor_fin),
            },
        )
        puntos: list[dict[str, Any]] = []
        for fila in datos:
            valor = fila.get(fuente)
            if isinstance(valor, list):
                puntos.extend(valor)
            elif isinstance(valor, dict):
                puntos.append(valor)
        puntos = [p for p in puntos if p.get("value") is not None and p.get("time")]
        if puntos:
            return min(puntos, key=lambda p: p["time"])
        if cursor_fin >= fin:
            break
        cursor_inicio = cursor_fin
    return None


def lectura(fila: dict[str, Any] | None, fuente: str) -> dict[str, Any] | None:
    valor = (fila or {}).get(fuente)
    if isinstance(valor, dict) and valor.get("value") is not None:
        return valor
    return None


def calcular_fila(
    vehiculo: dict[str, Any],
    inicial: dict[str, Any] | None,
    final: dict[str, Any] | None,
    inicio_periodo: datetime,
    fin_periodo: datetime,
    sesion: requests.Session,
) -> dict[str, Any]:
    vehicle_id = str(vehiculo.get("id") or "")
    creado = parsear_fecha(vehiculo.get("createdAtTime"))
    candidatos: list[dict[str, Any]] = []

    for campo, metodo in FUENTES:
        lectura_inicial = lectura(inicial, campo)
        lectura_final = lectura(final, campo)
        inicio_efectivo = inicio_periodo
        base_parcial = False

        if lectura_final and not lectura_inicial and creado and creado >= inicio_periodo:
            lectura_inicial = primera_lectura_posterior(
                sesion, vehicle_id, campo, max(creado, inicio_periodo), fin_periodo
            )
            if lectura_inicial:
                inicio_efectivo = parsear_fecha(lectura_inicial.get("time")) or creado
                base_parcial = True

        if not lectura_inicial or not lectura_final:
            continue
        valor_inicial = float(lectura_inicial["value"])
        valor_final = float(lectura_final["value"])
        delta = valor_final - valor_inicial
        if delta < 0:
            continue
        candidatos.append(
            {
                "campo": campo,
                "metodo": metodo,
                "lectura_inicial": lectura_inicial,
                "lectura_final": lectura_final,
                "inicio_efectivo": inicio_efectivo,
                "base_parcial": base_parcial,
                "delta": delta,
            }
        )

    elegido = candidatos[0] if candidatos else None
    ecu = next((x for x in candidatos if x["metodo"] == "ECU"), None)
    gps = next((x for x in candidatos if x["metodo"] == "GPS"), None)
    control_calidad = ""
    if ecu and gps and gps["delta"] > 1000:
        ratio = ecu["delta"] / gps["delta"]
        if ratio < 0.5 or ratio > 1.5:
            elegido = gps
            control_calidad = (
                "Divergencia mayor al 50% entre ECU y GPS; se uso GPS y se "
                "recomienda revisar odometro/asociacion del gateway."
            )
    if elegido and elegido["metodo"] == "ECU" and gps:
        tiempo_ecu = parsear_fecha(elegido["lectura_final"].get("time"))
        tiempo_gps = parsear_fecha(gps["lectura_final"].get("time"))
        ecu_estancado = elegido["delta"] == 0 and gps["delta"] > 1000
        ecu_muy_atras = bool(
            tiempo_ecu and tiempo_gps and tiempo_gps - tiempo_ecu > timedelta(days=1)
        )
        if ecu_estancado or ecu_muy_atras:
            elegido = gps
            control_calidad = (
                "La lectura ECU estaba detenida o atrasada; se uso GPS para "
                "cubrir el periodo."
            )

    tags = ", ".join(
        str(tag.get("name") or "").strip()
        for tag in vehiculo.get("tags") or []
        if tag.get("name")
    )
    fila: dict[str, Any] = {
        "Unidad": str(vehiculo.get("name") or "").strip(),
        "SamsaraVehicleId": vehicle_id,
        "Placa": str(vehiculo.get("licensePlate") or "").strip(),
        "VIN": str(vehiculo.get("vin") or "").strip(),
        "Marca": str(vehiculo.get("make") or "").strip(),
        "Modelo": str(vehiculo.get("model") or "").strip(),
        "Ano": vehiculo.get("year"),
        "Etiquetas": tags,
        "CreadaEnSamsara": vehiculo.get("createdAtTime"),
        "InicioSolicitado": iso_utc(inicio_periodo),
        "FinSolicitado": iso_utc(fin_periodo),
        "Metodo": "SIN DATOS",
        "FuenteAPI": "",
        "FechaLecturaInicial": None,
        "LecturaInicialMetros": None,
        "FechaLecturaFinal": None,
        "LecturaFinalMetros": None,
        "Kilometros": None,
        "KilometrosECU": ecu["delta"] / 1000 if ecu else None,
        "KilometrosGPS": gps["delta"] / 1000 if gps else None,
        "RatioECUvsGPS": (
            ecu["delta"] / gps["delta"]
            if ecu and gps and gps["delta"] > 0
            else None
        ),
        "Cobertura": "Sin lecturas comparables",
        "Observacion": "No existen dos lecturas acumuladas comparables en el periodo.",
    }
    if not elegido:
        return fila

    fila.update(
        {
            "Metodo": elegido["metodo"],
            "FuenteAPI": elegido["campo"],
            "FechaLecturaInicial": elegido["lectura_inicial"].get("time"),
            "LecturaInicialMetros": elegido["lectura_inicial"].get("value"),
            "FechaLecturaFinal": elegido["lectura_final"].get("time"),
            "LecturaFinalMetros": elegido["lectura_final"].get("value"),
            "Kilometros": elegido["delta"] / 1000,
            "Cobertura": "Parcial" if elegido["base_parcial"] else "Periodo completo",
            "Observacion": (
                control_calidad
                or (
                    "Unidad creada durante el periodo; el calculo inicia en su primera lectura."
                    if elegido["base_parcial"]
                    else (
                        "Respaldo GPS usado porque la lectura ECU no cubre el periodo."
                        if elegido["metodo"] == "GPS"
                        else "Odometro ECU, fuente primaria recomendada por Samsara."
                    )
                )
            ),
        }
    )
    return fila


def argumentos() -> argparse.Namespace:
    ahora = datetime.now(MX_TZ)
    parser = argparse.ArgumentParser(description="Kilometros por unidad desde una fecha.")
    parser.add_argument(
        "--inicio",
        default=f"{ahora.year}-01-01T00:00:00-06:00",
        help="Inicio RFC 3339; por defecto, 1 de enero del ano actual en Mexico.",
    )
    parser.add_argument(
        "--fin", default=ahora.isoformat(), help="Fin RFC 3339; por defecto, ahora."
    )
    parser.add_argument(
        "--salida",
        default=str(REPO_ROOT / "outputs" / "kilometros_samsara.json"),
        help="Archivo JSON de salida.",
    )
    return parser.parse_args()


def main() -> None:
    args = argumentos()
    inicio = parsear_fecha(args.inicio)
    fin = parsear_fecha(args.fin)
    if not inicio or not fin or inicio >= fin:
        raise SystemExit("El periodo solicitado no es valido.")

    load_dotenv(REPO_ROOT / ".env")
    token = os.getenv("SAMSARA_API_TOKEN") or os.getenv("SAMSARA_TOKEN")
    if not token:
        raise SystemExit("Falta SAMSARA_API_TOKEN o SAMSARA_TOKEN en .env")

    sesion = requests.Session()
    sesion.headers.update(
        {"Accept": "application/json", "Authorization": f"Bearer {token}"}
    )
    vehiculos = catalogo_vehiculos(sesion)
    inicial = snapshot(sesion, inicio)
    final = snapshot(sesion, fin)

    filas = [
        calcular_fila(
            vehiculo,
            inicial.get(str(vehiculo.get("id"))),
            final.get(str(vehiculo.get("id"))),
            inicio,
            fin,
            sesion,
        )
        for vehiculo in vehiculos
    ]
    filas.sort(key=lambda x: (x["Unidad"].casefold(), x["SamsaraVehicleId"]))
    con_km = [f for f in filas if f["Kilometros"] is not None]
    resumen = {
        "generadoEn": iso_utc(datetime.now(timezone.utc)),
        "inicio": iso_utc(inicio),
        "fin": iso_utc(fin),
        "totalUnidades": len(filas),
        "unidadesConKilometros": len(con_km),
        "unidadesSinDatos": len(filas) - len(con_km),
        "unidadesECU": sum(f["Metodo"] == "ECU" for f in filas),
        "unidadesGPS": sum(f["Metodo"] == "GPS" for f in filas),
        "unidadesCoberturaParcial": sum(f["Cobertura"] == "Parcial" for f in filas),
        "kilometrosTotales": sum(float(f["Kilometros"]) for f in con_km),
    }
    salida = Path(args.salida)
    salida.parent.mkdir(parents=True, exist_ok=True)
    salida.write_text(
        json.dumps({"resumen": resumen, "unidades": filas}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps(resumen, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
