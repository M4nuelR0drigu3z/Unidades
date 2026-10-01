"""Exporta el historico diario de kilometraje de Samsara a un archivo Excel.

El programa no modifica SQL. Descarga ``obdOdometerMeters`` y
``gpsDistanceMeters`` desde la fecha mas antigua disponible, conserva ambas
fuentes por separado y calcula:

* kilometros observados dentro de cada dia;
* kilometros entre la ultima lectura anterior y la primera lectura actual;
* validacion de saltos por fuente, tiempo transcurrido y respaldo GPS/OBD;
* resumen mensual por unidad;
* listado de incidencias para revision.

Samsara recomienda usar OBD como fuente primaria y ``gpsDistanceMeters`` como
respaldo. No se utiliza ``gpsOdometerMeters`` porque puede tener ajustes
manuales.

Ejemplos:

    python scripts/exportar_historico_km_samsara_excel.py
    python scripts/exportar_historico_km_samsara_excel.py --desde 2024-01-01
    python scripts/exportar_historico_km_samsara_excel.py --hasta 2026-08-31
"""

from __future__ import annotations

import argparse
import importlib.util
import os
import sys
import time
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime, time as hora, timedelta, timezone
from pathlib import Path
from typing import Any, Iterable
from zoneinfo import ZoneInfo

import pandas as pd
import requests
from dotenv import load_dotenv


REPO_ROOT = Path(__file__).resolve().parents[1]
API_BASE = "https://api.samsara.com"
TIPOS = "obdOdometerMeters,gpsDistanceMeters"
FUENTES = (
    ("obdOdometerMeters", "OBD"),
    ("gpsDistanceMeters", "GPS"),
)
MAX_FILAS_HOJA = 900_000


def parsear_argumentos() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Descarga el historico de kilometraje de Samsara a Excel."
    )
    parser.add_argument(
        "--desde",
        help=(
            "Fecha inicial YYYY-MM-DD. Si se omite, se usa la fecha de creacion "
            "mas antigua de las unidades en Samsara."
        ),
    )
    parser.add_argument(
        "--hasta",
        help="Ultima fecha incluida YYYY-MM-DD. Por defecto: ayer.",
    )
    parser.add_argument(
        "--salida",
        type=Path,
        default=REPO_ROOT / "outputs" / "historico_kilometros_samsara.xlsx",
        help="Ruta del archivo Excel de salida.",
    )
    parser.add_argument(
        "--zona-horaria",
        default="America/Mexico_City",
        help="Zona horaria usada para separar los dias.",
    )
    parser.add_argument(
        "--ventana-dias",
        type=int,
        default=7,
        help="Dias consultados por peticion historica. Predeterminado: 7.",
    )
    parser.add_argument(
        "--velocidad-maxima-promedio",
        type=float,
        default=120.0,
        help=(
            "Velocidad promedio maxima para aceptar un salto sin evidencia "
            "adicional. Predeterminado: 120 km/h."
        ),
    )
    parser.add_argument(
        "--tolerancia-fuentes-pct",
        type=float,
        default=30.0,
        help="Diferencia porcentual tolerada entre OBD y GPS. Predeterminado: 30.",
    )
    parser.add_argument(
        "--tolerancia-fuentes-km",
        type=float,
        default=50.0,
        help="Tolerancia minima absoluta entre OBD y GPS. Predeterminado: 50 km.",
    )
    parser.add_argument(
        "--incluir-hoy",
        action="store_true",
        help="Incluye el dia actual aunque todavia este incompleto.",
    )
    args = parser.parse_args()
    if args.ventana_dias < 1 or args.ventana_dias > 31:
        parser.error("--ventana-dias debe estar entre 1 y 31")
    if args.velocidad_maxima_promedio <= 0:
        parser.error("--velocidad-maxima-promedio debe ser mayor que cero")
    return args


def fecha_iso(valor: str | None) -> date | None:
    return date.fromisoformat(valor) if valor else None


def parsear_datetime(valor: str | None) -> datetime | None:
    if not valor:
        return None
    return datetime.fromisoformat(valor.replace("Z", "+00:00"))


def iso_utc(valor: datetime) -> str:
    return valor.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")


def fecha_local_a_utc(valor: date, zona: ZoneInfo) -> datetime:
    return datetime.combine(valor, hora.min, tzinfo=zona).astimezone(timezone.utc)


def dormir_por_rate_limit(respuesta: requests.Response, intento: int) -> float:
    retry_after = respuesta.headers.get("Retry-After")
    if retry_after:
        try:
            return max(1.0, float(retry_after))
        except ValueError:
            pass
    return min(60.0, 2.0**intento)


def solicitar_json(
    sesion: requests.Session,
    ruta: str,
    params: dict[str, Any],
    intentos: int = 7,
) -> dict[str, Any]:
    ultima_excepcion: Exception | None = None
    for intento in range(intentos):
        try:
            respuesta = sesion.get(
                f"{API_BASE}{ruta}", params=params, timeout=(20, 180)
            )
            if respuesta.status_code == 429 or respuesta.status_code >= 500:
                espera = dormir_por_rate_limit(respuesta, intento)
                print(
                    f"  Samsara respondio {respuesta.status_code}; "
                    f"nuevo intento en {espera:.0f} s..."
                )
                time.sleep(espera)
                continue
            respuesta.raise_for_status()
            return respuesta.json()
        except (requests.Timeout, requests.ConnectionError) as exc:
            ultima_excepcion = exc
            espera = min(60.0, 2.0**intento)
            print(f"  Error temporal de conexion; nuevo intento en {espera:.0f} s...")
            time.sleep(espera)
    if ultima_excepcion:
        raise RuntimeError("No fue posible consultar Samsara") from ultima_excepcion
    raise RuntimeError("Samsara no respondio correctamente despues de varios intentos")


def obtener_paginas(
    sesion: requests.Session,
    ruta: str,
    params: dict[str, Any] | None = None,
) -> Iterable[list[dict[str, Any]]]:
    after: str | None = None
    while True:
        pagina_params = dict(params or {})
        if after:
            pagina_params["after"] = after
        payload = solicitar_json(sesion, ruta, pagina_params)
        yield payload.get("data") or []
        paginacion = payload.get("pagination") or {}
        if not paginacion.get("hasNextPage"):
            break
        after = paginacion.get("endCursor")
        if not after:
            raise RuntimeError(f"{ruta} indico otra pagina sin proporcionar endCursor")
        time.sleep(0.05)


def catalogo_vehiculos(sesion: requests.Session) -> list[dict[str, Any]]:
    vehiculos: list[dict[str, Any]] = []
    for pagina in obtener_paginas(sesion, "/fleet/vehicles", {"limit": 512}):
        vehiculos.extend(pagina)
    return vehiculos


@dataclass
class LecturasDia:
    cantidad: int = 0
    primera_fecha: datetime | None = None
    primera_valor: float | None = None
    ultima_fecha: datetime | None = None
    ultima_valor: float | None = None

    def agregar(self, fecha: datetime, valor: float) -> None:
        self.cantidad += 1
        if self.primera_fecha is None or fecha < self.primera_fecha:
            self.primera_fecha = fecha
            self.primera_valor = valor
        if self.ultima_fecha is None or fecha > self.ultima_fecha:
            self.ultima_fecha = fecha
            self.ultima_valor = valor

    def delta_km(self) -> float | None:
        if (
            self.primera_valor is None
            or self.ultima_valor is None
            or self.cantidad < 2
        ):
            return None
        return (self.ultima_valor - self.primera_valor) / 1000.0

    def horas(self) -> float | None:
        if self.primera_fecha is None or self.ultima_fecha is None:
            return None
        return (self.ultima_fecha - self.primera_fecha).total_seconds() / 3600.0


def extraer_historico(
    sesion: requests.Session,
    desde: date,
    hasta_inclusivo: date,
    zona: ZoneInfo,
    ventana_dias: int,
) -> dict[tuple[str, date], dict[str, Any]]:
    dias: dict[tuple[str, date], dict[str, Any]] = {}
    cursor = desde
    fin_exclusivo = hasta_inclusivo + timedelta(days=1)
    numero_ventana = 0

    while cursor < fin_exclusivo:
        fin = min(cursor + timedelta(days=ventana_dias), fin_exclusivo)
        numero_ventana += 1
        print(f"Ventana {numero_ventana}: {cursor} a {fin - timedelta(days=1)}")
        params = {
            "types": TIPOS,
            "startTime": iso_utc(fecha_local_a_utc(cursor, zona)),
            "endTime": iso_utc(fecha_local_a_utc(fin, zona)),
        }
        for pagina in obtener_paginas(
            sesion, "/fleet/vehicles/stats/history", params
        ):
            for vehiculo in pagina:
                vehicle_id = str(vehiculo.get("id") or "")
                if not vehicle_id:
                    continue
                nombre = str(vehiculo.get("name") or "").strip()
                for campo, fuente in FUENTES:
                    puntos = vehiculo.get(campo) or []
                    if isinstance(puntos, dict):
                        puntos = [puntos]
                    for punto in puntos:
                        fecha_utc = parsear_datetime(punto.get("time"))
                        valor = punto.get("value")
                        if fecha_utc is None or valor is None:
                            continue
                        fecha_dia = fecha_utc.astimezone(zona).date()
                        if fecha_dia < desde or fecha_dia > hasta_inclusivo:
                            continue
                        clave = (vehicle_id, fecha_dia)
                        registro = dias.setdefault(
                            clave,
                            {
                                "SamsaraVehicleId": vehicle_id,
                                "UnidadAPI": nombre,
                                "Fecha": fecha_dia,
                                "OBD": LecturasDia(),
                                "GPS": LecturasDia(),
                            },
                        )
                        if nombre:
                            registro["UnidadAPI"] = nombre
                        try:
                            registro[fuente].agregar(fecha_utc, float(valor))
                        except (TypeError, ValueError):
                            continue
        cursor = fin
    return dias


def tolerancia_fuentes(
    a: float,
    b: float,
    tolerancia_pct: float,
    tolerancia_km: float,
) -> bool:
    limite = max(tolerancia_km, max(abs(a), abs(b)) * tolerancia_pct / 100.0)
    return abs(a - b) <= limite


def validar_delta(
    delta: float | None,
    horas: float | None,
    respaldo: float | None,
    velocidad_maxima: float,
    tolerancia_pct: float,
    tolerancia_km: float,
) -> tuple[bool, str, float | None]:
    if delta is None:
        return False, "LECTURAS_INSUFICIENTES", None
    if delta < 0:
        return False, "REINICIO_O_RETROCESO_CONTADOR", None
    if horas is None or horas <= 0:
        return False, "TIEMPO_INVALIDO", None
    velocidad = delta / horas
    if velocidad > velocidad_maxima:
        return False, "VELOCIDAD_PROMEDIO_IMPOSIBLE", velocidad
    if respaldo is not None and respaldo >= 0:
        if tolerancia_fuentes(
            delta, respaldo, tolerancia_pct, tolerancia_km
        ):
            return True, "VALIDADO_CON_OTRA_FUENTE", velocidad
        return False, "DIVERGENCIA_ENTRE_OBD_Y_GPS", velocidad
    return True, "VALIDADO_POR_CONTINUIDAD_Y_TIEMPO", velocidad


def seleccionar_fuente_dia(
    obd: LecturasDia,
    gps: LecturasDia,
    args: argparse.Namespace,
) -> tuple[str | None, dict[str, tuple[bool, str, float | None]]]:
    obd_delta = obd.delta_km()
    gps_delta = gps.delta_km()
    # Primero se valida cada contador por si mismo. La comparacion entre fuentes
    # se usa despues para confirmar OBD o para activar el respaldo GPS.
    validaciones = {
        "OBD": validar_delta(
            obd_delta,
            obd.horas(),
            None,
            args.velocidad_maxima_promedio,
            args.tolerancia_fuentes_pct,
            args.tolerancia_fuentes_km,
        ),
        "GPS": validar_delta(
            gps_delta,
            gps.horas(),
            None,
            args.velocidad_maxima_promedio,
            args.tolerancia_fuentes_pct,
            args.tolerancia_fuentes_km,
        ),
    }
    obd_valido = validaciones["OBD"][0]
    gps_valido = validaciones["GPS"][0]
    if obd_valido and gps_valido and obd_delta is not None and gps_delta is not None:
        if tolerancia_fuentes(
            obd_delta,
            gps_delta,
            args.tolerancia_fuentes_pct,
            args.tolerancia_fuentes_km,
        ):
            validaciones["OBD"] = (
                True,
                "VALIDADO_CON_GPS",
                validaciones["OBD"][2],
            )
            return "OBD", validaciones
        validaciones["GPS"] = (
            True,
            "GPS_USADO_POR_DIVERGENCIA_OBD",
            validaciones["GPS"][2],
        )
        return "GPS", validaciones
    if obd_valido:
        return "OBD", validaciones
    if gps_valido:
        return "GPS", validaciones
    return None, validaciones


def construir_historico_diario(
    dias: dict[tuple[str, date], dict[str, Any]],
    catalogo: dict[str, dict[str, Any]],
    args: argparse.Namespace,
) -> pd.DataFrame:
    filas: list[dict[str, Any]] = []
    anterior_elegido: dict[str, dict[str, Any]] = {}

    for (vehicle_id, fecha_dia), registro in sorted(
        dias.items(), key=lambda item: (item[0][0], item[0][1])
    ):
        obd: LecturasDia = registro["OBD"]
        gps: LecturasDia = registro["GPS"]
        fuente, validaciones = seleccionar_fuente_dia(obd, gps, args)
        elegido = registro[fuente] if fuente else None
        delta_dia = elegido.delta_km() if elegido else 0.0
        motivo_dia = validaciones[fuente][1] if fuente else (
            f"OBD:{validaciones['OBD'][1]} | GPS:{validaciones['GPS'][1]}"
        )

        gap_candidato = 0.0
        gap_aceptado = 0.0
        velocidad_gap: float | None = None
        motivo_gap = "SIN_LECTURA_ANTERIOR"
        anterior = anterior_elegido.get(vehicle_id)

        if fuente and elegido and anterior:
            if anterior["Fuente"] != fuente:
                motivo_gap = "CAMBIO_DE_FUENTE_NO_COMPARABLE"
            elif elegido.primera_fecha and anterior["FechaFinal"]:
                gap_candidato = (
                    (elegido.primera_valor or 0.0) - anterior["ValorFinal"]
                ) / 1000.0
                horas_gap = (
                    elegido.primera_fecha - anterior["FechaFinal"]
                ).total_seconds() / 3600.0

                otra = "GPS" if fuente == "OBD" else "OBD"
                actual_otra: LecturasDia = registro[otra]
                anterior_otra: LecturasDia | None = anterior.get(otra)
                respaldo_gap: float | None = None
                if (
                    anterior_otra
                    and anterior_otra.ultima_valor is not None
                    and actual_otra.primera_valor is not None
                ):
                    respaldo_gap = (
                        actual_otra.primera_valor - anterior_otra.ultima_valor
                    ) / 1000.0

                aceptado, motivo_gap, velocidad_gap = validar_delta(
                    gap_candidato,
                    horas_gap,
                    respaldo_gap,
                    args.velocidad_maxima_promedio,
                    args.tolerancia_fuentes_pct,
                    args.tolerancia_fuentes_km,
                )
                if aceptado:
                    gap_aceptado = gap_candidato

        if fuente and elegido:
            anterior_elegido[vehicle_id] = {
                "Fuente": fuente,
                "FechaFinal": elegido.ultima_fecha,
                "ValorFinal": elegido.ultima_valor,
                "OBD": obd,
                "GPS": gps,
            }

        vehiculo = catalogo.get(vehicle_id) or {}
        unidad = str(vehiculo.get("name") or registro["UnidadAPI"] or "").strip()
        incidencias = []
        if fuente is None:
            incidencias.append("SIN_FUENTE_VALIDA")
        if motivo_gap not in {
            "SIN_LECTURA_ANTERIOR",
            "VALIDADO_CON_OTRA_FUENTE",
            "VALIDADO_POR_CONTINUIDAD_Y_TIEMPO",
        }:
            incidencias.append(motivo_gap)

        filas.append(
            {
                "SamsaraVehicleId": vehicle_id,
                "Unidad": unidad,
                "VIN": str(vehiculo.get("vin") or ""),
                "Placa": str(vehiculo.get("licensePlate") or ""),
                "Fecha": fecha_dia,
                "FuenteElegida": fuente or "SIN DATOS VALIDOS",
                "CantidadLecturasOBD": obd.cantidad,
                "FechaInicialOBD": obd.primera_fecha,
                "LecturaInicialOBDKm": (
                    obd.primera_valor / 1000.0 if obd.primera_valor is not None else None
                ),
                "FechaFinalOBD": obd.ultima_fecha,
                "LecturaFinalOBDKm": (
                    obd.ultima_valor / 1000.0 if obd.ultima_valor is not None else None
                ),
                "KmDentroDiaOBD": obd.delta_km(),
                "CantidadLecturasGPS": gps.cantidad,
                "FechaInicialGPS": gps.primera_fecha,
                "LecturaInicialGPSKm": (
                    gps.primera_valor / 1000.0 if gps.primera_valor is not None else None
                ),
                "FechaFinalGPS": gps.ultima_fecha,
                "LecturaFinalGPSKm": (
                    gps.ultima_valor / 1000.0 if gps.ultima_valor is not None else None
                ),
                "KmDentroDiaGPS": gps.delta_km(),
                "KmDentroDiaAceptados": max(0.0, delta_dia or 0.0),
                "MotivoKmDentroDia": motivo_dia,
                "KmNoRegistradosCandidatos": gap_candidato,
                "KmNoRegistradosAceptados": gap_aceptado,
                "VelocidadPromedioGapKmh": velocidad_gap,
                "MotivoKmNoRegistrados": motivo_gap,
                "KmTotalesAceptados": max(0.0, delta_dia or 0.0) + gap_aceptado,
                "Estatus": "REVISAR" if incidencias else "OK",
                "Incidencia": " | ".join(incidencias),
            }
        )
    return pd.DataFrame(filas)


def construir_resumen_mensual(diario: pd.DataFrame) -> pd.DataFrame:
    if diario.empty:
        return pd.DataFrame(
            columns=[
                "SamsaraVehicleId",
                "Unidad",
                "Mes",
                "KmDentroDiaAceptados",
                "KmNoRegistradosAceptados",
                "KilometrosTotalesMes",
                "DiasConDatos",
                "RegistrosRevisar",
                "FuentePrincipal",
            ]
        )
    base = diario.copy()
    base["Mes"] = pd.to_datetime(base["Fecha"]).dt.strftime("%Y-%m")
    agrupado = base.groupby(
        ["SamsaraVehicleId", "Unidad", "Mes"], dropna=False, as_index=False
    ).agg(
        KmDentroDiaAceptados=("KmDentroDiaAceptados", "sum"),
        KmNoRegistradosAceptados=("KmNoRegistradosAceptados", "sum"),
        KilometrosTotalesMes=("KmTotalesAceptados", "sum"),
        DiasConDatos=("Fecha", "nunique"),
        RegistrosRevisar=("Estatus", lambda serie: int((serie == "REVISAR").sum())),
        FuentePrincipal=(
            "FuenteElegida",
            lambda serie: Counter(x for x in serie if x != "SIN DATOS VALIDOS").most_common(1)[0][0]
            if any(x != "SIN DATOS VALIDOS" for x in serie)
            else "SIN DATOS",
        ),
    )
    return agrupado.sort_values(["Unidad", "Mes"], kind="stable")


def catalogo_dataframe(vehiculos: list[dict[str, Any]]) -> pd.DataFrame:
    filas = []
    for vehiculo in vehiculos:
        filas.append(
            {
                "SamsaraVehicleId": str(vehiculo.get("id") or ""),
                "Unidad": str(vehiculo.get("name") or ""),
                "VIN": str(vehiculo.get("vin") or ""),
                "Placa": str(vehiculo.get("licensePlate") or ""),
                "Marca": str(vehiculo.get("make") or ""),
                "Modelo": str(vehiculo.get("model") or ""),
                "Ano": vehiculo.get("year"),
                "CreadaEnSamsara": vehiculo.get("createdAtTime"),
            }
        )
    return pd.DataFrame(filas).sort_values("Unidad", kind="stable")


def escribir_tabla(
    writer: pd.ExcelWriter,
    df: pd.DataFrame,
    nombre: str,
    formato_encabezado: Any,
    formato_fecha: Any,
    formato_datetime: Any,
    formato_numero: Any,
) -> None:
    df_excel = df.copy()
    for columna in df_excel.columns:
        if pd.api.types.is_datetime64_any_dtype(df_excel[columna]):
            # Excel no acepta datetimes con zona horaria.
            df_excel[columna] = pd.to_datetime(df_excel[columna], utc=True).dt.tz_localize(None)
    df_excel.to_excel(writer, sheet_name=nombre[:31], index=False)
    hoja = writer.sheets[nombre[:31]]
    hoja.freeze_panes(1, 2 if len(df_excel.columns) > 2 else 1)
    hoja.autofilter(0, 0, max(0, len(df_excel)), max(0, len(df_excel.columns) - 1))
    hoja.set_row(0, 32, formato_encabezado)
    for indice, columna in enumerate(df_excel.columns):
        muestra = df_excel[columna].dropna().astype(str).head(2000)
        ancho = max([len(str(columna))] + [len(x) for x in muestra]) + 2
        ancho = min(max(ancho, 11), 42)
        formato = None
        if columna == "Fecha":
            formato = formato_fecha
        elif "Fecha" in columna:
            formato = formato_datetime
        elif columna.startswith("Km") or columna.endswith("Kmh"):
            formato = formato_numero
        hoja.set_column(indice, indice, ancho, formato)


def escribir_excel(
    salida: Path,
    resumen: pd.DataFrame,
    diario: pd.DataFrame,
    vehiculos: pd.DataFrame,
    args: argparse.Namespace,
    desde: date,
    hasta: date,
) -> None:
    salida.parent.mkdir(parents=True, exist_ok=True)
    incidencias = diario[diario["Estatus"] == "REVISAR"].copy()
    with pd.ExcelWriter(
        salida,
        engine="xlsxwriter",
        datetime_format="yyyy-mm-dd hh:mm:ss",
        date_format="yyyy-mm-dd",
    ) as writer:
        libro = writer.book
        encabezado = libro.add_format(
            {
                "bold": True,
                "font_color": "#FFFFFF",
                "bg_color": "#1F4E78",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
            }
        )
        titulo = libro.add_format(
            {"bold": True, "font_size": 16, "font_color": "#1F4E78"}
        )
        etiqueta = libro.add_format({"bold": True, "bg_color": "#D9EAF7"})
        fecha_fmt = libro.add_format({"num_format": "yyyy-mm-dd"})
        datetime_fmt = libro.add_format({"num_format": "yyyy-mm-dd hh:mm:ss"})
        numero_fmt = libro.add_format({"num_format": "#,##0.00"})

        hoja_info = libro.add_worksheet("Informacion")
        writer.sheets["Informacion"] = hoja_info
        hoja_info.write("A1", "Histórico de kilómetros Samsara", titulo)
        informacion = [
            ("Fecha inicial", desde.isoformat()),
            ("Fecha final", hasta.isoformat()),
            ("Zona horaria", args.zona_horaria),
            ("Fuente primaria", "obdOdometerMeters (OBD/ECU)"),
            ("Fuente de respaldo", "gpsDistanceMeters"),
            ("Velocidad máxima promedio", args.velocidad_maxima_promedio),
            ("Tolerancia entre fuentes %", args.tolerancia_fuentes_pct),
            ("Tolerancia mínima entre fuentes km", args.tolerancia_fuentes_km),
            (
                "Definición Km no registrados",
                "Diferencia validada entre la última lectura elegida y la primera lectura siguiente de la misma fuente.",
            ),
            (
                "Fuente oficial",
                "https://developers.samsara.com/docs/mileage-and-distance",
            ),
        ]
        for fila, (nombre, valor) in enumerate(informacion, start=2):
            hoja_info.write(fila, 0, nombre, etiqueta)
            hoja_info.write(fila, 1, valor)
        hoja_info.set_column("A:A", 38)
        hoja_info.set_column("B:B", 105)

        escribir_tabla(
            writer, resumen, "Resumen mensual", encabezado, fecha_fmt,
            datetime_fmt, numero_fmt
        )
        escribir_tabla(
            writer, vehiculos, "Unidades", encabezado, fecha_fmt,
            datetime_fmt, numero_fmt
        )

        for numero, inicio in enumerate(range(0, len(diario), MAX_FILAS_HOJA), start=1):
            nombre = "Historico diario" if numero == 1 else f"Historico diario {numero}"
            escribir_tabla(
                writer,
                diario.iloc[inicio : inicio + MAX_FILAS_HOJA],
                nombre,
                encabezado,
                fecha_fmt,
                datetime_fmt,
                numero_fmt,
            )

        for numero, inicio in enumerate(
            range(0, len(incidencias), MAX_FILAS_HOJA), start=1
        ):
            nombre = "Incidencias" if numero == 1 else f"Incidencias {numero}"
            escribir_tabla(
                writer,
                incidencias.iloc[inicio : inicio + MAX_FILAS_HOJA],
                nombre,
                encabezado,
                fecha_fmt,
                datetime_fmt,
                numero_fmt,
            )

        if diario.empty:
            hoja = libro.add_worksheet("Historico diario")
            hoja.write("A1", "No se encontraron lecturas en el periodo.", titulo)
        if incidencias.empty:
            hoja = libro.add_worksheet("Incidencias")
            hoja.write("A1", "No se detectaron incidencias.", titulo)

        for nombre_hoja in writer.sheets:
            if nombre_hoja.startswith("Resumen mensual"):
                hoja = writer.sheets[nombre_hoja]
                columna = resumen.columns.get_loc("RegistrosRevisar")
                hoja.conditional_format(
                    1,
                    columna,
                    max(1, len(resumen)),
                    columna,
                    {
                        "type": "cell",
                        "criteria": ">",
                        "value": 0,
                        "format": libro.add_format(
                            {"bg_color": "#FFF2CC", "font_color": "#9C6500"}
                        ),
                    },
                )


def determinar_periodo(
    args: argparse.Namespace,
    vehiculos: list[dict[str, Any]],
    zona: ZoneInfo,
) -> tuple[date, date]:
    hoy = datetime.now(zona).date()
    hasta = fecha_iso(args.hasta) or (hoy if args.incluir_hoy else hoy - timedelta(days=1))
    desde = fecha_iso(args.desde)
    if desde is None:
        fechas_creacion = [
            parsear_datetime(v.get("createdAtTime"))
            for v in vehiculos
            if v.get("createdAtTime")
        ]
        fechas_creacion = [x for x in fechas_creacion if x is not None]
        if not fechas_creacion:
            raise RuntimeError(
                "Samsara no devolvio createdAtTime. Ejecuta nuevamente indicando "
                "--desde YYYY-MM-DD."
            )
        desde = min(fechas_creacion).astimezone(zona).date()
    if desde > hasta:
        raise ValueError(f"La fecha inicial {desde} es posterior a la final {hasta}")
    return desde, hasta


def main() -> int:
    args = parsear_argumentos()
    if importlib.util.find_spec("xlsxwriter") is None:
        print(
            "Falta XlsxWriter. Instalalo antes de iniciar la descarga con:\n"
            "  .\\.venv\\Scripts\\python.exe -m pip install XlsxWriter",
            file=sys.stderr,
        )
        return 2
    load_dotenv(REPO_ROOT / ".env")
    token = os.getenv("SAMSARA_API_TOKEN") or os.getenv("SAMSARA_TOKEN")
    if not token:
        print(
            "Falta SAMSARA_API_TOKEN o SAMSARA_TOKEN en el archivo .env.",
            file=sys.stderr,
        )
        return 2

    zona = ZoneInfo(args.zona_horaria)
    sesion = requests.Session()
    sesion.headers.update(
        {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    )

    print("Consultando catalogo de unidades...")
    vehiculos = catalogo_vehiculos(sesion)
    if not vehiculos:
        raise RuntimeError("Samsara no devolvio unidades")
    desde, hasta = determinar_periodo(args, vehiculos, zona)
    print(f"Unidades encontradas: {len(vehiculos):,}")
    print(f"Periodo: {desde} a {hasta}")
    print("La descarga puede tardar dependiendo de la antiguedad y el volumen.")

    dias = extraer_historico(
        sesion, desde, hasta, zona, args.ventana_dias
    )
    catalogo = {str(v.get("id") or ""): v for v in vehiculos}
    diario = construir_historico_diario(dias, catalogo, args)
    resumen = construir_resumen_mensual(diario)
    unidades = catalogo_dataframe(vehiculos)
    escribir_excel(args.salida.resolve(), resumen, diario, unidades, args, desde, hasta)

    print()
    print(f"Excel generado: {args.salida.resolve()}")
    print(f"Registros diarios: {len(diario):,}")
    print(f"Registros mensuales: {len(resumen):,}")
    if not diario.empty:
        print(f"Kilometros aceptados: {diario['KmTotalesAceptados'].sum():,.2f}")
        print(f"Registros para revisar: {(diario['Estatus'] == 'REVISAR').sum():,}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
