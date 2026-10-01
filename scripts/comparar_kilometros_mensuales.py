"""Compara kilometros mensuales de Samsara API contra HistoricoKmSamsara.

Genera un Excel y un CSV con una fila por unidad y mes. La API usa
``obdOdometerMeters`` como fuente primaria y ``gpsDistanceMeters`` como
respaldo. La columna SQL suma ``DiferenciaKm + KmNoRegistrados``.
"""

from __future__ import annotations

import argparse
import os
from datetime import date, datetime, time, timezone
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import pandas as pd
import pyodbc
import requests
from dotenv import load_dotenv
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from extraer_kilometros_samsara import (
    API_BASE,
    calcular_fila,
    catalogo_vehiculos,
    iso_utc,
    obtener_paginas,
)


REPO_ROOT = Path(__file__).resolve().parents[1]
MX_TZ = ZoneInfo("America/Mexico_City")
TIPOS = "obdOdometerMeters,gpsDistanceMeters"


def argumentos() -> argparse.Namespace:
    hoy = datetime.now(MX_TZ)
    parser = argparse.ArgumentParser(
        description="Compara kilometros mensuales de Samsara API contra SQL."
    )
    parser.add_argument(
        "--inicio",
        default=f"{hoy.year}-01-01",
        help="Fecha inicial YYYY-MM-DD; por defecto, inicio del ano actual.",
    )
    parser.add_argument(
        "--fin",
        default=hoy.date().isoformat(),
        help="Fecha final inclusiva YYYY-MM-DD; por defecto, hoy.",
    )
    parser.add_argument(
        "--salida",
        default=str(REPO_ROOT / "outputs" / "comparativa_kilometros_mensuales.xlsx"),
        help="Ruta del Excel de salida.",
    )
    parser.add_argument(
        "--sin-sql",
        action="store_true",
        help="Genera solo los datos de la API, sin intentar conectarse a SQL.",
    )
    parser.add_argument(
        "--api-csv",
        help=(
            "CSV previamente generado con datos Samsara. Si se indica, evita "
            "repetir la extraccion de la API y consulta solamente SQL."
        ),
    )
    return parser.parse_args()


def fecha_local(valor: str, fin: bool = False) -> datetime:
    dia = date.fromisoformat(valor)
    if fin:
        return datetime.combine(dia, time.max, tzinfo=MX_TZ)
    return datetime.combine(dia, time.min, tzinfo=MX_TZ)


def inicio_mes(fecha: datetime) -> datetime:
    return fecha.replace(day=1, hour=0, minute=0, second=0, microsecond=0)


def siguiente_mes(fecha: datetime) -> datetime:
    if fecha.month == 12:
        return fecha.replace(year=fecha.year + 1, month=1, day=1)
    return fecha.replace(month=fecha.month + 1, day=1)


def periodos_mensuales(inicio: datetime, fin: datetime) -> list[tuple[datetime, datetime]]:
    periodos: list[tuple[datetime, datetime]] = []
    cursor = inicio
    while cursor < fin:
        cierre = min(siguiente_mes(inicio_mes(cursor)), fin)
        periodos.append((cursor, cierre))
        cursor = cierre
    return periodos


def snapshot(
    sesion: requests.Session, fecha: datetime
) -> dict[str, dict[str, Any]]:
    datos = obtener_paginas(
        sesion,
        "/fleet/vehicles/stats",
        {"types": TIPOS, "time": iso_utc(fecha)},
    )
    return {str(fila["id"]): fila for fila in datos}


def extraer_api(
    sesion: requests.Session,
    vehiculos: list[dict[str, Any]],
    inicio: datetime,
    fin: datetime,
) -> pd.DataFrame:
    filas: list[dict[str, Any]] = []
    cache: dict[str, dict[str, dict[str, Any]]] = {}

    def leer_corte(fecha: datetime) -> dict[str, dict[str, Any]]:
        clave = iso_utc(fecha)
        if clave not in cache:
            print(f"Consultando Samsara al corte {clave}...")
            cache[clave] = snapshot(sesion, fecha)
        return cache[clave]

    for desde, hasta in periodos_mensuales(inicio, fin):
        inicial = leer_corte(desde)
        final = leer_corte(hasta)
        mes = desde.strftime("%Y-%m")
        for vehiculo in vehiculos:
            vehicle_id = str(vehiculo.get("id") or "")
            calculo = calcular_fila(
                vehiculo,
                inicial.get(vehicle_id),
                final.get(vehicle_id),
                desde,
                hasta,
                sesion,
            )
            filas.append(
                {
                    "SamsaraVehicleId": vehicle_id,
                    "Unidad": calculo["Unidad"],
                    "Mes": mes,
                    "InicioPeriodo": desde,
                    "FinPeriodo": hasta,
                    "KilometrosAPI": calculo["Kilometros"],
                    "MetodoAPI": calculo["Metodo"],
                    "CoberturaAPI": calculo["Cobertura"],
                    "LecturaInicialAPI": calculo["FechaLecturaInicial"],
                    "LecturaFinalAPI": calculo["FechaLecturaFinal"],
                    "ObservacionAPI": calculo["Observacion"],
                }
            )
    return pd.DataFrame(filas)


def conectar_sql() -> pyodbc.Connection:
    driver = os.getenv("SQL_DRIVER", "ODBC Driver 18 for SQL Server")
    cifrado = "Optional" if "Driver 18" in driver else "no"
    cadena = (
        f"DRIVER={{{driver}}};"
        f"SERVER={os.environ['SQL_SERVER']};"
        f"DATABASE={os.environ['SQL_DATABASE']};"
        f"UID={os.environ['SQL_USER']};"
        f"PWD={os.environ['SQL_PASSWORD']};"
        f"Encrypt={cifrado};TrustServerCertificate=yes;Connection Timeout=15;"
    )
    return pyodbc.connect(cadena)


def localizar_tabla_historica(conexion: pyodbc.Connection) -> str:
    """Localiza HistoricoKmSamsara aunque SQL_DATABASE apunte a otra base."""
    cursor = conexion.cursor()
    bases = [
        str(fila[0])
        for fila in cursor.execute(
            "SELECT name FROM sys.databases "
            "WHERE state = 0 AND HAS_DBACCESS(name) = 1"
        ).fetchall()
    ]
    coincidencias: list[str] = []
    for base in bases:
        base_segura = base.replace("]", "]]" )
        consulta = (
            f"SELECT s.name, t.name FROM [{base_segura}].sys.tables t "
            f"JOIN [{base_segura}].sys.schemas s ON s.schema_id = t.schema_id "
            "WHERE t.name = ?"
        )
        try:
            filas = cursor.execute(consulta, "HistoricoKmSamsara").fetchall()
        except pyodbc.Error:
            continue
        for esquema, tabla in filas:
            esquema_seguro = str(esquema).replace("]", "]]" )
            tabla_segura = str(tabla).replace("]", "]]" )
            coincidencias.append(
                f"[{base_segura}].[{esquema_seguro}].[{tabla_segura}]"
            )
    if not coincidencias:
        raise RuntimeError(
            "No se encontro HistoricoKmSamsara en ninguna base SQL accesible."
        )
    if len(coincidencias) > 1:
        raise RuntimeError(
            "Se encontro HistoricoKmSamsara en mas de una base: "
            + ", ".join(coincidencias)
        )
    return coincidencias[0]


def extraer_sql(inicio: datetime, fin: datetime) -> pd.DataFrame:
    with conectar_sql() as conexion:
        tabla = localizar_tabla_historica(conexion)
        print(f"Consultando historico SQL en {tabla}...")
        consulta = f"""
        SELECT
            CAST(IdTracto AS varchar(100)) AS SamsaraVehicleId,
            MAX(Unidad) AS UnidadSQL,
            CONVERT(char(7), TRY_CONVERT(datetime2, FechaInicio), 120) AS Mes,
            CAST(SUM(COALESCE(DiferenciaKm, 0) + COALESCE(KmNoRegistrados, 0))
                AS decimal(18, 3)) AS KilometrosSQL,
            COUNT(*) AS RegistrosSQL
        FROM {tabla}
        WHERE TRY_CONVERT(datetime2, FechaInicio) >= ?
          AND TRY_CONVERT(datetime2, FechaInicio) < ?
        GROUP BY
            CAST(IdTracto AS varchar(100)),
            CONVERT(char(7), TRY_CONVERT(datetime2, FechaInicio), 120)
        """
        cursor = conexion.cursor()
        cursor.execute(
            consulta,
            inicio.astimezone(timezone.utc).replace(tzinfo=None),
            fin.astimezone(timezone.utc).replace(tzinfo=None),
        )
        columnas = [descripcion[0] for descripcion in cursor.description]
        return pd.DataFrame.from_records(cursor.fetchall(), columns=columnas)


def comparar(api: pd.DataFrame, sql: pd.DataFrame | None) -> pd.DataFrame:
    # Permite reutilizar un CSV generado anteriormente sin conservar columnas
    # SQL vacias o antiguas que colisionen con la nueva consulta.
    columnas_sql_previas = {
        "UnidadSQL",
        "KilometrosSQL",
        "RegistrosSQL",
        "DiferenciaKm",
        "DiferenciaPct",
        "EstadoComparacion",
    }
    api = api.drop(
        columns=[c for c in columnas_sql_previas if c in api.columns],
        errors="ignore",
    ).copy()
    if sql is None:
        resultado = api.copy()
        resultado["KilometrosSQL"] = pd.NA
        resultado["DiferenciaKm"] = pd.NA
        resultado["DiferenciaPct"] = pd.NA
        resultado["EstadoComparacion"] = "SQL no disponible"
        return resultado

    resultado = api.merge(
        sql,
        how="outer",
        on=["SamsaraVehicleId", "Mes"],
        suffixes=("", "_SQL"),
    )
    resultado["Unidad"] = resultado["Unidad"].fillna(resultado.get("UnidadSQL"))
    resultado["KilometrosAPI"] = pd.to_numeric(
        resultado["KilometrosAPI"], errors="coerce"
    )
    resultado["KilometrosSQL"] = pd.to_numeric(
        resultado["KilometrosSQL"], errors="coerce"
    )
    if "RegistrosSQL" in resultado:
        resultado["RegistrosSQL"] = pd.to_numeric(
            resultado["RegistrosSQL"], errors="coerce"
        )
    resultado["DiferenciaKm"] = resultado["KilometrosAPI"] - resultado["KilometrosSQL"]
    resultado["DiferenciaPct"] = (
        resultado["DiferenciaKm"] / resultado["KilometrosSQL"].replace(0, pd.NA)
    ) * 100

    def estado(fila: pd.Series) -> str:
        if pd.isna(fila.get("KilometrosAPI")) and pd.isna(fila.get("KilometrosSQL")):
            return "Sin datos en ambas fuentes"
        if pd.isna(fila.get("KilometrosAPI")):
            return "Solo SQL"
        if pd.isna(fila.get("KilometrosSQL")):
            return "Solo API"
        if float(fila["KilometrosSQL"]) == 0:
            return (
                "Coincide (ambos 0)"
                if float(fila["KilometrosAPI"]) == 0
                else "Revisar (SQL=0)"
            )
        if pd.isna(fila.get("DiferenciaPct")):
            return "Revisar (sin porcentaje)"
        if abs(float(fila["DiferenciaPct"])) <= 5:
            return "Coincide (<=5%)"
        return "Revisar (>5%)"

    resultado["EstadoComparacion"] = resultado.apply(estado, axis=1)
    return resultado


def guardar(resultado: pd.DataFrame, salida: Path, error_sql: str | None) -> None:
    salida.parent.mkdir(parents=True, exist_ok=True)
    resultado = resultado.copy()
    for columna_fecha in ("InicioPeriodo", "FinPeriodo"):
        if columna_fecha in resultado:
            resultado[columna_fecha] = pd.to_datetime(
                resultado[columna_fecha], utc=True, errors="coerce"
            ).dt.tz_localize(None)
    columnas_principales = [
        "Unidad",
        "Mes",
        "KilometrosAPI",
        "KilometrosSQL",
        "DiferenciaKm",
        "DiferenciaPct",
        "MetodoAPI",
        "CoberturaAPI",
        "EstadoComparacion",
        "SamsaraVehicleId",
        "RegistrosSQL",
        "InicioPeriodo",
        "FinPeriodo",
        "LecturaInicialAPI",
        "LecturaFinalAPI",
        "ObservacionAPI",
    ]
    for columna in columnas_principales:
        if columna not in resultado:
            resultado[columna] = pd.NA
    resultado = resultado[columnas_principales].sort_values(["Unidad", "Mes"])

    resumen = pd.DataFrame(
        [
            {"Indicador": "Filas unidad-mes", "Valor": len(resultado)},
            {
                "Indicador": "Kilometros totales API",
                "Valor": pd.to_numeric(resultado["KilometrosAPI"], errors="coerce").sum(),
            },
            {
                "Indicador": "Kilometros totales SQL",
                "Valor": pd.to_numeric(resultado["KilometrosSQL"], errors="coerce").sum(),
            },
            {
                "Indicador": "Estado SQL",
                "Valor": "Disponible" if not error_sql else f"No disponible: {error_sql}",
            },
        ]
    )
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        resultado.to_excel(writer, sheet_name="Comparativa", index=False)
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        hoja = writer.sheets["Comparativa"]
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions
        for columna in hoja.columns:
            ancho = min(max(len(str(celda.value or "")) for celda in columna) + 2, 55)
            hoja.column_dimensions[columna[0].column_letter].width = ancho

    resultado.to_csv(salida.with_suffix(".csv"), index=False, encoding="utf-8-sig")


def main() -> None:
    args = argumentos()
    inicio = fecha_local(args.inicio)
    fin = fecha_local(args.fin, fin=True)
    if inicio >= fin:
        raise SystemExit("El periodo solicitado no es valido.")

    load_dotenv(REPO_ROOT / ".env")
    vehiculos: list[dict[str, Any]] = []
    if args.api_csv:
        print(f"Reutilizando datos Samsara desde {args.api_csv}...")
        datos_api = pd.read_csv(
            args.api_csv,
            dtype={"SamsaraVehicleId": "string"},
        )
    else:
        token = os.getenv("SAMSARA_API_TOKEN") or os.getenv("SAMSARA_TOKEN")
        if not token:
            raise SystemExit("Falta SAMSARA_API_TOKEN o SAMSARA_TOKEN en .env")

        sesion = requests.Session()
        sesion.headers.update(
            {"Accept": "application/json", "Authorization": f"Bearer {token}"}
        )
        reintentos = Retry(
            total=6,
            connect=6,
            read=6,
            status=6,
            backoff_factor=1,
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=frozenset({"GET"}),
            respect_retry_after_header=True,
        )
        sesion.mount("https://", HTTPAdapter(max_retries=reintentos))
        print(f"Consultando catalogo en {API_BASE}...")
        vehiculos = catalogo_vehiculos(sesion)
        datos_api = extraer_api(sesion, vehiculos, inicio, fin)

    datos_sql: pd.DataFrame | None = None
    error_sql: str | None = None
    if not args.sin_sql:
        try:
            datos_sql = extraer_sql(inicio, fin)
        except Exception as exc:  # La API puede reportarse aunque SQL no sea accesible.
            error_sql = f"{type(exc).__name__}: {exc}"
            print("SQL no disponible; se generara el reporte con datos API.")

    resultado = comparar(datos_api, datos_sql)
    salida = Path(args.salida)
    guardar(resultado, salida, error_sql)
    print(f"Reporte Excel: {salida}")
    print(f"Reporte CSV: {salida.with_suffix('.csv')}")
    total_unidades = (
        len(vehiculos)
        if vehiculos
        else datos_api["SamsaraVehicleId"].nunique(dropna=True)
    )
    print(f"Unidades: {total_unidades}; filas unidad-mes: {len(resultado)}")


if __name__ == "__main__":
    main()
