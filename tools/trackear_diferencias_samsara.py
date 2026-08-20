import argparse
import math
import os
import re
import sys
from datetime import timedelta
from pathlib import Path

import pandas as pd
import pytz
from dotenv import load_dotenv

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from detenciones import consultar_historial_gps


MX_TZ = pytz.timezone("America/Mexico_City")

COLUMNAS_RESUMEN = [
    "Unidad",
    "SamsaraVehicleId",
    "EstatusExterno",
    "EstatusPython",
    "UltimoReporteExterno",
    "GpsTimeSamsara",
    "DiferenciaMinutos",
    "DistanciaMetros",
    "VelocidadSamsaraMph",
    "IsEcuSpeed",
    "LatitudExterna",
    "LongitudExterna",
    "LatitudSamsara",
    "LongitudSamsara",
    "DireccionExterna",
    "DireccionSamsara",
]

COLUMNAS_LINEA_TIEMPO = [
    "Unidad",
    "SamsaraVehicleId",
    "Fuente",
    "FechaHoraMexico",
    "Estatus",
    "VelocidadMph",
    "IsEcuSpeed",
    "Latitud",
    "Longitud",
    "Direccion",
    "Detalle",
]

COLUMNAS_TRANSICIONES = (
    COLUMNAS_LINEA_TIEMPO
    + ["DuracionHastaSiguienteMinutos"]
)


def argumentos():
    parser = argparse.ArgumentParser(
        description=(
            "Detecta diferencias de estatus entre un reporte externo "
            "y la auditoria de EnvioMain1, y genera una linea de tiempo."
        )
    )
    parser.add_argument(
        "--samsara",
        default=str(REPO_ROOT / "comparacion_telemetria_samsara.csv"),
        help="CSV generado por EnvioMain1.",
    )
    parser.add_argument(
        "--externo",
        required=True,
        help="Reporte externo .xls, .xlsx o .csv.",
    )
    parser.add_argument(
        "--salida",
        default=str(REPO_ROOT / "outputs" / "seguimiento_estatus"),
        help="Directorio para los CSV resultantes.",
    )
    parser.add_argument(
        "--margen-minutos",
        type=int,
        default=15,
        help="Minutos adicionales antes y despues para consultar historial.",
    )
    parser.add_argument(
        "--sin-historial",
        action="store_true",
        help="No consulta la API; genera la linea de tiempo con ambos reportes.",
    )
    return parser.parse_args()


def normalizar_unidad(valor):
    texto = "" if pd.isna(valor) else str(valor).strip()
    texto = re.sub(r"\.0$", "", texto)
    coincidencia = re.match(r"^(\d+)", texto)
    return coincidencia.group(1) if coincidencia else texto.upper()


def normalizar_estatus_externo(valor):
    texto = "" if pd.isna(valor) else str(valor).strip().upper()

    if "MOVIMIENTO" in texto or texto == "RUTA":
        return "RUTA"

    if "DETENIDO" in texto:
        return "DETENIDO"

    return texto or "SIN ESTATUS"


def estatus_punto_samsara(velocidad, is_ecu_speed):
    try:
        velocidad_num = float(velocidad or 0)
    except (TypeError, ValueError):
        velocidad_num = 0.0

    if velocidad_num > 0:
        return "RUTA"

    if is_ecu_speed is True:
        return "DETENIDO"

    return "SIN ECU"


def buscar_columna(df, opciones):
    mapa = {
        str(columna).strip().casefold(): columna
        for columna in df.columns
    }

    for opcion in opciones:
        encontrada = mapa.get(opcion.casefold())
        if encontrada is not None:
            return encontrada

    raise ValueError(
        f"No se encontro ninguna columna de {opciones}. "
        f"Columnas disponibles: {list(df.columns)}"
    )


def leer_reporte_externo(ruta):
    extension = ruta.suffix.lower()

    if extension == ".csv":
        df = pd.read_csv(ruta)
    elif extension == ".xls":
        df = pd.read_excel(
            ruta,
            engine="xlrd",
            engine_kwargs={
                "ignore_workbook_corruption": True,
            },
        )
    elif extension == ".xlsx":
        df = pd.read_excel(
            ruta,
            engine="openpyxl",
        )
    else:
        raise ValueError(
            f"Formato externo no soportado: {extension}"
        )

    columnas = {
        "UnidadExterna": buscar_columna(
            df,
            ["Unidad"],
        ),
        "EstatusExternoOriginal": buscar_columna(
            df,
            ["Estatus_actual", "Estatus", "Estado"],
        ),
        "LatitudExterna": buscar_columna(
            df,
            ["Latitud", "Latitude"],
        ),
        "LongitudExterna": buscar_columna(
            df,
            ["Longitud", "Longitude"],
        ),
        "UltimoReporteExterno": buscar_columna(
            df,
            ["Ultimo reporte", "Último reporte", "Fecha"],
        ),
        "DireccionExterna": buscar_columna(
            df,
            ["Direccion", "Dirección", "Ubicacion", "Ubicación"],
        ),
    }

    externo = df[
        list(columnas.values())
    ].rename(
        columns={
            original: nuevo
            for nuevo, original in columnas.items()
        }
    )

    externo["UnidadClave"] = externo[
        "UnidadExterna"
    ].map(normalizar_unidad)
    externo["EstatusExterno"] = externo[
        "EstatusExternoOriginal"
    ].map(normalizar_estatus_externo)
    externo["FechaExterna"] = pd.to_datetime(
        externo["UltimoReporteExterno"],
        dayfirst=True,
        errors="coerce",
    )

    return externo


def leer_reporte_samsara(ruta):
    samsara = pd.read_csv(
        ruta,
        dtype={
            "Unidad": str,
            "SamsaraVehicleId": str,
        },
    )

    requeridas = {
        "Unidad",
        "SamsaraVehicleId",
        "Decision",
        "EstatusInicial",
        "GpsTimeUTC",
        "Latitude",
        "Longitude",
        "ReverseGeo",
        "SpeedMilesPerHour",
        "IsEcuSpeed",
    }
    faltantes = requeridas - set(samsara.columns)

    if faltantes:
        raise ValueError(
            "Faltan columnas en el CSV de Samsara: "
            + ", ".join(sorted(faltantes))
        )

    samsara["UnidadClave"] = samsara[
        "Unidad"
    ].map(normalizar_unidad)
    samsara["FechaSamsara"] = pd.to_datetime(
        samsara["GpsTimeUTC"],
        utc=True,
        errors="coerce",
    ).dt.tz_convert(MX_TZ).dt.tz_localize(None)

    return samsara


def haversine_metros(latitud_1, longitud_1, latitud_2, longitud_2):
    try:
        lat_1 = math.radians(float(latitud_1))
        lon_1 = math.radians(float(longitud_1))
        lat_2 = math.radians(float(latitud_2))
        lon_2 = math.radians(float(longitud_2))
    except (TypeError, ValueError):
        return None

    delta_lat = lat_2 - lat_1
    delta_lon = lon_2 - lon_1
    a = (
        math.sin(delta_lat / 2) ** 2
        + math.cos(lat_1)
        * math.cos(lat_2)
        * math.sin(delta_lon / 2) ** 2
    )
    return 6371000 * 2 * math.asin(math.sqrt(a))


def detectar_diferencias(externo, samsara):
    columna_estatus = "EstatusInicial"

    if "EstatusFinal" in samsara.columns:
        samsara["EstatusComparado"] = samsara[
            "EstatusFinal"
        ].where(
            samsara["EstatusFinal"].notna()
            & samsara["EstatusFinal"].ne(""),
            samsara["EstatusInicial"],
        )
        columna_estatus = "EstatusComparado"

    incluidas = samsara[
        samsara["Decision"].eq("INCLUIDA")
        & samsara[columna_estatus].isin(
            ["RUTA", "DETENIDO"]
        )
    ].copy()

    cruce = externo.merge(
        incluidas,
        on="UnidadClave",
        how="inner",
    )

    diferencias = cruce[
        cruce["EstatusExterno"].ne(
            cruce[columna_estatus]
        )
    ].copy()

    diferencias["EstatusPythonComparado"] = diferencias[
        columna_estatus
    ]

    diferencias["DiferenciaMinutos"] = (
        diferencias["FechaSamsara"]
        - diferencias["FechaExterna"]
    ).dt.total_seconds() / 60

    diferencias["DistanciaMetros"] = diferencias.apply(
        lambda row: haversine_metros(
            row["LatitudExterna"],
            row["LongitudExterna"],
            row["Latitude"],
            row["Longitude"],
        ),
        axis=1,
    )

    return diferencias


def construir_resumen(diferencias):
    if diferencias.empty:
        return pd.DataFrame(
            columns=COLUMNAS_RESUMEN
        )

    resumen = pd.DataFrame({
        "Unidad": diferencias["UnidadClave"],
        "SamsaraVehicleId": diferencias[
            "SamsaraVehicleId"
        ],
        "EstatusExterno": diferencias[
            "EstatusExterno"
        ],
        "EstatusPython": diferencias[
            "EstatusPythonComparado"
        ],
        "UltimoReporteExterno": diferencias[
            "FechaExterna"
        ],
        "GpsTimeSamsara": diferencias[
            "FechaSamsara"
        ],
        "DiferenciaMinutos": diferencias[
            "DiferenciaMinutos"
        ].round(2),
        "DistanciaMetros": diferencias[
            "DistanciaMetros"
        ].round(1),
        "VelocidadSamsaraMph": diferencias[
            "SpeedMilesPerHour"
        ],
        "IsEcuSpeed": diferencias[
            "IsEcuSpeed"
        ],
        "LatitudExterna": diferencias[
            "LatitudExterna"
        ],
        "LongitudExterna": diferencias[
            "LongitudExterna"
        ],
        "LatitudSamsara": diferencias[
            "Latitude"
        ],
        "LongitudSamsara": diferencias[
            "Longitude"
        ],
        "DireccionExterna": diferencias[
            "DireccionExterna"
        ],
        "DireccionSamsara": diferencias[
            "ReverseGeo"
        ],
    })

    return resumen.sort_values(
        by="Unidad"
    )


def construir_linea_tiempo_base(diferencias):
    eventos = []

    for _, row in diferencias.iterrows():
        datos_comunes = {
            "Unidad": row["UnidadClave"],
            "SamsaraVehicleId": row[
                "SamsaraVehicleId"
            ],
        }

        eventos.append({
            **datos_comunes,
            "Fuente": "REPORTE EXTERNO",
            "FechaHoraMexico": row["FechaExterna"],
            "Estatus": row["EstatusExterno"],
            "VelocidadMph": None,
            "IsEcuSpeed": None,
            "Latitud": row["LatitudExterna"],
            "Longitud": row["LongitudExterna"],
            "Direccion": row["DireccionExterna"],
            "Detalle": (
                f"Estatus original: "
                f"{row['EstatusExternoOriginal']}"
            ),
        })

        eventos.append({
            **datos_comunes,
            "Fuente": "SAMSARA ACTUAL",
            "FechaHoraMexico": row["FechaSamsara"],
            "Estatus": row["EstatusPythonComparado"],
            "VelocidadMph": row[
                "SpeedMilesPerHour"
            ],
            "IsEcuSpeed": row["IsEcuSpeed"],
            "Latitud": row["Latitude"],
            "Longitud": row["Longitude"],
            "Direccion": row["ReverseGeo"],
            "Detalle": (
                "Clasificacion final de EnvioMain1"
            ),
        })

    return eventos


def agregar_historial_samsara(
    eventos,
    diferencias,
    token,
    margen_minutos,
):
    if diferencias.empty:
        return eventos

    fechas = pd.concat([
        diferencias["FechaExterna"],
        diferencias["FechaSamsara"],
    ]).dropna()

    if fechas.empty:
        raise ValueError(
            "No hay fechas validas para consultar historial."
        )

    inicio_mx = MX_TZ.localize(
        fechas.min().to_pydatetime()
    ) - timedelta(minutes=margen_minutos)
    fin_mx = MX_TZ.localize(
        fechas.max().to_pydatetime()
    ) + timedelta(minutes=margen_minutos)

    ids = diferencias[
        "SamsaraVehicleId"
    ].dropna().astype(str).unique().tolist()

    historial = consultar_historial_gps(
        token=token,
        vehicle_ids=ids,
        start_time=inicio_mx.astimezone(pytz.UTC),
        end_time=fin_mx.astimezone(pytz.UTC),
    )

    unidad_por_id = {
        str(row["SamsaraVehicleId"]): row["UnidadClave"]
        for _, row in diferencias.iterrows()
    }

    for vehiculo in historial:
        vehicle_id = str(vehiculo.get("id", ""))
        unidad = unidad_por_id.get(vehicle_id)

        if not unidad:
            continue

        for punto in vehiculo.get("gps", []) or []:
            fecha = pd.to_datetime(
                punto.get("time"),
                utc=True,
                errors="coerce",
            )

            if pd.isna(fecha):
                continue

            is_ecu_speed = punto.get(
                "isEcuSpeed"
            )
            velocidad = punto.get(
                "speedMilesPerHour"
            )
            reverse_geo = punto.get(
                "reverseGeo"
            ) or {}

            eventos.append({
                "Unidad": unidad,
                "SamsaraVehicleId": vehicle_id,
                "Fuente": "SAMSARA HISTORIAL",
                "FechaHoraMexico": (
                    fecha.tz_convert(MX_TZ)
                    .tz_localize(None)
                ),
                "Estatus": estatus_punto_samsara(
                    velocidad,
                    is_ecu_speed,
                ),
                "VelocidadMph": velocidad,
                "IsEcuSpeed": is_ecu_speed,
                "Latitud": punto.get("latitude"),
                "Longitud": punto.get("longitude"),
                "Direccion": reverse_geo.get(
                    "formattedLocation",
                    "",
                ),
                "Detalle": (
                    "Punto historico consultado en Samsara"
                ),
            })

    return eventos


def guardar_resultados(
    salida,
    resumen,
    eventos,
):
    salida.mkdir(
        parents=True,
        exist_ok=True,
    )

    ruta_resumen = (
        salida / "resumen_diferencias.csv"
    )
    ruta_linea = (
        salida / "linea_tiempo_samsara.csv"
    )
    ruta_transiciones = (
        salida / "transiciones_estatus_samsara.csv"
    )

    resumen.to_csv(
        ruta_resumen,
        index=False,
        encoding="utf-8-sig",
    )

    linea_tiempo = pd.DataFrame(
        eventos,
        columns=COLUMNAS_LINEA_TIEMPO,
    )

    if not linea_tiempo.empty:
        linea_tiempo = (
            linea_tiempo.drop_duplicates(
                subset=[
                    "Unidad",
                    "Fuente",
                    "FechaHoraMexico",
                ]
            )
            .sort_values(
                by=[
                    "Unidad",
                    "FechaHoraMexico",
                    "Fuente",
                ]
            )
        )

    linea_tiempo.to_csv(
        ruta_linea,
        index=False,
        encoding="utf-8-sig",
    )

    historial = linea_tiempo[
        linea_tiempo["Fuente"].eq(
            "SAMSARA HISTORIAL"
        )
    ].copy()

    if historial.empty:
        historial = linea_tiempo[
            linea_tiempo["Fuente"].eq(
                "SAMSARA ACTUAL"
            )
        ].copy()

    if historial.empty:
        transiciones = pd.DataFrame(
            columns=COLUMNAS_TRANSICIONES
        )
    else:
        historial = historial.sort_values(
            by=[
                "Unidad",
                "FechaHoraMexico",
            ]
        )
        cambio = historial.groupby(
            "Unidad"
        )["Estatus"].transform(
            lambda serie: serie.ne(
                serie.shift()
            )
        )
        transiciones = historial[
            cambio
        ].copy()
        siguiente = transiciones.groupby(
            "Unidad"
        )["FechaHoraMexico"].shift(-1)
        transiciones[
            "DuracionHastaSiguienteMinutos"
        ] = (
            pd.to_datetime(siguiente)
            - pd.to_datetime(
                transiciones["FechaHoraMexico"]
            )
        ).dt.total_seconds().div(60).round(2)
        transiciones = transiciones[
            COLUMNAS_TRANSICIONES
        ]

    transiciones.to_csv(
        ruta_transiciones,
        index=False,
        encoding="utf-8-sig",
    )

    return (
        ruta_resumen,
        ruta_linea,
        ruta_transiciones,
        linea_tiempo,
        transiciones,
    )


def main():
    args = argumentos()
    ruta_samsara = Path(args.samsara).resolve()
    ruta_externo = Path(args.externo).resolve()
    ruta_salida = Path(args.salida).resolve()

    if not ruta_samsara.exists():
        raise FileNotFoundError(ruta_samsara)

    if not ruta_externo.exists():
        raise FileNotFoundError(ruta_externo)

    externo = leer_reporte_externo(
        ruta_externo
    )
    samsara = leer_reporte_samsara(
        ruta_samsara
    )
    diferencias = detectar_diferencias(
        externo,
        samsara,
    )
    resumen = construir_resumen(
        diferencias
    )
    eventos = construir_linea_tiempo_base(
        diferencias
    )

    historial_consultado = False

    if not args.sin_historial and not diferencias.empty:
        load_dotenv(
            REPO_ROOT / ".env"
        )
        token = os.getenv("SAMSARA_API_TOKEN")

        if not token:
            raise RuntimeError(
                "Falta SAMSARA_API_TOKEN en .env"
            )

        eventos = agregar_historial_samsara(
            eventos=eventos,
            diferencias=diferencias,
            token=token,
            margen_minutos=args.margen_minutos,
        )
        historial_consultado = True

    (
        ruta_resumen,
        ruta_linea,
        ruta_transiciones,
        linea_tiempo,
        transiciones,
    ) = guardar_resultados(
        salida=ruta_salida,
        resumen=resumen,
        eventos=eventos,
    )

    print(
        f"Unidades con estatus diferente: "
        f"{len(resumen)}"
    )
    print(
        "Unidades: "
        + (
            ", ".join(
                resumen["Unidad"].astype(str)
            )
            if not resumen.empty
            else "ninguna"
        )
    )
    print(
        f"Historial Samsara consultado: "
        f"{historial_consultado}"
    )
    print(
        f"Eventos en linea de tiempo: "
        f"{len(linea_tiempo)}"
    )
    print(
        f"Transiciones de estatus: "
        f"{len(transiciones)}"
    )
    print(f"Resumen: {ruta_resumen}")
    print(f"Linea de tiempo: {ruta_linea}")
    print(f"Transiciones: {ruta_transiciones}")


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(
            f"ERROR: {error}",
            file=sys.stderr,
        )
        sys.exit(1)
