"""Diagnostico de solo lectura del cruce Samsara vs Logitrack.

Consulta las unidades del reporte, aplica la misma decision que EnvioMain y
exporta un Excel con distancia, desfase y estatus de cada unidad. No envia nada.

Uso:
    .\\.venv\\Scripts\\python.exe .\\scripts\\diagnosticar_samsara_logitrack.py
    .\\.venv\\Scripts\\python.exe .\\scripts\\diagnosticar_samsara_logitrack.py --reporte "Reporte EC-05"
"""
from __future__ import annotations

import argparse
from datetime import datetime
import os
from pathlib import Path
import sys

import pandas as pd
import pytz
import requests

BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))

import EnvioMain as envio  # noqa: E402
from LogitrackEstatus import fetch_last_status, request_access_token  # noqa: E402

COLUMNAS = [
    "Unidad", "Estatus Samsara Actual", "Estatus Logitrack", "Estatus",
    "Velocidad Mph", "Velocidad Logitrack", "Fecha GPS", "Fecha Logitrack",
    "Desfase Lecturas Min", "Distancia GPS Metros", "Distancia Esperada Metros",
    "Ubicación Coincide", "Doble Comprobación Actual", "Fuente Confirmación",
    "Motivo Decisión",
]


def describir(serie: pd.Series) -> str:
    if serie.empty:
        return "n=0"
    return (
        f"n={len(serie)} promedio={serie.mean():.0f} m mediana={serie.median():.0f} m "
        f"p95={serie.quantile(.95):.0f} m max={serie.max():.0f} m"
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--reporte", default="Reporte EC-05")
    args = parser.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    config = envio.cargar_configuracion(envio.DEFAULT_CONFIG_PATH)
    reporte = next(r for r in config["reportes"] if r["nombre"] == args.reporte)
    settings = reporte.get("telemetria_logitrack") or {}
    token = os.getenv("SAMSARA_API_TOKEN") or os.getenv("SAMSARA_TOKEN", "").removeprefix("Bearer ")
    sesion = envio.crear_sesion_samsara(token)
    now_mx = datetime.now(pytz.timezone(config.get("zona_horaria", envio.DEFAULT_TIMEZONE)))

    ids, _ = envio.obtener_ids_etiquetas_reporte_remoto(
        sesion, reporte, envio.cargar_cache_catalogo(envio.DEFAULT_TAG_CATALOG_PATH)
    )
    vehicles = envio.obtener_vehiculos(sesion, ids, reporte.get("tipo_filtro_etiqueta", "tagIds"))
    # Sin filtros de geocerca para medir todas las unidades del reporte.
    results, excluidas = envio.procesar_vehiculos(vehicles, now_mx, {"gps_max_minutos": 60, "excluir_speed_cero_sin_ecu": False}, {})

    with requests.Session() as sesion_logitrack:
        api_url = os.getenv("LOGITRACK_API_URL", envio.LOGITRACK_DEFAULT_API_URL).strip().rstrip("/")
        token_url = os.getenv("LOGITRACK_TOKEN_URL", envio.LOGITRACK_DEFAULT_TOKEN_URL).strip()
        token_logitrack = request_access_token(sesion_logitrack, token_url, 60)
        statuses = fetch_last_status(sesion_logitrack, api_url, token_logitrack, 60)

    results = envio.enriquecer_con_logitrack(results, statuses, now_mx, settings)
    filtros = envio.combinar_filtros(
        config.get("filtros_base") or {}, reporte, config.get("perfiles_filtros") or {}
    )
    etiquetas_geocercas = envio.resolver_etiquetas_samsara(
        sesion, filtros.get("etiquetas_geocercas_excluidas") or [], {}
    )
    geometrias = envio.obtener_geometrias_geocercas(
        sesion, etiquetas_geocercas.values(), filtros.get("geocercas_especiales_ids") or []
    )
    rescatadas, excluidas = envio.rescatar_con_logitrack(
        excluidas, statuses, now_mx, settings, geometrias
    )
    df = pd.DataFrame(results + rescatadas).reindex(columns=COLUMNAS)

    frescas = df[df["Desfase Lecturas Min"].notna() & (df["Estatus Logitrack"] != "DESACTUALIZADO")]
    detenidas = frescas[frescas["Doble Comprobación Actual"] == "SI"]
    movimiento = frescas[frescas["Estatus Logitrack"] == "RUTA"]
    print(f"\nSamsara={len(vehicles)} Logitrack={len(statuses)} GPS Samsara viejo={len(excluidas) + len(rescatadas)}")
    print("Distancia ambas detenidas:", describir(detenidas["Distancia GPS Metros"].dropna()))
    print("Distancia en movimiento:  ", describir(movimiento["Distancia GPS Metros"].dropna()))
    print("\nEstatus resultante (antes del historial Samsara):")
    print(df["Estatus"].value_counts().to_string())

    salida = BASE_DIR / "outputs" / f"diagnostico_samsara_logitrack_{now_mx:%Y%m%d_%H%M%S}.xlsx"
    salida.parent.mkdir(parents=True, exist_ok=True)
    for columna in ("Fecha GPS", "Fecha Logitrack"):
        df[columna] = pd.to_datetime(df[columna], errors="coerce", utc=True).dt.tz_convert(now_mx.tzinfo).dt.tz_localize(None)
    df.to_excel(salida, index=False)
    print(f"\nDetalle: {salida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
