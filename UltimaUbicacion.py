import requests
import pandas as pd
import os
from datetime import datetime, timezone
from pathlib import Path
from zoneinfo import ZoneInfo
import unicodedata
from dotenv import load_dotenv

# ===== Config =====
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")
URL = "https://api.samsara.com/fleet/vehicles/stats/feed"
PARAMS = {"types": "gps,engineStates"}   # gps + estado de motor

MX_TZ = ZoneInfo("America/Mexico_City")

# ===== Utilidades =====
def utc_to_mx(ts_utc: str) -> str:
    """Convierte ISO-8601 UTC a string local CDMX 'YYYY-MM-DD HH:MM:SS'."""
    try:
        dt_utc = datetime.fromisoformat(ts_utc.replace("Z", "+00:00"))
    except ValueError:
        dt_naive = datetime.fromisoformat(ts_utc.split(".")[0])
        dt_utc = dt_naive.replace(tzinfo=timezone.utc)
    return dt_utc.astimezone(MX_TZ).strftime("%Y-%m-%d %H:%M:%S")

def norm(s):
    """Normaliza texto a NFC y maneja None."""
    if s is None:
        return ""
    return unicodedata.normalize("NFC", str(s))

def normalize_engine_state(raw: str) -> str:
    """Normaliza estado de motor a on/off/idle cuando es posible."""
    if not raw:
        return ""
    s = str(raw).strip().lower()
    if "idle" in s:
        return "idle"
    if s in ("on", "engine_on", "engineon", "running"):
        return "on"
    if s in ("off", "engine_off", "engineoff", "stopped"):
        return "off"
    return s  # fallback

def extract_latest_engine_state(vehicle: dict) -> str:
    """Obtiene el estado de motor más reciente del arreglo engineStates."""
    states = vehicle.get("engineStates") or []
    if not states:
        return ""
    latest = max(states, key=lambda e: e.get("time", ""))
    raw = latest.get("state")
    if raw is None:
        raw = latest.get("value")
    if raw is None:
        raw = latest.get("engineState")
    return normalize_engine_state(raw)

# ===== Main (una sola llamada y exporta a .xlsx) =====
def main():
    token = os.getenv("SAMSARA_API_TOKEN")
    if not token:
        raise RuntimeError("Falta SAMSARA_API_TOKEN en .env")
    headers = {
        "Accept": "application/json",
        "Authorization": f"Bearer {token}",
    }
    # Llamada única
    resp = requests.get(URL, headers=headers, params=PARAMS, timeout=60)
    resp.raise_for_status()
    payload = resp.json()
    data = payload.get("data", []) or []

    rows = []
    for v in data:
        vid = v.get("id")
        name = norm(v.get("name"))
        # ext = v.get("externalIds") or {}
        # serial = ext.get("samsara.serial", "")
        # vin = ext.get("samsara.vin", "")

        engine_state = extract_latest_engine_state(v)

        gps_list = v.get("gps") or []
        if not gps_list:
            continue

        # Toma el ping más reciente por 'time'
        g = max(gps_list, key=lambda x: x.get("time", ""))
        t_utc = g.get("time", "")
        t_mx = utc_to_mx(t_utc) if t_utc else ""

        rev = g.get("reverseGeo", {}) or {}
        loc = norm(rev.get("formattedLocation"))

        rows.append({
            "ID": vid,
            "Unidad": name,
            # "VIN": vin,
            # "Serial": serial,
            # "Tiempo_UTC": t_utc,
            "Tiempo_MX": t_mx,
            "Latitud": g.get("latitude"),
            "Longitud": g.get("longitude"),
            "Locación": loc,
            "engine_state": engine_state,           # on/off/idle (si disponible)
            "isEcuSpeed": g.get("isEcuSpeed"),      # True/False si la velocidad viene de ECU
            # "Velocidad_mph": g.get("speedMilesPerHour"),
            # "Heading_deg": g.get("headingDegrees"),
        })

    # Exportar a Excel
    df = pd.DataFrame(rows)
    # Orden visual de columnas (si alguna no existe, se ignora)
    cols = [
        "ID", "Unidad",
        "Tiempo_MX",
        "Latitud", "Longitud", "Locación",
        "engine_state", "isEcuSpeed"
    ]
    df = df[[c for c in cols if c in df.columns]]

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    outfile = f"feed_ultimas_ubicaciones_{ts}.xlsx"
    # Un solo sheet
    with pd.ExcelWriter(outfile, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="ultima_ubicacion", index=False)

    print(f"Listo: {len(df)} unidades -> {outfile}")

if __name__ == "__main__":
    main()
