# -*- coding: utf-8 -*-
from pathlib import Path
import os
import unicodedata
import requests
import pyodbc
from dotenv import load_dotenv
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
import logging

# =========================
# Logging (archivo + consola)
# =========================
LOG_FILE = "ubicaciones_update.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler()
    ],
)

# =========================
# Carga .env
# =========================
BASE_DIR = Path(__file__).parent
ENV_PATH = BASE_DIR / ".env"
load_dotenv(dotenv_path=ENV_PATH)

SAMSARA_TOKEN = os.getenv("SAMSARA_TOKEN")
SQL_SERVER     = os.getenv("SQL_SERVER")
SQL_DATABASE   = os.getenv("SQL_DATABASE")
SQL_USER       = os.getenv("SQL_USER")
SQL_PASSWORD   = os.getenv("SQL_PASSWORD")
SQL_DRIVER_ENV = os.getenv("SQL_DRIVER")  # ej. "ODBC Driver 17 for SQL Server"

# =========================
# Helpers ODBC
# =========================
def pick_sql_driver(preferred=None):
    available = [d.strip() for d in pyodbc.drivers()]
    logging.info(f"Drivers ODBC instalados: {available}")

    candidates = []
    if preferred:
        candidates.append(preferred)

    candidates += [
        "ODBC Driver 18 for SQL Server",
        "ODBC Driver 17 for SQL Server",
        "ODBC Driver 13 for SQL Server",
        "SQL Server Native Client 11.0",
        "SQL Server",
    ]

    for name in candidates:
        if name in available:
            return name

    raise RuntimeError("No se encontró un driver ODBC de SQL Server instalado.")


def build_conn_str(server, database, user=None, password=None, driver=None, trust_cert=True):
    drv = driver or pick_sql_driver(SQL_DRIVER_ENV)

    cs = f"DRIVER={{{drv}}};SERVER={server};DATABASE={database};Encrypt=yes;"

    if trust_cert:
        cs += "TrustServerCertificate=yes;"

    if user and password:
        cs += f"UID={user};PWD={password};"
    else:
        cs += "Trusted_Connection=yes;"

    logging.info(f"Usando driver ODBC: {drv}")
    return cs


def normalize_token(raw: str) -> str:
    raw = (raw or "").strip()
    return raw if raw.startswith("Bearer ") else f"Bearer {raw}"


# =========================
# Config API
# =========================
URL = "https://api.samsara.com/fleet/vehicles/stats/feed"
PARAMS = {"types": "gps,engineStates"}   # última ubicación + estado motor
HEADERS = {
    "Accept": "application/json",
    "Authorization": normalize_token(SAMSARA_TOKEN),
}

# =========================
# Utilidades
# =========================
MX_TZ = ZoneInfo("America/Mexico_City")


def utc_to_mx_dt(ts_utc: str) -> datetime:
    """Convierte ISO-8601 UTC a datetime naive en hora CDMX para SQL DATETIME."""
    try:
        dt_utc = datetime.fromisoformat(ts_utc.replace("Z", "+00:00"))
    except ValueError:
        dt_naive = datetime.fromisoformat(ts_utc.split(".")[0])
        dt_utc = dt_naive.replace(tzinfo=timezone.utc)

    return dt_utc.astimezone(MX_TZ).replace(tzinfo=None)


def norm(s):
    if s is None:
        return ""
    return unicodedata.normalize("NFC", str(s)).strip()


def normalize_engine_state(raw: str) -> str:
    if not raw:
        return ""

    s = str(raw).strip().lower()

    if "idle" in s:
        return "idle"

    if s in ("on", "engine_on", "engineon", "running"):
        return "on"

    if s in ("off", "engine_off", "engineoff", "stopped"):
        return "off"

    return s


def extract_latest_engine_state(vehicle: dict) -> str:
    states = vehicle.get("engineStates") or []

    if not states:
        return ""

    latest = max(states, key=lambda e: e.get("time", ""))
    raw = latest.get("state") or latest.get("value") or latest.get("engineState")

    return normalize_engine_state(raw)


# =========================
# 1) Llamada API y armado de filas
# =========================
def fetch_rows():
    logging.info(f"GET {URL} params={PARAMS}")

    resp = requests.get(URL, headers=HEADERS, params=PARAMS, timeout=60)

    if resp.status_code == 401:
        logging.error("401 Unauthorized — revisa token/scopes/región.")
        logging.error(resp.text[:800])

    resp.raise_for_status()

    payload = resp.json()
    data = payload.get("data", []) or []

    logging.info(f"Unidades recibidas: {len(data)}")

    rows = []  # (Unidad, TiempoMX_dt, Latitud, Longitud, Locacion)

    for v in data:
        unidad = norm(v.get("name"))

        if not unidad:
            continue

        gps_list = v.get("gps") or []

        if not gps_list:
            continue

        # Ping más reciente
        g = max(gps_list, key=lambda x: x.get("time", ""))

        t_utc = g.get("time", "")

        if not t_utc:
            continue

        t_mx_dt = utc_to_mx_dt(t_utc)

        lat = g.get("latitude")
        lon = g.get("longitude")
        loc = norm((g.get("reverseGeo") or {}).get("formattedLocation") or "X")
        rows.append((unidad, t_mx_dt, lat, lon, loc))

    logging.info(f"Filas preparadas para staging: {len(rows)}")

    return rows


# =========================
# 2) Actualización SQL + ejecución SP checkcall
# =========================
def update_sql_with_temp(rows):
    if not rows:
        logging.info("No hay filas para actualizar.")
        return

    conn_str = build_conn_str(
        server=SQL_SERVER,
        database=SQL_DATABASE,
        user=SQL_USER or None,
        password=SQL_PASSWORD or None,
        driver=SQL_DRIVER_ENV,
        trust_cert=True
    )

    with pyodbc.connect(conn_str, autocommit=False) as cn:
        cur = cn.cursor()

        try:
            # 2.1) Crear tabla temporal y tablas para capturar OUTPUT
            cur.execute("""
IF OBJECT_ID('tempdb..#LatestGPS_Stg') IS NOT NULL DROP TABLE #LatestGPS_Stg;
CREATE TABLE #LatestGPS_Stg(
    Unidad         NVARCHAR(128) NOT NULL,
    TiempoMX_dt    DATETIME      NOT NULL,
    Latitud        FLOAT         NULL,
    Longitud       FLOAT         NULL,
    Locacion       NVARCHAR(400) NULL
);

IF OBJECT_ID('tempdb..#MP_Updated') IS NOT NULL DROP TABLE #MP_Updated;
CREATE TABLE #MP_Updated(
    Unidad NVARCHAR(128),
    OldDate DATETIME NULL,
    NewDate DATETIME NULL,
    Locacion NVARCHAR(400) NULL,
    Latitud FLOAT NULL,
    Longitud FLOAT NULL
);

IF OBJECT_ID('tempdb..#TR_Updated') IS NOT NULL DROP TABLE #TR_Updated;
CREATE TABLE #TR_Updated(
    Unidad NVARCHAR(128),
    OldDate DATETIME NULL,
    NewDate DATETIME NULL,
    Locacion NVARCHAR(400) NULL,
    Latitud FLOAT NULL,
    Longitud FLOAT NULL
);
            """)

            # 2.2) Insert masivo a la #temp
            cur.fast_executemany = True
            cur.executemany(
                """
                INSERT INTO #LatestGPS_Stg
                    (Unidad, TiempoMX_dt, Latitud, Longitud, Locacion)
                VALUES (?, ?, ?, ?, ?)
                """,
                rows
            )
            logging.info("Ejmplos de filas cargadas en #LatestGPS_Stg:")
            for row in rows[:10]:
                logging.info(
                    f"[STAGING] Unidad={row[0]} | TiempoMX_dt={row[1]} | Latitud={row[2]} | Longitud={row[3]} | Locacion={row[4]}"
                )
            cur.execute("SELECT COUNT(*) FROM #LatestGPS_Stg;")
            sttaging_count = cur.fetchone()[0]
            logging.info(f"Filas en #LatestGPS_Stg: {sttaging_count}")
            # 2.3) Actualiza manpowerprofile
            cur.execute("""
;WITH Parsed AS (
    SELECT
        Unidad,
        TiempoMX_dt,
        Latitud,
        Longitud,
        Locacion,
        ROW_NUMBER() OVER (PARTITION BY Unidad ORDER BY TiempoMX_dt DESC) AS rn
    FROM #LatestGPS_Stg
),
Latest AS (
    SELECT
        Unidad,
        TiempoMX_dt,
        Latitud,
        Longitud,
        Locacion
    FROM Parsed
    WHERE rn = 1
)
UPDATE mp
SET
    mp.mpp_gps_date      = L.TiempoMX_dt,
    mp.mpp_gps_desc      = L.Locacion,
    mp.mpp_gps_latitude  = L.Latitud,
    mp.mpp_gps_longitude = L.Longitud
OUTPUT
    inserted.mpp_tractornumber,
    deleted.mpp_gps_date,
    inserted.mpp_gps_date,
    inserted.mpp_gps_desc,
    inserted.mpp_gps_latitude,
    inserted.mpp_gps_longitude
INTO #MP_Updated
    (Unidad, OldDate, NewDate, Locacion, Latitud, Longitud)
FROM dbo.manpowerprofile AS mp
JOIN Latest AS L
    ON L.Unidad = mp.mpp_tractornumber
WHERE L.TiempoMX_dt > ISNULL(mp.mpp_gps_date, '19000101');
            """)
            

            # 2.4) Actualiza tractorprofile
            cur.execute("""
;WITH Parsed AS (
    SELECT
        Unidad,
        TiempoMX_dt,
        Latitud,
        Longitud,
        Locacion,
        ROW_NUMBER() OVER (PARTITION BY Unidad ORDER BY TiempoMX_dt DESC) AS rn
    FROM #LatestGPS_Stg
),
Latest AS (
    SELECT
        Unidad,
        TiempoMX_dt,
        Latitud,
        Longitud,
        Locacion
    FROM Parsed
    WHERE rn = 1
)
UPDATE tr
SET
    tr.trc_gps_date      = L.TiempoMX_dt,
    tr.trc_gps_desc      = L.Locacion,
    tr.trc_gps_latitude  = L.Latitud,
    tr.trc_gps_longitude = L.Longitud
OUTPUT
    inserted.trc_number,
    deleted.trc_gps_date,
    inserted.trc_gps_date,
    inserted.trc_gps_desc,
    inserted.trc_gps_latitude,
    inserted.trc_gps_longitude
INTO #TR_Updated
    (Unidad, OldDate, NewDate, Locacion, Latitud, Longitud)
FROM dbo.tractorprofile AS tr
JOIN Latest AS L
    ON L.Unidad = tr.trc_number
WHERE L.TiempoMX_dt > ISNULL(tr.trc_gps_date, '19000101');
            """)

            # 2.5) Traer listas de actualizados
            cur.execute("""
                SELECT
                    Unidad,
                    OldDate,
                    NewDate,
                    Locacion,
                    Latitud,
                    Longitud
                FROM #MP_Updated
                ORDER BY NewDate DESC;
            """)
            mp_rows = cur.fetchall()

            cur.execute("""
                SELECT
                    Unidad,
                    OldDate,
                    NewDate,
                    Locacion,
                    Latitud,
                    Longitud
                FROM #TR_Updated
                ORDER BY NewDate DESC;
            """)
            tr_rows = cur.fetchall()
            
            cur.execute("""
                SELECT 
                    mpp_tractornumber,
                    mpp_gps_date,
                    mpp_gps_desc,
                    mpp_gps_latitude,
                    mpp_gps_longitude
                FROM dbo.manpowerprofile
                WHERE mpp_tractornumber IN (
                    SELECT TOP 10 Unidad FROM #MP_Updated ORDER BY NewDate DESC
                );
            """)

            debug_mp_before_sp = cur.fetchall()

            logging.info("MP después del UPDATE, antes del SP:")
            for r in debug_mp_before_sp:
                logging.info(
                    f"[MP BEFORE SP] Unidad={r[0]} | Fecha={r[1]} | Desc={r[2]} | Lat={r[3]} | Lon={r[4]}"
                )
            # 2.6) Insertar checkcall usando SP
            # El SP actual recibe:
            # @P_IDOPERA varchar(8),
            # @P_fechamov datetime,
            # @P_ubicacion varchar(100),
            # @P_unidad varchar(10)
            # 2.6) Insertar checkcall usando SP
            checkcall_insertados = 0
            
            for r in tr_rows:
                unidad, oldd, newd, loc, lat, lon = r
            
                unidad_sp = unidad[:10] if unidad else None
                ubicacion_sp = loc[:100] if loc else None

                logging.info(
                    "[CHECKCALL/SP] Insertando | "
                    f"P_IDOPERA={None} | "
                    f"P_fechamov={newd} | "
                    f"P_ubicacion={ubicacion_sp} | "
                    f"P_unidad={unidad_sp} | "
                    f"P_TIPOMOV=GPS | "
                    f"P_latitud={lat} | "
                    f"P_longitud={lon}"
                )
                print(
                    "[CHECKCALL/SP]",
                    {
                        "P_IDOPERA": None,
                        "P_fechamov": newd,
                        "P_ubicacion": ubicacion_sp,
                        "P_unidad": unidad_sp,
                        "P_TIPOMOV": "GPS",
                        "P_latitud": lat,
                        "P_longitud": lon,
                    }
                )
                cur.execute(
                    """
                    EXEC dbo.sp_inserta_checkcal_RC_JR
                        @P_IDOPERA = ?,
                        @P_fechamov = ?,
                        @P_ubicacion = ?,
                        @P_unidad = ?,
                        @P_TIPOMOV = ?,
                        @P_latitud = ?,
                        @P_longitud = ?
                    """,
                    None,
                    newd,
                    ubicacion_sp,
                    unidad_sp,
                    "GPS",
                    lat,
                    lon
                )
            
            cur.execute("""
                SELECT TOP 10
                    ckc_number,
                    ckc_tractor,
                    ckc_date,
                    ckc_event,
                    ckc_city,
                    ckc_cityname,
                    ckc_state,
                    ckc_zip,
                    ckc_comment,
                    ckc_commentlarge,
                    ckc_latseconds,
                    ckc_longseconds
                FROM tmwSuite..checkcall
                WHERE ckc_tractor IN (
                    SELECT TOP 10 Unidad FROM #TR_Updated ORDER BY NewDate DESC
                )
                ORDER BY ckc_number DESC;
            """)
            
            debug_ckc = cur.fetchall()
            
            logging.info("Checkcall insertado por el SP:")
            for r in debug_ckc:
                logging.info(
                    f"[CKC] number={r[0]} | unidad={r[1]} | fecha={r[2]} | event={r[3]} | "
                    f"city={r[4]} | cityname={r[5]} | state={r[6]} | zip={r[7]} | "
                    f"comment={r[8]} | commentlarge={r[9]} | lat={r[10]} | lon={r[11]}"
                )
                
                cur.execute("""
                    SELECT 
                        mpp_tractornumber,
                        mpp_gps_date,
                        mpp_gps_desc,
                        mpp_gps_latitude,
                        mpp_gps_longitude
                    FROM dbo.manpowerprofile
                    WHERE mpp_tractornumber IN (
                        SELECT TOP 10 Unidad FROM #MP_Updated ORDER BY NewDate DESC
                    );
                """)
                
                debug_mp_after_sp = cur.fetchall()
                
                logging.info("MP después del SP, antes del COMMIT:")
                for r in debug_mp_after_sp:
                    logging.info(
                        f"[MP AFTER SP] Unidad={r[0]} | Fecha={r[1]} | Desc={r[2]} | Lat={r[3]} | Lon={r[4]}"
                    )
    
                checkcall_insertados += 1

            # 2.7) Confirmar transacción completa
            cn.commit()

            # 2.8) Logs legibles
            logging.info(f"MP actualizadas: {len(mp_rows)}")
            for r in mp_rows:
                unidad, oldd, newd, loc, lat, lon = r
                logging.info(
                    f"[MP] {unidad} | {oldd} -> {newd} | {loc} | ({lat}, {lon})"
                )

            logging.info(f"TR actualizadas: {len(tr_rows)}")
            for r in tr_rows:
                unidad, oldd, newd, loc, lat, lon = r
                logging.info(
                    f"[TR] {unidad} | {oldd} -> {newd} | {loc} | ({lat}, {lon})"
                )

            logging.info(f"Checkcalls insertados: {checkcall_insertados}")

            print(
                f"Actualizadas MP: {len(mp_rows)} | "
                f"TR: {len(tr_rows)} | "
                f"Checkcalls: {checkcall_insertados}"
            )

            if mp_rows:
                print("Ejemplos MP:", mp_rows[:3])

            if tr_rows:
                print("Ejemplos TR:", tr_rows[:3])

        except Exception as e:
            cn.rollback()
            logging.exception("Error durante la actualización SQL. Se hizo rollback.")
            raise


# =========================
# Main
# =========================
def main():
    rows = fetch_rows()
    logging.info(f"Filas obtenidas de Samsara: {len(rows)}")

    update_sql_with_temp(rows)

    logging.info("Actualización en SQL completada.")


if __name__ == "__main__":
    main()