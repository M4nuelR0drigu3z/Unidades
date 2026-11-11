"""Catalogos de TMWSuite: unidades (tractores) y operadores (manpowerprofile).

Lo usa api_telemetria.py (GET /telemetria/consultarUnidades y el operador de consultarEstatus).
Solo lectura: consultas fijas, sin datos del cliente en el SQL; los filtros se aplican en
memoria sobre el catalogo en cache.
"""
from __future__ import annotations

import logging
import os
import re
import threading
import time
from typing import Any

CACHE_SEGUNDOS_DEFAULT = 600
DRIVERS_PREFERIDOS = [
    "ODBC Driver 18 for SQL Server",
    "ODBC Driver 17 for SQL Server",
    "ODBC Driver 13 for SQL Server",
    "SQL Server Native Client 11.0",
    "SQL Server",
]

CONSULTA_UNIDADES = """
SELECT
    LTRIM(RTRIM(tp.trc_number)) AS unidad,
    LTRIM(RTRIM(tp.trc_division)) AS division,
    TRY_CONVERT(int, lf.abbr) AS flota_id,
    COALESCE(fd.Nombre_Flota, lf.name) AS flota,
    LTRIM(RTRIM(tp.trc_licnum)) AS placas,
    LTRIM(RTRIM(tp.trc_status)) AS estatus_tmw,
    CASE
        WHEN LTRIM(RTRIM(tp.trc_fleet)) = '17' THEN 0
        WHEN COALESCE(LTRIM(RTRIM(tp.trc_division)), '') IN ('', 'UNK') THEN 0
        ELSE 1
    END AS activa
FROM [TMWSuite].dbo.tractorprofile AS tp
LEFT JOIN [TMWSuite].dbo.labelfile AS lf
    ON lf.labeldefinition = 'Fleet' AND lf.abbr = tp.trc_fleet
LEFT JOIN [TMWSuite].dbo.VW_Fleet_Division AS fd
    ON fd.Id_Flota = tp.trc_fleet
WHERE tp.trc_number <> 'UNKNOWN'
ORDER BY tp.trc_number
"""

# El usuario del conductor en Samsara es el mpp_id de TMW.
CONSULTA_OPERADORES = """
SELECT
    UPPER(LTRIM(RTRIM(mpp_id))) AS mpp_id,
    LTRIM(RTRIM(mpp_firstname)) AS nombre,
    LTRIM(RTRIM(mpp_middlename)) AS segundo_nombre,
    LTRIM(RTRIM(mpp_lastname)) AS apellido,
    LTRIM(RTRIM(mpp_currentphone)) AS telefono
FROM [TMWSuite].dbo.manpowerprofile
WHERE NULLIF(LTRIM(RTRIM(mpp_id)), '') IS NOT NULL
"""

_cache: dict[str, Any] = {"momento": None, "unidades": []}
_candado = threading.Lock()
_cache_operadores: dict[str, Any] = {"momento": None, "operadores": {}}
_candado_operadores = threading.Lock()


def _driver_sql() -> str:
    import pyodbc

    configurado = os.getenv("SQL_DRIVER", "").strip()
    if configurado:
        return configurado
    instalados = {d.strip() for d in pyodbc.drivers()}
    for nombre in DRIVERS_PREFERIDOS:
        if nombre in instalados:
            return nombre
    raise RuntimeError("No hay un driver ODBC de SQL Server instalado")


def conectar():
    import pyodbc

    faltantes = [v for v in ("SQL_SERVER", "SQL_USER", "SQL_PASSWORD") if not os.getenv(v)]
    if faltantes:
        raise RuntimeError(f"Faltan variables de SQL: {', '.join(faltantes)}")
    driver = _driver_sql()
    cifrado = "Optional" if "Driver 18" in driver else "no"
    conexion = pyodbc.connect(
        f"DRIVER={{{driver}}};"
        f"SERVER={os.environ['SQL_SERVER']};"
        f"DATABASE={os.getenv('SQL_DATABASE') or 'TMWSuite'};"
        f"UID={os.environ['SQL_USER']};"
        f"PWD={os.environ['SQL_PASSWORD']};"
        f"Encrypt={cifrado};TrustServerCertificate=yes;Connection Timeout=15;"
        "ApplicationIntent=ReadOnly;",
        timeout=15,
    )
    # Columnas varchar con acentos en Latin1 (p. ej. "México").
    conexion.setdecoding(pyodbc.SQL_CHAR, encoding="cp1252")
    return conexion


def _texto(valor: Any) -> str:
    return str(valor).strip() if valor is not None else ""


def leer_unidades() -> list[dict[str, Any]]:
    with conectar() as conexion:
        cursor = conexion.cursor()
        cursor.execute(CONSULTA_UNIDADES)
        columnas = [c[0] for c in cursor.description]
        filas = [dict(zip(columnas, fila)) for fila in cursor.fetchall()]
    return [
        {
            "unidad": _texto(fila["unidad"]),
            "division": _texto(fila["division"]),
            "flota": (
                {"id": fila["flota_id"], "nombre": _texto(fila["flota"])}
                if fila["flota_id"] is not None or fila["flota"] else None
            ),
            "placas": _texto(fila["placas"]),
            "activa": bool(fila["activa"]),
        }
        for fila in filas
    ]


def unidades_con_cache() -> list[dict[str, Any]]:
    """Catalogo completo; se vuelve a leer de SQL cada TELEMETRIA_CACHE_UNIDADES_SEGUNDOS (600)."""
    ttl = int(os.getenv("TELEMETRIA_CACHE_UNIDADES_SEGUNDOS", CACHE_SEGUNDOS_DEFAULT))
    with _candado:
        momento = _cache["momento"]
        if momento is not None and time.monotonic() - momento < ttl:
            return _cache["unidades"]
        unidades = leer_unidades()
        _cache.update(momento=time.monotonic(), unidades=unidades)
        logging.info("Catalogo de unidades leido de SQL: %s unidades", len(unidades))
        return unidades


def limpiar_telefono(valor: Any) -> str | None:
    """Solo digitos; algunos vienen con espacios o sin lada (8 digitos) y se dejan asi."""
    digitos = re.sub(r"\D", "", _texto(valor))
    return digitos or None


def leer_operadores() -> dict[str, dict[str, Any]]:
    """mpp_id -> {"nombre", "telefono"}."""
    with conectar() as conexion:
        cursor = conexion.cursor()
        cursor.execute(CONSULTA_OPERADORES)
        columnas = [c[0] for c in cursor.description]
        filas = [dict(zip(columnas, fila)) for fila in cursor.fetchall()]
    operadores = {}
    for fila in filas:
        partes = (fila["nombre"], fila["segundo_nombre"], fila["apellido"])
        operadores[_texto(fila["mpp_id"])] = {
            "nombre": " ".join(_texto(p) for p in partes if _texto(p)),
            "telefono": limpiar_telefono(fila["telefono"]),
        }
    return operadores


def operadores_con_cache() -> dict[str, dict[str, Any]]:
    """Operadores de TMW; se vuelven a leer cada TELEMETRIA_CACHE_UNIDADES_SEGUNDOS (600)."""
    ttl = int(os.getenv("TELEMETRIA_CACHE_UNIDADES_SEGUNDOS", CACHE_SEGUNDOS_DEFAULT))
    with _candado_operadores:
        momento = _cache_operadores["momento"]
        if momento is not None and time.monotonic() - momento < ttl:
            return _cache_operadores["operadores"]
        operadores = leer_operadores()
        _cache_operadores.update(momento=time.monotonic(), operadores=operadores)
        logging.info("Catalogo de operadores leido de SQL: %s operadores", len(operadores))
        return operadores


def filtrar_unidades(
    unidades: list[dict[str, Any]],
    unidad: str | None = None,
    division: str | None = None,
    flota: int | None = None,
    incluir_bajas: bool = False,
) -> list[dict[str, Any]]:
    def lista(texto: str | None) -> set[str]:
        return {x.strip().upper() for x in (texto or "").split(",") if x.strip()}

    unidades_buscadas, divisiones = lista(unidad), lista(division)
    return [
        u for u in unidades
        if (incluir_bajas or u["activa"])
        and (not unidades_buscadas or u["unidad"].upper() in unidades_buscadas)
        and (not divisiones or u["division"].upper() in divisiones)
        and (flota is None or (u["flota"] or {}).get("id") == flota)
    ]
