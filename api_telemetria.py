"""API de estatus de unidades: cruce Samsara + Logitrack del reporte configurado.

Ejecutar (produccion, ver servidor_api.py):
    pythonw.exe servidor_api.py

Consultar:
    GET /telemetria/consultarEstatus?centro_operativo=EC-05
    Encabezado: X-API-Key: <TELEMETRIA_API_KEY>
"""
from __future__ import annotations

from contextlib import asynccontextmanager
from datetime import datetime
import logging
import os
import secrets
import threading
import time
from typing import Any

from fastapi import Depends, FastAPI, Header, HTTPException, Query, Request
import pytz

import EnvioMain as envio

CACHE_SEGUNDOS_DEFAULT = 300
REFRESCO_SEGUNDOS_DEFAULT = 180
CLAVES_RESUMEN = {
    "DETENIDO": "detenido",
    "REVISAR": "revisar",
    "TRAFICO LENTO": "trafico_lento",
    "RETEN": "reten",
    "RUTA": "ruta",
}

MAX_INTENTOS_FALLIDOS_DEFAULT = 10
VENTANA_FALLOS_SEGUNDOS_DEFAULT = 600
BLOQUEO_SEGUNDOS_DEFAULT = 900

# /docs y /openapi.json describen la API a cualquiera; en produccion quedan apagados.
_DOCS = os.getenv("TELEMETRIA_DOCS", "false").strip().lower() in {"1", "true", "si", "sí"}
app = FastAPI(
    title="Telemetría de unidades",
    version="1.0.0",
    docs_url="/docs" if _DOCS else None,
    redoc_url=None,
    openapi_url="/openapi.json" if _DOCS else None,
)
_fallos: dict[str, list[float]] = {}
_bloqueados: dict[str, float] = {}
_candado_fallos = threading.Lock()


@app.middleware("http")
async def encabezados_seguridad(request: Request, call_next):
    respuesta = await call_next(request)
    # Ubicaciones de la flota: que ningun proxy ni navegador las guarde.
    respuesta.headers["Cache-Control"] = "no-store"
    respuesta.headers["X-Content-Type-Options"] = "nosniff"
    return respuesta
_cache: dict[str, tuple[float, dict[str, Any]]] = {}
_candados: dict[str, threading.Lock] = {}
_candado_global = threading.Lock()


def claves_configuradas() -> dict[str, str]:
    """Consumidor -> clave.

    TELEMETRIA_API_KEYS="erp:clave1,portal:clave2" da una clave por consumidor para poder
    revocarlas por separado; TELEMETRIA_API_KEY sigue aceptandose como consumidor "general".
    """
    claves: dict[str, str] = {}
    for par in os.getenv("TELEMETRIA_API_KEYS", "").split(","):
        nombre, separador, clave = par.strip().partition(":")
        if separador and nombre.strip() and clave.strip():
            claves[nombre.strip()] = clave.strip()
    unica = os.getenv("TELEMETRIA_API_KEY", "").strip()
    if unica:
        claves.setdefault("general", unica)
    return claves


def _ip_cliente(request: Request) -> str:
    # Detras de IIS, uvicorn ya resuelve la IP real desde X-Forwarded-For (servidor_api.py).
    return request.client.host if request.client else "desconocida"


def _registrar_fallo(ip: str, ahora: float) -> None:
    maximo = int(os.getenv("TELEMETRIA_MAX_INTENTOS_FALLIDOS", MAX_INTENTOS_FALLIDOS_DEFAULT))
    ventana = int(os.getenv("TELEMETRIA_VENTANA_FALLOS_SEGUNDOS", VENTANA_FALLOS_SEGUNDOS_DEFAULT))
    bloqueo = int(os.getenv("TELEMETRIA_BLOQUEO_SEGUNDOS", BLOQUEO_SEGUNDOS_DEFAULT))
    with _candado_fallos:
        recientes = [t for t in _fallos.get(ip, []) if ahora - t < ventana] + [ahora]
        _fallos[ip] = recientes
        if len(recientes) >= maximo:
            _bloqueados[ip] = ahora + bloqueo
            _fallos.pop(ip, None)
            logging.warning("IP %s bloqueada %s s por %s claves invalidas", ip, bloqueo, maximo)


def validar_api_key(
    request: Request, x_api_key: str | None = Header(default=None)
) -> str:
    """Regresa el nombre del consumidor; bloquea temporalmente IPs que prueban claves."""
    ip, ahora = _ip_cliente(request), time.monotonic()
    with _candado_fallos:
        hasta = _bloqueados.get(ip)
        if hasta and hasta <= ahora:
            _bloqueados.pop(ip, None)
            hasta = None
    if hasta:
        raise HTTPException(
            429, "Demasiados intentos fallidos; intente más tarde",
            headers={"Retry-After": str(int(hasta - ahora) + 1)},
        )
    claves = claves_configuradas()
    if not claves:
        raise HTTPException(503, "La API no tiene claves configuradas")
    consumidor = None
    # Se comparan todas las claves para que el tiempo de respuesta no revele cual coincide.
    for nombre, clave in claves.items():
        if secrets.compare_digest((x_api_key or "").encode(), clave.encode()):
            consumidor = nombre
    if consumidor is None:
        logging.warning("Clave invalida desde %s", ip)
        _registrar_fallo(ip, ahora)
        raise HTTPException(401, "API key inválida")
    return consumidor


def buscar_reporte(config: dict[str, Any], centro_operativo: str) -> dict[str, Any] | None:
    """Ubica el reporte por perfil (EC-05) o por nombre (Reporte EC-05)."""
    buscado = envio.normalizar_texto(centro_operativo)
    for reporte in config["reportes"]:
        nombres = {
            envio.normalizar_texto(reporte.get("perfil_filtros")),
            envio.normalizar_texto(reporte["nombre"]),
            envio.normalizar_texto(reporte["nombre"]).removeprefix("reporte "),
        }
        if buscado in nombres:
            return reporte
    return None


def _a_float(valor: Any) -> float | None:
    return envio.convertir_float(valor)


def _minutos(valor: Any) -> int | None:
    numero = _a_float(valor)
    return round(numero) if numero is not None else None


def unidad_a_json(row: dict[str, Any]) -> dict[str, Any]:
    estatus = envio.estatus_chat(row.get("Estatus"))
    if estatus == "DETENIDO":
        tiempo = _minutos(row.get("Minutos Detenido"))
    elif estatus == "TRAFICO LENTO":
        tiempo = _minutos(row.get("Minutos Trafico"))
    else:
        tiempo = None
    latitud, longitud = _a_float(row.get("Latitud")), _a_float(row.get("Longitud"))
    return {
        "unidad": str(row.get("Unidad") or ""),
        "estatus": estatus,
        "tiempo_minutos": tiempo,
        "coordenadas": (
            {"latitud": latitud, "longitud": longitud}
            if latitud is not None and longitud is not None else None
        ),
        "ubicacion": str(row.get("Ubicación") or ""),
    }


def construir_respuesta_estatus(
    centro_operativo: str, results: list[dict[str, Any]], now_mx: datetime
) -> dict[str, Any]:
    unidades = [unidad_a_json(row) for row in results]
    resumen = {clave: 0 for clave in CLAVES_RESUMEN.values()}
    for unidad in unidades:
        clave = CLAVES_RESUMEN.get(unidad["estatus"])
        if clave:
            resumen[clave] += 1
    return {
        "centro_operativo": centro_operativo,
        "fecha": now_mx.strftime("%Y-%m-%d"),
        "hora": now_mx.strftime("%H:%M:%S"),
        "total_unidades": len(unidades),
        "resumen": resumen,
        "unidades": unidades,
    }


def calcular_estatus(centro_operativo: str) -> dict[str, Any]:
    config = envio.cargar_configuracion(envio.DEFAULT_CONFIG_PATH)
    reporte = buscar_reporte(config, centro_operativo)
    if reporte is None:
        raise HTTPException(404, f"Centro operativo no configurado: {centro_operativo}")
    zona = pytz.timezone(config.get("zona_horaria", envio.DEFAULT_TIMEZONE))
    now_mx = datetime.now(zona)
    contexto = envio.preparar_contexto_samsara(config)
    # Google Sheets solo aporta columnas del Excel; no se consulta para la API.
    results, _, _ = envio.generar_resultados_reporte(
        contexto, reporte, now_mx, enriquecer_google_sheets=False
    )
    nombre = reporte.get("perfil_filtros") or centro_operativo.upper()
    return construir_respuesta_estatus(nombre, results, now_mx)


def _candado(clave: str) -> threading.Lock:
    with _candado_global:
        return _candados.setdefault(clave, threading.Lock())


def recalcular(centro_operativo: str) -> dict[str, Any]:
    """Calcula y guarda en cache; lo usan el refresco en segundo plano y las peticiones."""
    clave = envio.normalizar_texto(centro_operativo)
    with _candado(clave):
        respuesta = calcular_estatus(centro_operativo)
        _cache[clave] = (time.monotonic(), respuesta)
        return respuesta


def obtener_estatus_con_cache(centro_operativo: str) -> dict[str, Any]:
    """Reutiliza el ultimo calculo para no saturar Samsara ni Logitrack.

    Con el refresco en segundo plano activo, el cache siempre esta vigente y la
    respuesta es inmediata. Si el refresco falla y el cache envejece mas de
    TELEMETRIA_CACHE_SEGUNDOS, se recalcula en la peticion para no servir datos viejos.
    """
    clave = envio.normalizar_texto(centro_operativo)
    ttl = int(os.getenv("TELEMETRIA_CACHE_SEGUNDOS", CACHE_SEGUNDOS_DEFAULT))
    # Un solo calculo a la vez por centro; las demas peticiones esperan y usan el cache.
    with _candado(clave):
        guardado = _cache.get(clave)
        if guardado and time.monotonic() - guardado[0] < ttl:
            return guardado[1]
        respuesta = calcular_estatus(centro_operativo)
        _cache[clave] = (time.monotonic(), respuesta)
        return respuesta


def refrescar_periodicamente(centros: list[str], intervalo: int, detener: threading.Event) -> None:
    while not detener.is_set():
        for centro in centros:
            try:
                recalcular(centro)
                logging.info("Estatus de %s recalculado", centro)
            except Exception:
                logging.exception("Fallo el refresco en segundo plano de %s", centro)
        detener.wait(intervalo)


def centros_a_refrescar() -> list[str]:
    return [x.strip() for x in os.getenv("TELEMETRIA_CENTROS", "EC-05").split(",") if x.strip()]


@asynccontextmanager
async def ciclo_de_vida(_app: FastAPI):
    intervalo = int(os.getenv("TELEMETRIA_REFRESCO_SEGUNDOS", REFRESCO_SEGUNDOS_DEFAULT))
    detener = threading.Event()
    if intervalo > 0:
        threading.Thread(
            target=refrescar_periodicamente,
            args=(centros_a_refrescar(), intervalo, detener),
            name="refresco-telemetria",
            daemon=True,
        ).start()
    yield
    detener.set()


app.router.lifespan_context = ciclo_de_vida


@app.get("/salud")
def salud() -> dict[str, Any]:
    edades = {
        clave: round(time.monotonic() - guardado[0])
        for clave, guardado in _cache.items()
    }
    return {"estado": "ok", "segundos_desde_ultimo_calculo": edades}


@app.get("/telemetria/consultarEstatus")
@app.get("/telemetria/consultarEstatus/", include_in_schema=False)
def consultar_estatus(
    request: Request,
    centro_operativo: str = Query(
        "EC-05", max_length=40, description="Centro operativo, por ejemplo EC-05"
    ),
    consumidor: str = Depends(validar_api_key),
) -> dict[str, Any]:
    logging.info(
        "Consulta de estatus %s por %s desde %s", centro_operativo, consumidor, _ip_cliente(request)
    )
    try:
        return obtener_estatus_con_cache(centro_operativo)
    except HTTPException:
        raise
    except envio.ConfiguracionError as error:
        # El detalle de configuracion queda en el log, no en la respuesta.
        logging.error("Configuracion invalida para %s: %s", centro_operativo, error)
        raise HTTPException(500, "Error de configuración del servidor") from error
    except Exception as error:
        logging.exception("Fallo la consulta de estatus de %s", centro_operativo)
        raise HTTPException(502, "No se pudo consultar la telemetría") from error
