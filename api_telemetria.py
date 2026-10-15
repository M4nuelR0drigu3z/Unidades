"""API de estatus de unidades: cruce Samsara + Logitrack del reporte configurado.

Ejecutar (produccion, ver servidor_api.py):
    pythonw.exe servidor_api.py

Consultar:
    GET /telemetria/consultarEstatus?centro_operativo=EC-05
    Encabezado: X-API-Key: <TELEMETRIA_API_KEY>
"""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from contextlib import asynccontextmanager
from datetime import datetime
import logging
import os
import re
import secrets
import threading
import time
from typing import Any

from fastapi import Depends, FastAPI, Header, HTTPException, Query, Request
import pytz

import catalogo_unidades
import EnvioMain as envio

CACHE_SEGUNDOS_DEFAULT = 300
REFRESCO_SEGUNDOS_DEFAULT = 180
HILOS_REFRESCO_DEFAULT = 4
CLAVES_RESUMEN = {"DETENIDO": "detenido", "RUTA": "ruta"}
# El cliente solo recibe DETENIDO o RUTA. REVISAR (detenida hace menos del minimo o con
# ubicaciones que no coinciden) y RETEN cuentan como detenida; TRAFICO LENTO avanza, es RUTA.
# El detalle sigue disponible en el Excel y el log del reporte.
ESTATUS_DETENIDO_API = {"DETENIDO", "REVISAR", "RETEN"}


def estatus_api(estatus: Any) -> str:
    return "DETENIDO" if envio.estatus_chat(estatus) in ESTATUS_DETENIDO_API else "RUTA"


ORDEN_ESTADOS = [
    "DETENIDO CONFIRMADO", "DETENIDO SAMSARA", "DETENIDO LOGITRACK", "DETENIDO",
    "REVISAR", "TRAFICO LENTO", "RETEN", "RUTA",
]
# Solo nombres de centro operativo; evita consultar en Samsara etiquetas arbitrarias.
PATRON_CENTRO = re.compile(r"EC-\d{2}", re.IGNORECASE)
MAX_CENTROS_POR_CONSULTA = 10
CATALOGO_CENTROS_SEGUNDOS = 1800
_catalogo_centros: dict[str, Any] = {"momento": None, "centros": []}
_candado_catalogo = threading.Lock()

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
    """Reporte con configuracion propia para la API: el que cruza con Logitrack (hoy EC-05).

    Los reportes pensados solo para Excel (Sayer, EC-02 con todas las unidades) no aplican:
    esos centros usan las reglas generales de reporte_generico.
    """
    buscado = envio.normalizar_texto(centro_operativo)
    for reporte in config["reportes"]:
        nombres = {
            envio.normalizar_texto(reporte.get("perfil_filtros")),
            envio.normalizar_texto(reporte["nombre"]),
            envio.normalizar_texto(reporte["nombre"]).removeprefix("reporte "),
        }
        if buscado in nombres and (reporte.get("telemetria_logitrack") or {}).get("activo"):
            return reporte
    return None


def reporte_generico(centro: str) -> dict[str, Any]:
    """Reglas para centros sin configuracion propia: solo Samsara y su historial.

    Igual que EC-05, excluye las geocercas que en Samsara tienen la etiqueta del centro
    (si no tiene ninguna, no excluye nada por ese lado), ademas de las geocercas especiales
    de filtros_base (bases TDR, CEDIS); y omite las unidades detenidas a 400 m de ellas.
    """
    return {
        "nombre": f"Reporte {centro}",
        "etiquetas": [centro],
        "tipo_filtro_etiqueta": "parentTagIds",
        "analizar_detenciones": True,
        "enriquecer_google_sheets": False,
        "velocidad_detenido_mph": 3,
        "filtros": {
            "excluir_speed_cero_sin_ecu": False,
            "excluir_unidades_en_geocerca": True,
            "etiquetas_geocercas_excluidas": [centro],
            "distancia_cercania_geocerca_metros": 400,
        },
        "contenido": {"orden_estados": ORDEN_ESTADOS},
    }


def centros_samsara(config: dict[str, Any]) -> list[str]:
    """Centros EC-xx que existen como etiqueta en Samsara; se vuelve a consultar cada 30 min."""
    with _candado_catalogo:
        momento = _catalogo_centros["momento"]
        if momento is not None and time.monotonic() - momento < CATALOGO_CENTROS_SEGUNDOS:
            return _catalogo_centros["centros"]
        sesion = envio.preparar_contexto_samsara(config)["sesion"]
        etiquetas = envio.obtener_catalogo_etiquetas_samsara(sesion, 512)
        centros = sorted({
            str(etiqueta.get("nombre")).strip().upper() for etiqueta in etiquetas
            if PATRON_CENTRO.fullmatch(str(etiqueta.get("nombre") or "").strip())
        })
        _catalogo_centros.update(momento=time.monotonic(), centros=centros)
        return centros


def resolver_reporte(
    config: dict[str, Any], centro_operativo: str
) -> tuple[dict[str, Any] | None, str]:
    """(reporte, nombre del centro); reporte es None si el centro no existe."""
    reporte = buscar_reporte(config, centro_operativo)
    if reporte is not None:
        return reporte, reporte.get("perfil_filtros") or centro_operativo.strip().upper()
    nombre = centro_operativo.strip().upper()
    if PATRON_CENTRO.fullmatch(nombre) and nombre in centros_samsara(config):
        return reporte_generico(nombre), nombre
    return None, nombre


def separar_centros(texto: str, config: dict[str, Any]) -> list[str]:
    """'EC-05, ec-10' -> ['EC-05', 'EC-10']; 'TODOS' -> todos los centros de Samsara."""
    nombres = list(dict.fromkeys(x.strip().upper() for x in texto.split(",") if x.strip()))
    if "TODOS" in nombres:
        return centros_samsara(config)
    if not nombres:
        raise HTTPException(422, "Indique al menos un centro operativo")
    if len(nombres) > MAX_CENTROS_POR_CONSULTA:
        raise HTTPException(422, f"Máximo {MAX_CENTROS_POR_CONSULTA} centros por consulta")
    desconocidos = [nombre for nombre in nombres if resolver_reporte(config, nombre)[0] is None]
    if desconocidos:
        raise HTTPException(404, f"Centro operativo no disponible: {', '.join(desconocidos)}")
    return nombres


def _a_float(valor: Any) -> float | None:
    return envio.convertir_float(valor)


def _minutos(valor: Any) -> int | None:
    numero = _a_float(valor)
    return round(numero) if numero is not None else None


def fecha_gps(row: dict[str, Any]) -> str | None:
    """Hora de la lectura de donde salen las coordenadas (hora de Mexico).

    Si Samsara no tenia GPS vigente y la unidad se rescato con Logitrack, las coordenadas
    son de Logitrack y su fecha tambien.
    """
    if row.get("Estatus Samsara Actual") in envio.MOTIVOS_SAMSARA_SIN_GPS_VIGENTE:
        fecha = row.get("Fecha Logitrack")
    else:
        fecha = row.get("Fecha GPS")
    return fecha.strftime("%Y-%m-%d %H:%M:%S") if isinstance(fecha, datetime) else None


def unidad_a_json(row: dict[str, Any]) -> dict[str, Any]:
    estatus = estatus_api(row.get("Estatus"))
    tiempo = _minutos(row.get("Minutos Detenido")) if estatus == "DETENIDO" else None
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
        "fecha_gps": fecha_gps(row),
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
    reporte, nombre = resolver_reporte(config, centro_operativo)
    if reporte is None:
        raise HTTPException(404, f"Centro operativo no disponible: {nombre}")
    zona = pytz.timezone(config.get("zona_horaria", envio.DEFAULT_TIMEZONE))
    now_mx = datetime.now(zona)
    contexto = envio.preparar_contexto_samsara(config)
    # Google Sheets solo aporta columnas del Excel; no se consulta para la API.
    results, _, _ = envio.generar_resultados_reporte(
        contexto, reporte, now_mx, enriquecer_google_sheets=False
    )
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


def expandir_centros(centros: list[str]) -> list[str]:
    if any(centro.strip().upper() == "TODOS" for centro in centros):
        return centros_samsara(envio.cargar_configuracion(envio.DEFAULT_CONFIG_PATH))
    return centros


def segundos_para_siguiente_marca(intervalo: int, ahora: float | None = None) -> float:
    """Segundos hasta el siguiente multiplo de `intervalo` en el reloj.

    Con 180 s las marcas son 12:00:00, 12:03:00 ... 12:30:00. Los husos de Mexico difieren
    de UTC en horas enteras, asi que los multiplos sobre la hora epoch caen igual en hora local.
    """
    ahora = time.time() if ahora is None else ahora
    return intervalo - (ahora % intervalo)


def _recalcular_en_segundo_plano(centro: str) -> None:
    try:
        recalcular(centro)
        logging.info("Estatus de %s recalculado", centro)
    except Exception:
        logging.exception("Fallo el refresco en segundo plano de %s", centro)


def refrescar_centros(centros: list[str], hilos: int) -> None:
    try:
        lista = expandir_centros(centros)
    except Exception:
        logging.exception("No se pudo obtener la lista de centros de Samsara")
        lista = [centro for centro in centros if centro.strip().upper() != "TODOS"]
    if not lista:
        return
    # Varios centros a la vez: el ciclo termina pocos segundos despues de la marca y todos
    # los centros quedan con la misma hora de corte.
    with ThreadPoolExecutor(max_workers=max(1, min(hilos, len(lista)))) as ejecutor:
        list(ejecutor.map(_recalcular_en_segundo_plano, lista))


def refrescar_periodicamente(
    centros: list[str], intervalo: int, detener: threading.Event,
    hilos: int = HILOS_REFRESCO_DEFAULT,
) -> None:
    """Primer calculo al arrancar; despues, cada ciclo empieza en una marca del reloj.

    Si un ciclo tarda mas que el intervalo, el siguiente espera a la marca que sigue.
    """
    while not detener.is_set():
        refrescar_centros(centros, hilos)
        detener.wait(segundos_para_siguiente_marca(intervalo))


def centros_a_refrescar() -> list[str]:
    return [x.strip() for x in os.getenv("TELEMETRIA_CENTROS", "TODOS").split(",") if x.strip()]


@asynccontextmanager
async def ciclo_de_vida(_app: FastAPI):
    intervalo = int(os.getenv("TELEMETRIA_REFRESCO_SEGUNDOS", REFRESCO_SEGUNDOS_DEFAULT))
    hilos = int(os.getenv("TELEMETRIA_HILOS_REFRESCO", HILOS_REFRESCO_DEFAULT))
    detener = threading.Event()
    if intervalo > 0:
        if 1800 % intervalo:
            logging.warning(
                "TELEMETRIA_REFRESCO_SEGUNDOS=%s no divide 30 minutos: los calculos no "
                "coincidiran con las :00 y :30", intervalo,
            )
        threading.Thread(
            target=refrescar_periodicamente,
            args=(centros_a_refrescar(), intervalo, detener, hilos),
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
        "EC-05", max_length=120,
        description="Un centro (EC-05), varios separados por coma (EC-05,EC-10) o TODOS",
    ),
    consumidor: str = Depends(validar_api_key),
) -> Any:
    logging.info(
        "Consulta de estatus %s por %s desde %s", centro_operativo, consumidor, _ip_cliente(request)
    )
    try:
        config = envio.cargar_configuracion(envio.DEFAULT_CONFIG_PATH)
        centros = separar_centros(centro_operativo, config)
        respuestas = [obtener_estatus_con_cache(centro) for centro in centros]
        # Un centro: el objeto tal cual (compatibilidad); varios o TODOS: una lista.
        es_lista = len(centros) > 1 or centro_operativo.strip().upper() == "TODOS"
        return respuestas if es_lista else respuestas[0]
    except HTTPException:
        raise
    except envio.ConfiguracionError as error:
        # El detalle de configuracion queda en el log, no en la respuesta.
        logging.error("Configuracion invalida para %s: %s", centro_operativo, error)
        raise HTTPException(500, "Error de configuración del servidor") from error
    except Exception as error:
        logging.exception("Fallo la consulta de estatus de %s", centro_operativo)
        raise HTTPException(502, "No se pudo consultar la telemetría") from error


@app.get("/telemetria/consultarUnidades")
@app.get("/telemetria/consultarUnidades/", include_in_schema=False)
def consultar_unidades(
    request: Request,
    unidad: str | None = Query(None, max_length=200, description="Una o varias separadas por coma"),
    division: str | None = Query(None, max_length=120, description="EC-05, o varias separadas por coma"),
    flota: int | None = Query(None, ge=0, description="ID de flota"),
    incluir_bajas: bool = Query(False, description="Incluye unidades dadas de baja (OUT)"),
    consumidor: str = Depends(validar_api_key),
) -> dict[str, Any]:
    """Catalogo de unidades de TMWSuite: numero, division, flota y placas."""
    logging.info(
        "Consulta de unidades unidad=%s division=%s flota=%s por %s desde %s",
        unidad, division, flota, consumidor, _ip_cliente(request),
    )
    try:
        unidades = catalogo_unidades.unidades_con_cache()
    except Exception as error:
        logging.exception("Fallo la consulta del catalogo de unidades")
        raise HTTPException(502, "No se pudo consultar el catálogo de unidades") from error
    filtradas = catalogo_unidades.filtrar_unidades(unidades, unidad, division, flota, incluir_bajas)
    return {"total_unidades": len(filtradas), "unidades": filtradas}
