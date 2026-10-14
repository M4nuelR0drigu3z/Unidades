"""Arranca la API de telemetria para produccion local (Programador de tareas con pythonw).

pythonw no tiene consola, asi que todo el log (peticiones, refresco y errores) va a
api_telemetria.log, con rotacion para que no crezca sin limite.

Variables de entorno (.env): TELEMETRIA_HOST (0.0.0.0), TELEMETRIA_PUERTO (8000),
TELEMETRIA_API_KEY, TELEMETRIA_REFRESCO_SEGUNDOS, TELEMETRIA_CACHE_SEGUNDOS, TELEMETRIA_CENTROS.
"""
from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path
import socket
import sys

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent


def configurar_log() -> None:
    formato = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
    manejador = RotatingFileHandler(
        BASE_DIR / "api_telemetria.log", maxBytes=5_000_000, backupCount=5, encoding="utf-8"
    )
    manejador.setFormatter(formato)
    raiz = logging.getLogger()
    raiz.setLevel(logging.INFO)
    raiz.addHandler(manejador)
    # Con python.exe (prueba manual) tambien se ve en consola; con pythonw no hay stdout.
    if sys.stdout is not None:
        consola = logging.StreamHandler(sys.stdout)
        consola.setFormatter(formato)
        raiz.addHandler(consola)


def puerto_ocupado(host: str, puerto: int) -> bool:
    """True si otra aplicacion ya atiende ese puerto.

    En Windows uvicorn puede abrir un puerto ya usado sin marcar error, y las peticiones
    siguen llegando a la otra aplicacion; por eso se revisa antes de arrancar.
    """
    destino = "127.0.0.1" if host in ("0.0.0.0", "") else host
    try:
        with socket.create_connection((destino, puerto), timeout=1):
            return True
    except OSError:
        return False


def main() -> None:
    load_dotenv(BASE_DIR / ".env")
    # Antes de importar la API: EnvioMain usa basicConfig, que no hace nada si ya hay manejadores.
    configurar_log()
    import uvicorn

    from api_telemetria import app

    # Por defecto solo escucha dentro del servidor: el acceso externo entra por IIS con HTTPS.
    host = os.getenv("TELEMETRIA_HOST", "127.0.0.1")
    puerto = int(os.getenv("TELEMETRIA_PUERTO", "8000"))
    if puerto_ocupado(host, puerto):
        logging.error(
            "El puerto %s ya lo usa otra aplicacion; cambie TELEMETRIA_PUERTO en .env", puerto
        )
        sys.exit(1)
    if not (os.getenv("TELEMETRIA_API_KEYS") or os.getenv("TELEMETRIA_API_KEY")):
        logging.error("Faltan TELEMETRIA_API_KEYS en .env; la API respondera 503")
    logging.info("API de telemetria en http://%s:%s", host, puerto)
    uvicorn.run(
        app,
        host=host,
        port=puerto,
        log_config=None,
        server_header=False,
        # Solo se confia en X-Forwarded-For cuando la peticion viene de IIS en el mismo equipo.
        proxy_headers=True,
        forwarded_allow_ips="127.0.0.1",
        limit_concurrency=int(os.getenv("TELEMETRIA_MAX_CONEXIONES", "50")),
    )


if __name__ == "__main__":
    main()
