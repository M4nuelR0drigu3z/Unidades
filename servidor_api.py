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

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent


def configurar_log() -> None:
    manejador = RotatingFileHandler(
        BASE_DIR / "api_telemetria.log", maxBytes=5_000_000, backupCount=5, encoding="utf-8"
    )
    manejador.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s"))
    raiz = logging.getLogger()
    raiz.setLevel(logging.INFO)
    raiz.addHandler(manejador)


def main() -> None:
    load_dotenv(BASE_DIR / ".env")
    # Antes de importar la API: EnvioMain usa basicConfig, que no hace nada si ya hay manejadores.
    configurar_log()
    import uvicorn

    from api_telemetria import app

    if not os.getenv("TELEMETRIA_API_KEY"):
        logging.error("Falta TELEMETRIA_API_KEY en .env; la API respondera 503")
    uvicorn.run(
        app,
        host=os.getenv("TELEMETRIA_HOST", "0.0.0.0"),
        port=int(os.getenv("TELEMETRIA_PUERTO", "8000")),
        log_config=None,
    )


if __name__ == "__main__":
    main()
