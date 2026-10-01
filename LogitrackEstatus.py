"""Consulta el ultimo estatus de las unidades en la API de Logitrack."""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import requests
from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_API_URL = "https://api.logitrack.mx/api/v2"
DEFAULT_TOKEN_URL = "https://api.logitrack.mx/api/v2/oauth/access_token"
DEFAULT_TIMEOUT_SECONDS = 60


def required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Falta configurar {name} en el archivo .env")
    return value


def request_access_token(session: requests.Session, token_url: str, timeout: int) -> str:
    """Solicita un access token usando las credenciales configuradas."""
    configured_token = os.getenv("LOGITRACK_ACCESS_TOKEN", "").strip()
    if configured_token:
        return configured_token

    payload = {
        "username": required_env("LOGITRACK_USERNAME"),
        "password": required_env("LOGITRACK_PASSWORD"),
        "grant_type": "password",
        "client_id": required_env("LOGITRACK_CLIENT_ID"),
        "client_secret": required_env("LOGITRACK_CLIENT_SECRET"),
    }
    response = session.post(
        token_url,
        headers={
            "Accept": "application/json",
            "Content-Type": "application/x-www-form-urlencoded",
        },
        data=payload,
        timeout=timeout,
    )
    if response.status_code == 404:
        raise RuntimeError(
            "El endpoint para generar el token devolvio 404: "
            f"{token_url}. Verifique LOGITRACK_TOKEN_URL."
        )
    response.raise_for_status()
    token = response.json().get("access_token")
    if not token:
        raise RuntimeError("Logitrack no devolvio access_token en la autenticacion")
    return str(token)


def fetch_last_status(
    session: requests.Session, api_url: str, token: str, timeout: int
) -> list[dict[str, Any]]:
    """Obtiene el ultimo estatus reportado por cada unidad."""
    response = session.get(
        f"{api_url}/units/last_status",
        headers={
            "Accept": "application/json",
            "Content-Type": "application/json",
            "Authorization": f"Bearer {token}",
        },
        timeout=timeout,
    )
    response.raise_for_status()
    data = response.json()
    if not isinstance(data, list):
        raise RuntimeError("La respuesta de last_status no es una lista de unidades")
    return data


def save_statuses(statuses: list[dict[str, Any]], output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    suffix = output.suffix.lower()

    if suffix == ".json":
        output.write_text(
            json.dumps(statuses, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return

    frame = pd.DataFrame(statuses)
    if suffix == ".csv":
        frame.to_csv(output, index=False, encoding="utf-8-sig")
        return
    if suffix == ".xlsx":
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            frame.to_excel(writer, sheet_name="estatus", index=False)
            worksheet = writer.sheets["estatus"]
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, max(len(frame), 1), max(len(frame.columns) - 1, 0))
        return
    raise ValueError("El archivo de salida debe terminar en .xlsx, .csv o .json")


def parse_args() -> argparse.Namespace:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    parser = argparse.ArgumentParser(
        description="Consulta el ultimo estatus de las unidades en Logitrack."
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=BASE_DIR / "outputs" / f"estatus_logitrack_{timestamp}.xlsx",
        help="Ruta de salida (.xlsx, .csv o .json).",
    )
    return parser.parse_args()


def main() -> int:
    load_dotenv(BASE_DIR / ".env")
    args = parse_args()
    api_url = os.getenv("LOGITRACK_API_URL", DEFAULT_API_URL).strip().rstrip("/")
    token_url = os.getenv("LOGITRACK_TOKEN_URL", DEFAULT_TOKEN_URL).strip()
    timeout = int(os.getenv("LOGITRACK_TIMEOUT_SECONDS", DEFAULT_TIMEOUT_SECONDS))

    try:
        with requests.Session() as session:
            token = request_access_token(session, token_url, timeout)
            statuses = fetch_last_status(session, api_url, token, timeout)
        save_statuses(statuses, args.output)
    except (requests.RequestException, RuntimeError, ValueError, OSError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1

    print(f"Listo: {len(statuses)} unidades -> {args.output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
