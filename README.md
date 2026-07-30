# Driver Scoreboard

Dashboard en Streamlit para visualizar ranking de operadores usando datos de Samsara.

## Archivos principales

- `app.py`: aplicación principal de Streamlit.
- `refresh_cache.py`: consulta la API de Samsara y genera el cache local.

## Variables de entorno requeridas

Crear un archivo `.env` local con:

```env
SAMSARA_TOKEN=tu_token
BASE_URL=https://api.samsara.com
DAYS_BACK=7
MIN_KM_STORE=0
GLOBAL_RPS=4.5