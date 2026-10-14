# API de estatus de unidades

`api_telemetria.py` expone la misma lógica del Reporte EC-05 (Samsara + Logitrack, geocercas,
historial de detención) como JSON. `servidor_api.py` la arranca en el servidor.

## Endpoint

```
GET /telemetria/consultarEstatus?centro_operativo=EC-05
X-API-Key: <TELEMETRIA_API_KEY>
```

`centro_operativo` acepta el perfil (`EC-05`) o el nombre del reporte (`Reporte EC-05`) definidos
en `config/reportes.json`. Respuesta:

```json
{
  "centro_operativo": "EC-05",
  "fecha": "2026-10-05",
  "hora": "11:30:07",
  "total_unidades": 35,
  "resumen": {"detenido": 14, "revisar": 0, "trafico_lento": 0, "reten": 0, "ruta": 21},
  "unidades": [
    {
      "unidad": "1939",
      "estatus": "DETENIDO",
      "tiempo_minutos": 1435,
      "coordenadas": {"latitud": 19.631638, "longitud": -99.17491},
      "ubicacion": "Avenida Independencia, Tultitlán, MEX, 54915"
    }
  ]
}
```

- `fecha` y `hora` son el momento del cálculo (hora de México).
- `estatus` usa los mismos valores que Google Chat: `DETENIDO`, `REVISAR`, `TRAFICO LENTO`,
  `RETEN`, `RUTA`. Las unidades vienen ordenadas con las detenidas primero.
- `tiempo_minutos` es el tiempo detenido (o en tráfico lento); `null` para las demás.
- Las unidades omitidas por patio, geocerca o GPS viejo no aparecen, igual que en el reporte.

| Código | Significado |
|---|---|
| 200 | Consulta correcta |
| 401 | Falta `X-API-Key` o no coincide |
| 404 | Centro operativo no configurado |
| 502 | Samsara no respondió (Logitrack caído no es error: se usa solo Samsara) |
| 503 | El servidor no tiene `TELEMETRIA_API_KEY` |

`GET /salud` responde sin clave, con los segundos desde el último cálculo de cada centro. La
documentación interactiva queda en `/docs`.

## Rendimiento

Un cálculo completo tarda ~15–20 s (Samsara, Logitrack e historial GPS). Por eso la API recalcula
en segundo plano cada `TELEMETRIA_REFRESCO_SEGUNDOS` (180) los centros de `TELEMETRIA_CENTROS`, y
las consultas responden al instante con el último cálculo.

Si el refresco falla (por ejemplo, Samsara no responde) y el último cálculo tiene más de
`TELEMETRIA_CACHE_SEGUNDOS` (300), la siguiente consulta recalcula en ese momento; si tampoco puede,
responde 502. Nunca se sirve un estatus más viejo que ese límite.

## Configuración

En `.env`:

```
TELEMETRIA_API_KEY=<cadena larga y aleatoria>
TELEMETRIA_HOST=0.0.0.0
TELEMETRIA_PUERTO=8000
TELEMETRIA_CENTROS=EC-05
TELEMETRIA_REFRESCO_SEGUNDOS=180
TELEMETRIA_CACHE_SEGUNDOS=300
```

Genere la clave con:

```powershell
.\.venv\Scripts\python.exe -c "import secrets; print(secrets.token_urlsafe(32))"
```

## Ejecutar en local

```powershell
.\.venv\Scripts\python.exe .\servidor_api.py
```

Consulte `http://127.0.0.1:8000/docs` (botón **Authorize** con la clave) o:

```powershell
Invoke-RestMethod "http://127.0.0.1:8000/telemetria/consultarEstatus?centro_operativo=EC-05" `
  -Headers @{"X-API-Key"="<clave>"} | ConvertTo-Json -Depth 5
```

## Producción en el servidor (Programador de tareas)

1. Actualizar código y dependencias:

   ```powershell
   cd C:\Users\Administrator\source\repos\Unidades-1
   git pull
   & "C:\Users\Administrator\AppData\Local\Programs\Python\Python313\python.exe" -m pip install -r requirements.txt
   ```

2. Agregar al `.env` del servidor las variables `TELEMETRIA_*` de arriba.
3. Probar a mano con `python.exe .\servidor_api.py` y consultar desde el mismo servidor.
4. Crear la tarea **API Telemetria**:
   - Desencadenador: **Al iniciar el sistema**.
   - Acción: programa
     `C:\Users\Administrator\AppData\Local\Programs\Python\Python313\pythonw.exe`, argumentos
     `servidor_api.py`, iniciar en `C:\Users\Administrator\source\repos\Unidades-1`.
   - General: **Ejecutar tanto si el usuario inició sesión como si no**.
   - Configuración: desmarcar **Detener la tarea si se ejecuta durante más de…**; **Si la tarea
     no se puede ejecutar, reiniciar cada 1 minuto**, hasta 3 veces; **No iniciar una instancia
     nueva** si ya se está ejecutando.
5. Abrir el puerto solo a la red que consumirá la API (ajuste la subred):

   ```powershell
   New-NetFirewallRule -DisplayName "API Telemetria" -Direction Inbound -Protocol TCP `
     -LocalPort 8000 -RemoteAddress 192.168.0.0/16 -Action Allow
   ```

El log queda en `api_telemetria.log` (rota a 5 MB y conserva 5 archivos). Después de un
`git pull`, reinicie la tarea para cargar el código nuevo.

La API viaja por HTTP. Dentro de la red interna es aceptable; si se expone fuera, publíquela detrás
de HTTPS (IIS como proxy inverso) para que la clave no viaje en claro.
