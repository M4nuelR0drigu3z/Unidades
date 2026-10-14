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
| 429 | IP bloqueada temporalmente por claves inválidas |
| 404 | Centro operativo no configurado |
| 502 | Samsara no respondió (Logitrack caído no es error: se usa solo Samsara) |
| 503 | El servidor no tiene claves configuradas |

`GET /salud` responde sin clave, con los segundos desde el último cálculo de cada centro. La
documentación interactiva (`/docs`) está apagada salvo con `TELEMETRIA_DOCS=true`.

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
TELEMETRIA_API_KEYS=erp:<clave1>,portal:<clave2>
TELEMETRIA_HOST=127.0.0.1
TELEMETRIA_PUERTO=8010
TELEMETRIA_CENTROS=EC-05
TELEMETRIA_REFRESCO_SEGUNDOS=180
TELEMETRIA_CACHE_SEGUNDOS=300
```

- `TELEMETRIA_API_KEYS`: una clave por consumidor (`nombre:clave`, separados por coma). Para
  revocar a un consumidor se borra su par y se reinicia la tarea. El log registra qué consumidor
  hizo cada consulta. `TELEMETRIA_API_KEY` (una sola clave) sigue funcionando como consumidor
  `general`.
- `TELEMETRIA_HOST=127.0.0.1`: la API solo escucha dentro del servidor; el acceso externo entra por
  IIS con HTTPS.
- Opcionales: `TELEMETRIA_DOCS=true` enciende `/docs` (apagado por defecto);
  `TELEMETRIA_MAX_INTENTOS_FALLIDOS` (10), `TELEMETRIA_VENTANA_FALLOS_SEGUNDOS` (600) y
  `TELEMETRIA_BLOQUEO_SEGUNDOS` (900) controlan el bloqueo por claves inválidas;
  `TELEMETRIA_MAX_CONEXIONES` (50).

Genere cada clave con:

```powershell
.\.venv\Scripts\python.exe -c "import secrets; print(secrets.token_urlsafe(32))"
```

## Seguridad

| Capa | Protección |
|---|---|
| IIS | HTTPS con el certificado del servidor, HSTS, solo `GET`, solo se publica `/telemetria/*`, límite de 60 peticiones/minuto y 10 simultáneas por IP, sin encabezado `Server` |
| Firewall | El puerto 8010 no se abre; solo IIS (443) es accesible desde fuera |
| API | Clave por consumidor; 10 claves inválidas en 10 min bloquean la IP 15 min (429); `/docs` apagado; respuestas `Cache-Control: no-store`; errores sin detalles internos; solo lectura |
| Servidor | Ejecutar la tarea con una cuenta sin privilegios de administrador y restringir el `.env` a esa cuenta |

## Ejecutar en local

```powershell
$env:TELEMETRIA_DOCS = "true"
.\.venv\Scripts\python.exe .\servidor_api.py
```

Consulte `http://127.0.0.1:8010/docs` (botón **Authorize** con la clave) o:

```powershell
Invoke-RestMethod "http://127.0.0.1:8010/telemetria/consultarEstatus?centro_operativo=EC-05" `
  -Headers @{"X-API-Key"="<clave>"} | ConvertTo-Json -Depth 5
```

## Producción en el servidor

### 1. Código y tarea programada

1. Actualizar código y dependencias:

   ```powershell
   cd C:\Users\Administrator\source\repos\Unidades-1
   git pull
   & "C:\Users\Administrator\AppData\Local\Programs\Python\Python313\python.exe" -m pip install -r requirements.txt
   ```

2. Agregar al `.env` del servidor las variables `TELEMETRIA_*` de arriba.
3. Crear la tarea **API Telemetria**:
   - Desencadenador: **Al iniciar el sistema**.
   - Acción: programa
     `C:\Users\Administrator\AppData\Local\Programs\Python\Python313\pythonw.exe`, argumentos
     `servidor_api.py`, iniciar en `C:\Users\Administrator\source\repos\Unidades-1`.
   - General: **Ejecutar tanto si el usuario inició sesión como si no**.
   - Configuración: desmarcar **Detener la tarea si se ejecuta durante más de…**; **Si la tarea
     no se puede ejecutar, reiniciar cada 1 minuto**, hasta 3 veces; **No iniciar una instancia
     nueva** si ya se está ejecutando.

### 2. IIS con HTTPS

1. Instalar en el servidor (una sola vez):
   - **URL Rewrite** y **Application Request Routing (ARR)** de Microsoft.
   - El rol **Servidor web > Seguridad > Restricciones de IP y dominio** (Administrador del
     servidor > Agregar roles y características).
2. En el Administrador de IIS, seleccionar el **servidor** > **Application Request Routing Cache**
   > **Server Proxy Settings** > marcar **Enable proxy** y dejar **Preserve client IP in the
   following header: X-Forwarded-For** > Aplicar.
3. Publicar la API, en una de dos formas:
   - **Sitio propio (recomendado):** crear la carpeta `C:\inetpub\api-telemetria`, copiar
     `deploy\iis\web.config` ahí y crear un sitio con esa carpeta y un enlace **https** con el
     certificado existente (puerto 443 con un nombre de host propio si el certificado lo cubre, o
     un puerto como 8443).
   - **Dentro de un sitio HTTPS existente:** copiar las reglas de `deploy\iis\web.config` al
     `web.config` de ese sitio (solo los bloques `rewrite` y `security`). La API queda en
     `https://<sitio>/telemetria/consultarEstatus`.
4. Probar desde otro equipo:

   ```powershell
   Invoke-RestMethod "https://<host>/telemetria/consultarEstatus?centro_operativo=EC-05" `
     -Headers @{"X-API-Key"="<clave>"} | ConvertTo-Json -Depth 5
   ```

5. Si existía la regla de firewall del puerto 8010, eliminarla:

   ```powershell
   Remove-NetFirewallRule -DisplayName "API Telemetria"
   ```

El log queda en `api_telemetria.log` (rota a 5 MB y conserva 5 archivos). Después de un
`git pull`, reinicie la tarea para cargar el código nuevo. Revise en el log los avisos
`Clave invalida desde` y `bloqueada`: muchas apariciones indican que alguien está probando claves.
