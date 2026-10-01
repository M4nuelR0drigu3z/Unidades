# Configuración de reportes

`EnvioMain.py` resuelve las etiquetas de Samsara por nombre y ejecuta cada bloque activo de
`config/reportes.json` de manera independiente.

## Validar sin enviar

```powershell
.\.venv\Scripts\python.exe .\EnvioMain.py --dry-run
```

Google Chat se imprime en consola. Si la entrega de correo está activa, su Excel se guarda en
`outputs/envio_previews/`, sin conectarse al servidor SMTP.

```powershell
.\.venv\Scripts\python.exe .\EnvioMain.py --dry-run --solo "Reporte EC-05"
.\.venv\Scripts\python.exe .\EnvioMain.py --solo "Reporte EC-05" --canal correo --dry-run
.\.venv\Scripts\python.exe .\EnvioMain.py --listar-etiquetas
.\.venv\Scripts\python.exe .\EnvioMain.py --sincronizar-etiquetas
```

## Qué controla cada reporte

- `activo`: habilita o deshabilita el reporte.
- `etiquetas`: nombres visibles en Samsara. Sus IDs se resuelven en cada ejecución.
- `tipo_filtro_etiqueta`: `tagIds` para pertenencia directa o `parentTagIds` para descendientes.
- `perfil_filtros`: aplica una política reutilizable definida en `perfiles_filtros`.
- `filtros`: sobrescribe opciones de `filtros_base` únicamente para ese reporte.
- `contenido.estados`: vacío incluye todos; `["DETENIDO"]` incluye solo detenidos.
- `contenido.unidades`: vacío incluye todas o se puede limitar a unidades concretas.
- `contenido.columnas`: define las columnas del Excel.
- `entregas`: permite uno o varios destinos para el mismo reporte.

`--canal google_chat` o `--canal correo` fuerza ese canal para la ejecución, aunque la entrega
esté inactiva en el JSON. Esto permite usar dos tareas programadas sin cambiar la configuración.

Google Chat toma el webhook de la variable indicada en `webhook_env`. El correo acepta una lista
en `destinatarios` y/o una variable `destinatarios_env` con direcciones separadas por coma. El
Excel incluye `Resumen`, `Unidades` y opcionalmente `Omitidas`.

## Padres, hijos y geocercas

`--sincronizar-etiquetas` actualiza `config/catalogo_etiquetas_samsara.json` con todas las etiquetas,
su `parentTagId` y la jerarquía completa de hijos. El catálogo también funciona como caché local
para resolver nombres a IDs.

El perfil `EC-05` tiene `excluir_unidades_en_geocerca: true`: obtiene de Samsara las geocercas
asociadas a ese tag y omite las unidades ubicadas dentro. El perfil `EC-02` usa
`excluir_unidades_en_geocerca: false` y vacía las geocercas especiales, por lo que no elimina
unidades por ubicación.

EC-02 conserva todas las unidades que pertenezcan directamente a `Sayer Full`,
`Sayer Patios y T.`, `Sayer Pipas`, `Sayer Sencillo`, `Sayer Thorton` o `Sayer Vuelteros`.
Usa `tagIds`, por lo que no incorpora automáticamente otras etiquetas hijas de EC-02. Dentro de
ese conjunto aplica `incluir_todas_las_unidades: true`, `gps_max_minutos: null` y
`excluir_speed_cero_sin_ecu: false`. No analiza detenciones ni genera hojas de resumen u omitidas.
Toda la información queda concentrada en `Unidades`, con las columnas
`Unidad`, `Estatus`, `Operador`, `Ubicación`, `Coordenadas` y `Geocerca`.
Las filas se ordenan con `DETENIDO` primero y `RUTA` después. La fila superior de indicadores
muestra `Total`, `En ruta` y `Detenidos`.
El operador corresponde a la asignación no pasajera más reciente encontrada en Samsara durante
las últimas 24 horas; queda vacío cuando Samsara no registra una asignación en esa ventana.

`Reporte EC-05` es el único reporte activo por defecto y su entrega activa es Google Chat mediante
`GOOGLE_CHAT_WEBHOOK_URL`. El correo puede forzarse con `--canal correo` y usa `MAIL_TO`.
`Reporte Sayer` y `Reporte EC-02` quedan desactivados. Al usar `parentTagIds`, EC-05 incluye
automáticamente los vehículos del padre y de todos sus descendientes.

### Doble telemetría en EC-05

EC-05 combina la lectura actual de Samsara con `units/last_status` de Logitrack.

**Vinculación de unidades.** La llave ignora mayúsculas, espacios, guiones, ceros a la izquierda y
el identificador `TDR` al inicio o al final: `2234 - TDR`, `2234-TDR`, `2234tdr`, `TDR 02234` y
`2234` son la misma unidad. Las unidades sin coincidencia se imprimen en consola y en el resumen
de Google Chat.

**Distancia entre equipos.** Con ambas unidades detenidas, los dos GPS quedan a ~20 m en promedio
(medido el 2026-10-01: mediana 17 m, p95 53 m). Por eso la tolerancia base es
`distancia_detenido_metros` (150 m). En movimiento, la distancia se explica por el desfase entre
lecturas, así que la tolerancia crece con la velocidad de la unidad en movimiento:
`150 m + velocidad x desfase x factor_tolerancia_movimiento`. Solo se marca `REVISAR` (ubicación
no coincide) cuando la distancia es físicamente imposible a `velocidad_maxima_plausible_kmh`.

**Elección del estatus.**

| Situación | Estatus | Fuente |
|---|---|---|
| Ambas detenidas y cerca + historial Samsara >= `detencion_minima_minutos` | `DETENIDO CONFIRMADO` | Histórico Samsara + Logitrack |
| Ambas en movimiento | `RUTA` | Samsara + Logitrack |
| No coinciden, Logitrack más reciente por >= `desfase_minimo_minutos` | `DETENIDO LOGITRACK` o `RUTA` | Logitrack |
| No coinciden, Samsara igual o más reciente | Lo que diga Samsara (detenido pasa por historial) | Samsara |
| GPS Samsara viejo y Logitrack vigente fuera de geocercas excluidas | `DETENIDO LOGITRACK` o `RUTA` | Logitrack |
| Historial detenido sin doble lectura | `DETENIDO SAMSARA` | Histórico Samsara |
| Historial insuficiente o ubicación imposible | `REVISAR` | — |

Las unidades con GPS Samsara viejo solo se cubren con Logitrack si sus coordenadas Logitrack no caen
dentro de una geocerca excluida (se descarga su geometría desde `/addresses`). Si la geometría no se
puede descargar, no se cubre ninguna. Se desactiva con `usar_logitrack_si_samsara_viejo: false`.

**Geocercas por coordenadas.** Samsara no siempre llena `gps.address` aunque la unidad esté dentro
de la geocerca (caso 2148 en TDR MEX1, 2026-10-01). Por eso, además del ID que reporta Samsara, se
valida la coordenada contra el polígono de las geocercas excluidas por etiqueta y de
`geocercas_especiales_ids`. Las mismas geocercas se usan para la cercanía y para validar las unidades
cubiertas por Logitrack.

**Vigencia de Logitrack.** `max_antiguedad_minutos` es 15: en operación normal el desfase entre
plataformas es de ~4 min. Si la mediana de antigüedad de Logitrack supera 10 min, la consola y el
mensaje de Chat muestran `Logitrack con retraso`; en ese caso las detenciones quedan como
`DETENIDO SAMSARA` en lugar de `DETENIDO CONFIRMADO`.

**Detenidas junto a un patio.** Una unidad detenida (`DETENIDO*` o `REVISAR`) a
`distancia_cercania_geocerca_metros` o menos del borde de una geocerca excluida se omite con motivo
`CERCA DE GEOCERCA`: está esperando afuera del patio y no lleva viaje. Las unidades en `RUTA` o
`TRAFICO LENTO` nunca se omiten por cercanía. En EC-05 el umbral es 400 m: el 2026-10-01 las
detenidas junto a patio quedaron a 86–266 m y la siguiente unidad detenida estaba a 909 m. Los
estados afectados se cambian con `estados_omitir_cerca_geocerca`; con `null` o `0` se desactiva.
Ambos valores van en `perfiles_filtros`. Si no se puede descargar la geometría, la regla no se
aplica.

Los umbrales se configuran dentro de `reportes[].telemetria_logitrack`
(`velocidad_detenido_samsara_mph` en mph y `velocidad_detenido_logitrack_kmh` en km/h). Las
credenciales se cargan desde `.env` mediante `LOGITRACK_USERNAME`, `LOGITRACK_PASSWORD`,
`LOGITRACK_CLIENT_ID` y `LOGITRACK_CLIENT_SECRET`; no deben agregarse a `reportes.json`.

Antes de habilitar un envío, valide la integración sin publicar mensajes:

```powershell
.\.venv\Scripts\python.exe .\EnvioMain.py --dry-run --solo "Reporte EC-05"
```

Para enviar una prueba real al chat de pruebas (`GOOGLE_CHAT_WEBHOOK_URL_PRUEBAS`), sin tocar el
chat de producción ni el correo:

```powershell
.\.venv\Scripts\python.exe .\EnvioMain.py --chat-prueba --solo "Reporte EC-05"
```

Para medir de nuevo distancia y desacuerdos entre plataformas (solo lectura, genera un Excel en
`outputs/`):

```powershell
.\.venv\Scripts\python.exe .\scripts\diagnosticar_samsara_logitrack.py
```

Variables SMTP: `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASSWORD`,
`EMAIL_FROM_ADDRESS` y `EMAIL_FROM_NAME`.

## Agregar EC-01, EC-02, EC-03 u otro equipo

Cada apartado debe ser un objeto independiente dentro de `reportes`. Use `parentTagIds` cuando
la etiqueta visible, por ejemplo `EC-01`, sea padre de otras etiquetas y deban incluirse todos sus
descendientes.

```json
{
  "nombre": "Reporte EC-01",
  "equipo": "Equipo EC-01",
  "activo": true,
  "perfil_filtros": "EC-01",
  "tipo_filtro_etiqueta": "parentTagIds",
  "etiquetas": ["EC-01"],
  "analizar_detenciones": true,
  "contenido": {
    "estados": [],
    "unidades": [],
    "columnas": [
      "Unidad", "Estatus", "Tiempo Detenido", "Motor", "Fecha GPS",
      "Ubicación", "Latitud", "Longitud", "Coordenadas", "Geocerca",
      "Velocidad Mph", "IsEcuSpeed", "PLACAS", "ID ROSTERING"
    ],
    "incluir_omitidas_excel": true
  },
  "entregas": [
    {
      "canal": "google_chat",
      "activo": true,
      "webhook_env": "GOOGLE_CHAT_WEBHOOK_URL_EC01"
    },
    {
      "canal": "correo",
      "activo": false,
      "destinatarios": [],
      "destinatarios_env": "MAIL_TO_EC01",
      "asunto": "{nombre} - {fecha} {hora}",
      "cuerpo": "Hola {equipo}, se adjunta el reporte generado el {fecha} a las {hora}."
    }
  ]
}
```

Los nombres de reporte deben ser únicos. Para EC-02 o EC-03 copie el bloque y sustituya todas las
referencias de `EC-01`, incluyendo nombres de variables de entorno.

### Perfil que excluye unidades en geocerca

Agregue el perfil dentro de `perfiles_filtros`:

```json
"EC-01": {
  "excluir_unidades_en_geocerca": true,
  "etiquetas_geocercas_excluidas": ["EC-01"],
  "etiqueta_ids_geocercas_excluidas": []
}
```

### Perfil que conserva unidades en geocerca

```json
"EC-02": {
  "excluir_unidades_en_geocerca": false,
  "etiquetas_geocercas_excluidas": [],
  "etiqueta_ids_geocercas_excluidas": [],
  "geocercas_especiales_ids": []
}
```

### Variables por equipo

Los secretos y destinatarios van en `.env`, nunca en el JSON versionado:

```dotenv
GOOGLE_CHAT_WEBHOOK_URL_EC01=https://chat.googleapis.com/...
GOOGLE_CHAT_WEBHOOK_URL_EC02=https://chat.googleapis.com/...
GOOGLE_CHAT_WEBHOOK_URL_EC03=https://chat.googleapis.com/...

MAIL_TO_EC01=ec01@empresa.com,supervisor@empresa.com
MAIL_TO_EC02=ec02@empresa.com
MAIL_TO_EC03=ec03@empresa.com
```

La cuenta emisora puede compartirse entre equipos mediante `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`,
`SMTP_PASSWORD`, `EMAIL_FROM_ADDRESS` y `EMAIL_FROM_NAME`.

### Diferencia entre los dos campos `activo`

- `reportes[].activo` decide si el reporte puede ejecutarse y si entra en una ejecución sin
  `--solo`.
- `reportes[].entregas[].activo` decide qué canales se utilizan cuando no se pasa `--canal`.
- `--solo` puede seleccionar explícitamente un reporte aunque su `activo` sea `false`.
- `--canal` puede forzar una entrega aunque su entrega tenga `activo: false`.

Si agrega EC-01, EC-02 y EC-03 con `activo: true`, ejecutar `EnvioMain.py` sin argumentos enviará
todos ellos. Para conservar EC-05 como único envío predeterminado, deje los demás con
`activo: false` y ejecútelos desde sus tareas con `--solo`. Valide cada alta con `--dry-run` antes
de habilitar su tarea programada.

## Comparar con un reporte externo

```powershell
.\.venv\Scripts\python.exe .\tools\trackear_diferencias_samsara.py `
  --externo "C:\ruta\Reporte_Unidades.xls"
```

La auditoría de Samsara se busca en la raíz del repositorio y los resultados se guardan en
`outputs/seguimiento_estatus/`.
