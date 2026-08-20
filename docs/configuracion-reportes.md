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
