# Reportes de unidades

Automatizaciones de telemetría para consultar Samsara, clasificar el estado de las unidades y
distribuir reportes por Google Chat o correo electrónico.

## Flujo principal

`EnvioMain.py` es el punto de entrada configurable. Permite seleccionar etiquetas directas o
padres, aplicar políticas distintas de geocercas por equipo y elegir los canales de entrega.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
Copy-Item .env.example .env
.\.venv\Scripts\python.exe .\EnvioMain.py --dry-run
```

El modo `--dry-run` consulta las fuentes y construye la salida, pero no manda mensajes ni correos.

## Estructura

```text
Unidades/
├── archive/                        Versiones históricas, fuera del flujo activo
├── assets/
│   ├── images/                     Logos e imágenes
│   └── templates/                  Plantillas Excel
├── config/                         Configuración y catálogo Samsara
├── docs/                           Guías operativas
├── outputs/                        Reportes generados (no versionados)
├── tests/                          Pruebas automatizadas
├── tools/                          Utilidades manuales de análisis
├── EnvioMain.py                    Envío configurable por equipo
├── EnvioMain1.py                   Auditoría/comparación de telemetría
├── detenciones.py                  Análisis histórico de detenciones
└── requirements.txt                Dependencias de Python
```

Los demás scripts de la raíz se conservaron en su ruta para no romper tareas programadas o
integraciones existentes.

Las plantillas y logos se resuelven desde `assets/`; no deben copiarse nuevamente a la raíz.

## Configuración

La guía completa está en [docs/configuracion-reportes.md](docs/configuracion-reportes.md).
Para Windows Task Scheduler, consulte [docs/task-scheduler.md](docs/task-scheduler.md).

Variables sensibles y credenciales nunca deben subirse. Use `.env.example` como referencia y
mantenga `.env` y `Credenciales.json` únicamente en el equipo de ejecución.

Antes de publicar, revise [docs/seguridad.md](docs/seguridad.md). El historial contiene versiones
anteriores de `.env`, pero la auditoría documentada confirmó que sus valores estaban vacíos y que
no hay credenciales preparadas para subir.

## Validación

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\.venv\Scripts\python.exe -m py_compile EnvioMain.py EnvioMain1.py detenciones.py
```

## Sincronizar etiquetas

```powershell
.\.venv\Scripts\python.exe .\EnvioMain.py --sincronizar-etiquetas
```

Esto actualiza el catálogo de padres e hijos dentro de `config/` sin modificar etiquetas en
Samsara.
