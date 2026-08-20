# Programar los envíos en Windows

La configuración predeterminada ejecuta únicamente `Reporte EC-05` y lo envía a Google Chat.
La tarea puede llamar directamente a `pythonw.exe`; no necesita un script `.cmd`.

## Tarea EC-05 por Google Chat

En el servidor actual, configure la acción **Iniciar un programa** así:

- Programa o script: `C:\Users\Administrator\AppData\Local\Programs\Python\Python313\pythonw.exe`
- Agregar argumentos: `"C:\Users\Administrator\source\repos\Unidades-1\EnvioMain.py" --solo "Reporte EC-05" --canal google_chat`
- Iniciar en: `C:\Users\Administrator\source\repos\Unidades-1`

El webhook se obtiene de `GOOGLE_CHAT_WEBHOOK_URL` en `.env`.

## Tarea EC-05 por correo

Para generar el Excel y mandarlo por correo, duplique la tarea y configure:

- Programa o script: `C:\Users\Administrator\AppData\Local\Programs\Python\Python313\pythonw.exe`
- Agregar argumentos: `"C:\Users\Administrator\source\repos\Unidades-1\EnvioMain.py" --solo "Reporte EC-05" --canal correo`
- Iniciar en: `C:\Users\Administrator\source\repos\Unidades-1`

Los destinatarios se obtienen de `MAIL_TO` en `.env`, separados por coma. También deben estar
configuradas `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER` y `SMTP_PASSWORD`.

## Recomendaciones

- Use la cuenta de Windows que tiene acceso al repositorio y a `.env`.
- Elija **Ejecutar tanto si el usuario inició sesión como si no** si debe operar sin sesión abierta.
- En **Si la tarea ya se está ejecutando**, seleccione **No iniciar una instancia nueva**.
- Configure reintentos si la tarea falla por conectividad.
- No habilite privilegios elevados salvo que la política del equipo lo requiera.
- Verifique que Python 3.13 tenga instaladas las dependencias de `requirements.txt`.
- Confirme que `.env`, `config/`, `assets/` y las credenciales existan dentro de `Unidades-1`.

`pythonw.exe` no abre consola. Para diagnosticar errores, ejecute temporalmente la misma acción con
`python.exe` o pruebe el comando desde PowerShell y revise `reporte_logs.log`.

Antes de habilitar cada tarea, valide manualmente sin enviar:

```powershell
cd C:\Users\Administrator\source\repos\Unidades-1
& "C:\Users\Administrator\AppData\Local\Programs\Python\Python313\python.exe" `
  ".\EnvioMain.py" --solo "Reporte EC-05" --canal google_chat --dry-run
& "C:\Users\Administrator\AppData\Local\Programs\Python\Python313\python.exe" `
  ".\EnvioMain.py" --solo "Reporte EC-05" --canal correo --dry-run
```

El primer comando imprime el mensaje. El segundo genera el Excel de prueba dentro de
`outputs/envio_previews/`; ninguno realiza un envío real.

## Agregar tareas para EC-01, EC-02, EC-03 u otro equipo

Duplique la tarea existente y cambie solamente el valor de `--solo` y, cuando corresponda, el de
`--canal`. Por ejemplo, para EC-02:

- Chat: `"C:\Users\Administrator\source\repos\Unidades-1\EnvioMain.py" --solo "Reporte EC-02" --canal google_chat`
- Correo: `"C:\Users\Administrator\source\repos\Unidades-1\EnvioMain.py" --solo "Reporte EC-02" --canal correo`

El campo **Programa o script** y el campo **Iniciar en** permanecen iguales para todos los equipos.

Después cree una tarea de Windows por cada combinación necesaria:

| Nombre sugerido | Reporte | Canal |
| --- | --- | --- |
| Unidades EC-01 Chat | Reporte EC-01 | Google Chat |
| Unidades EC-01 Correo | Reporte EC-01 | Excel por correo |
| Unidades EC-02 Chat | Reporte EC-02 | Google Chat |
| Unidades EC-02 Correo | Reporte EC-02 | Excel por correo |
| Unidades EC-03 Chat | Reporte EC-03 | Google Chat |
| Unidades EC-03 Correo | Reporte EC-03 | Excel por correo |

En todos los casos, **Iniciar en** debe apuntar a la raíz del repositorio. El nombre que aparece
después de `--solo` debe coincidir exactamente con `nombre` en `config/reportes.json` y el reporte
debe tener `activo: true`.

La configuración completa del bloque de reporte, perfiles de geocerca y variables de entorno se
encuentra en [configuracion-reportes.md](configuracion-reportes.md).
