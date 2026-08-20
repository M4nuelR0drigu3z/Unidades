# Seguridad antes de publicar

## Archivos locales

No deben versionarse:

- `.env`
- `Credenciales.json`
- archivos de cuenta de servicio
- webhooks, tokens o contraseñas
- reportes generados y logs

El repositorio incluye `.env.example` únicamente con nombres de variables y valores no sensibles.

## Historial existente

`.env` estuvo versionado en commits anteriores, pero la auditoría del 2026-08-20 confirmó que todas
sus variables tenían valores vacíos. El escaneo de todos los commits y referencias alcanzables,
del área de staging y del contenido XML de las plantillas Excel no encontró tokens, webhooks,
contraseñas, llaves privadas ni archivos de cuenta de servicio.

Con base en esa revisión, no es necesario reescribir el historial solamente por esos `.env`
vacíos. Si una revisión futura identifica un valor real, debe rotarse de inmediato y eliminarse del
historial antes del siguiente push.

## Revisión recomendada

```powershell
git status
git diff --cached --stat
git grep -n -I -E "BEGIN PRIVATE KEY|chat.googleapis.com|password=|token="
```

No copie resultados de la última búsqueda en tickets o conversaciones si contienen valores reales.
