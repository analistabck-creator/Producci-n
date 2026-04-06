# Sistema de Producción - versión lista para Render

Esta versión ya está preparada para subir a Render con estructura de carpetas, `render.yaml`, `Procfile` y almacenamiento persistente en disco montado.

## Estructura exacta

```text
almacen_produccion_render/
├── app.py
├── requirements.txt
├── render.yaml
├── Procfile
├── .env.example
├── .gitignore
├── README.md
├── static/
│   └── styles.css
└── templates/
    ├── base.html
    ├── dashboard.html
    ├── detail.html
    ├── drive.html
    ├── error.html
    ├── history.html
    ├── login.html
    ├── record_form.html
    └── users.html
```

## Qué cambia para Render

- La base de datos ya no depende de la carpeta del proyecto.
- El sistema usa `APP_DATA_DIR` para guardar:
  - `produccion.db`
  - `uploads/`
  - `exports/`
  - `exports/pdfs/`
- En Render, `APP_DATA_DIR` queda apuntando a `/var/data`.
- Ya incluye `gunicorn`.

## Subida a Render

### Opción A: desde GitHub
1. Sube esta carpeta tal cual a un repositorio.
2. En Render crea un nuevo **Web Service**.
3. Conecta el repo.
4. Render leerá `render.yaml` automáticamente.
5. Despliega.

### Opción B: manual desde panel
Si no usas Blueprint:
- Build Command:
  `pip install -r requirements.txt`
- Start Command:
  `gunicorn app:app`

## Variables de entorno recomendadas

Obligatorias:
- `SECRET_KEY`
- `APP_DATA_DIR=/var/data`

Opcionales para Drive:
- `GOOGLE_SERVICE_ACCOUNT_FILE=/var/data/google-service-account.json`
- `GOOGLE_DRIVE_FOLDER_ID=...`
- `GOOGLE_DRIVE_FILE_NAME=historial_produccion_maestro.xlsx`
- `GOOGLE_DRIVE_SYNC_MODE=off|excel|full`

## Disco persistente en Render

Muy importante:
si usas SQLite, fotos, videos y Excel, debes crear un **Persistent Disk**.
En esta versión ya está declarado así en `render.yaml`:

- Mount path: `/var/data`
- Size: `5 GB`

## Recomendación importante

Para producción seria:
- deja `GOOGLE_DRIVE_SYNC_MODE=off` al primer despliegue
- prueba el sistema
- luego activa Drive
- si vas a usar muchos videos, considera pasar luego a PostgreSQL + almacenamiento externo

## Usuarios demo

- admin / admin123
- supervisor / super123
- tecnico / tec123
- almacen / alm123

## Arranque local

```bash
pip install -r requirements.txt
python app.py
```
