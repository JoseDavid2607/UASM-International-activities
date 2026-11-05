# Agenda Internacional · Facultad de Administración, Uniandes

Cronograma informativo de actividades internacionales (`index.html`) y Manual del requisito internacional
(`manual-requisito-internacional/`). Sitio estático, sin build.

## Cómo llegan los datos a la página

```
Excel + carpeta "evidencias" en Drive ──(GitHub Action, cada 6 h)──▶ data/events.json + data/evidence.json ──▶ index.html
```

- El Excel (hoja `Actividades`) define las actividades. Columnas: Programa, Categoría, Título, Fecha inicio,
  Fecha fin (opcional), Descripción corta, Descripción larga, Estado, Enlace / formulario, Enlace de la sesión,
  Enlace de la grabación.
- Cada carpeta dentro de `evidencias/` cuyo **nombre coincide con el título** de una actividad aporta sus archivos
  como evidencias (se ignoran tildes, mayúsculas y signos). Puede haber carpetas de agrupación intermedias.
- `scripts/build-data.mjs` genera los JSON; la página solo lee esos archivos, así que **la API key no está en el navegador**.

## Configuración (una sola vez)

1. En Google Cloud crea una **API key nueva** con solo la *Drive API* habilitada y **sin** restricción por referrer
   (la usa un servidor, no el navegador). Rota/elimina la key anterior, que quedó en el historial de git.
2. Comparte el Excel y la carpeta `evidencias` como *Cualquier persona con el enlace: Lector*.
3. En GitHub: *Settings → Secrets and variables → Actions → New repository secret* con nombre `DRIVE_API_KEY`.
4. En *Actions → Actualizar datos del cronograma → Run workflow* para generar `data/` la primera vez.
   Después se actualiza solo cada 6 horas (o ejecútalo a mano tras editar el Excel).

## Desarrollo local

```bash
npm install
DRIVE_API_KEY=... node scripts/build-data.mjs      # o XLSX_FILE=ruta/local.xlsx node scripts/build-data.mjs
python3 -m http.server 8000                         # abrir http://localhost:8000
npm test
```
