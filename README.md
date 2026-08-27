# Tablero DEFINIT — Ventas y Otorgamiento por Ítem

Tablero interactivo (HTML/JS puro, sin frameworks) que lee datos agregados y
anonimizados desde `data/*.json`. Esos archivos se generan de dos formas
posibles:

1. **Automática (recomendada):** un GitHub Action (`.github/workflows/sync-data.yml`)
   corre diariamente, se loguea en la API EVUP con credenciales guardadas como
   *secrets* del repo, y sobreescribe `data/dicts.json`, `data/rows.json` y
   `data/meta.json`.
2. **Manual:** exportando el Excel de otorgamiento y regenerando los mismos
   tres archivos (como se hizo para la primera versión del tablero).

En ningún caso el HTML público contiene credenciales, ni nombre/teléfono/e-mail/
CPF/CNPJ/domicilio de clientes — solo campos agregados (vendedora, ítem,
estado, canal, montos, mes, género).

## Puesta en marcha (una sola vez)

1. **Creá el repositorio en GitHub** y subí esta carpeta completa (`index.html`,
   `data/`, `scripts/`, `.github/`).
2. **Activá GitHub Pages:** Settings → Pages → Source: rama `main`, carpeta `/ (root)`.
3. **Cargá las credenciales de la API como secrets** (nunca las escribas en el
   código ni las compartas por chat): Settings → Secrets and variables →
   Actions → New repository secret:
   - `EVUP_LOGIN`: usuario de API del ELOS
   - `EVUP_PASSWORD`: contraseña de ese usuario
4. **Corré el sync por primera vez a mano:** pestaña Actions → "Sync EVUP data"
   → "Run workflow". Revisá los logs — es la primera corrida contra la API real
   y puede necesitar ajustes menores (ver "Notas técnicas" abajo).
5. Cuando el job termine y haga commit de `data/*.json`, refrescá la página de
   GitHub Pages: el tablero va a mostrar "Sincronizado el ..." con la fecha real
   de esa corrida.

A partir de ahí el workflow corre solo todos los días (cron en UTC, ver el
comentario en `sync-data.yml` para cambiar el horario).

## Probar el sync localmente (opcional)

```bash
cd definit-dashboard
EVUP_LOGIN="usuario" EVUP_PASSWORD="clave" node scripts/sync.mjs
```

Requiere Node 18+. Nunca commitees un `.env` ni pegues la contraseña en un
archivo del repo — pasala solo por variable de entorno en la terminal.

## Notas técnicas / qué revisar en la primera corrida real

La documentación pública de la API (`/docs/index.html`) no especifica el
formato exacto de respuesta de los endpoints `List` (paginación, envoltorio),
así que `scripts/sync.mjs` fue escrito de forma defensiva (acepta array plano
o `{items:[...], hasNextPage:...}` con variantes de mayúsculas) pero **no fue
probado contra la API real** porque no tengo credenciales. Si la primera
corrida falla o trae datos incompletos, lo más probable es un desajuste en:

- El nombre exacto del campo de paginación (`hasNextPage` vs `HasNextPage` vs
  ausente) — ver `extractItems()` en `sync.mjs`.
- El esquema del token de login (string plano vs `{ "token": "..." }`).
- Nombres de campo con mayúscula/minúscula distinta a la documentada.

Revisar los logs del Action (o de la corrida local) contra la definición real
para ajustar `sync.mjs` en consecuencia.

## Estructura

```
index.html                     tablero (fetch a ./data/*.json en tiempo de carga)
data/dicts.json                catálogos (ítems, vendedoras, estados, meses, etc.)
data/rows.json                 líneas de ítem anonimizadas (sin PII de clientes)
data/meta.json                 fecha de la última sincronización y totales
scripts/sync.mjs               script Node que llama a la API EVUP
.github/workflows/sync-data.yml   cron diario + botón de ejecución manual
```
