# DEFINIT: Tablero Mensual de Ventas

**👉 Abrir el tablero:** https://ebudani.github.io/DEFINIT/
(lleva a https://script.google.com/macros/s/AKfycbwadhCyR2cirZ-wKxHdEAArIu2xZPN7maaCrI5kQi64YLSG_fgDNEnw6XSe942mSBjYVA/exec)
(ve los datos quien tenga acceso a la planilla "Facturación 2026")

Tablero web que lee **Facturación 2026** (Google Sheets) en vivo y muestra:

- Venta por local, por día y por mes
- % de cumplimiento de Meta y de Súper
- Ranking de locales y evolución mensual de cada uno
- Proyección de cierre del mes (venta a la fecha + ritmo diario)
- Participación de cada local en el total

Filtros: mes, grupo (propios / franquicias / todos) y local para los gráficos.

## Cómo funciona

Es una **web app de Google Apps Script** que corre con la cuenta de Google de
quien la abre: cada vez que se abre (o se toca *Actualizar*) lee la planilla directamente,
así que siempre muestra lo último que se cargó. No hay copias de los datos
en ningún servidor ni en este repositorio.

- Solo lee las pestañas **Ventas Mensuales** y las mensuales (`Septiembre26`, `Agosto26`, …).
  Cualquier otra pestaña (por ejemplo la de clientes con DNI) no se abre.
- Permisos pedidos, ambos de solo lectura (no puede modificar nada): `spreadsheets.readonly`
  para Facturación 2026 y `drive.readonly` para leer los Excel de los indicadores comerciales.

### Indicadores comerciales

Lee el Excel **"Descuentos por locales…"** más reciente de la carpeta compartida
(`CARPETA_COMERCIAL_ID` en `src/Comercial.js`) tal cual está, sin convertirlo ni copiarlo.
Cuando se sube uno nuevo, el tablero lo toma solo. El cálculo queda en caché 6 h por archivo.
Muestra cantidad de ventas, ticket promedio (sin IVA), % de descuento, mix de cuotas y
formas de pago, por local y mes. Es una sección al final de la página; solo ve los datos quien tiene acceso a esa carpeta en Drive.

| Indicador | Cálculo |
|---|---|
| Cantidad de ventas | Cobros con venta neta > 0 |
| Ticket promedio | Venta neta ÷ 1,21 ÷ cantidad de ventas |
| % de descuento | Descuento ÷ venta bruta, de todos los cobros (mismo criterio que la tabla dinámica del archivo) |
| Mix de cuotas / formas de pago | % de la venta neta según "Condición de Pago" / "Forma de Pago" |
- Acceso: **quien tenga acceso a la planilla en Drive** (`"executeAs": "USER_ACCESSING"`,
  `"access": "ANYONE"` en `src/appsscript.json`). Cualquiera con el link puede abrir la
  página, pero los datos se leen con los permisos de quien la abre: sin acceso a
  "Facturación 2026" no ve nada. Para dar o quitar acceso, se comparte o se deja de
  compartir la planilla. Cada persona autoriza el tablero la primera vez
  (permiso de solo lectura de hojas de cálculo).

```
src/
  Code.js           doGet() + getData(): lee la planilla
  Parser.js         entiende el formato de la planilla (probado con los datos reales)
  Index.html        estructura de la página
  Styles.html       estilos (modo claro y oscuro)
  App.html          KPIs, ranking y gráficos (Chart.js)
  appsscript.json   manifiesto: zona horaria, permisos, acceso
.github/workflows/
  deploy.yml        publica en Apps Script en cada push a main
dev/
  dump_xlsx.py      convierte una copia .xlsx a dev/raw.json (no se versiona)
  build-preview.js  genera dev/index.html para ver el tablero sin desplegar
```

## Puesta en marcha (una sola vez)

Requisitos: Node.js y la cuenta de Google que tiene acceso a la planilla.

1. Activá la API de Apps Script en tu cuenta: <https://script.google.com/home/usersettings> → *Google Apps Script API* → **Activada**.
2. Instalá las dependencias e iniciá sesión (se abre el navegador para autorizar):
   ```bash
   npm install
   npx clasp login
   ```
3. Creá el proyecto de Apps Script vinculado a esta carpeta:
   ```bash
   npx clasp create --type webapp --title "Tablero DEFINIT" --rootDir src
   ```
   Esto genera `.clasp.json` con el ID del script. **Commitealo**: la publicación
   automática lo necesita (el ID no es secreto; sin tu cuenta no da acceso a nada).
   Si `clasp` reescribió `src/appsscript.json`, restauralo con `git checkout src/appsscript.json`.
4. Subí el código y creá la implementación web (una sola vez):
   ```bash
   npm run crear-implementacion
   ```
5. Abrí el script con `npx clasp open`, andá a **Implementar → Administrar implementaciones**,
   copiá la **URL de la aplicación web** y abrila. La primera vez Google pide autorizar
   el acceso de solo lectura a tus hojas de cálculo.

Guardá esa URL en favoritos: es tu tablero.

## Publicación automática desde GitHub

Cada push a `main` que cambie algo en `src/` publica el tablero solo
(`.github/workflows/deploy.yml`): sube el código a Apps Script y actualiza la
misma implementación, así **la URL no cambia**. Se puede ver cada corrida en la
pestaña **Actions** del repo, y lanzarla a mano con *Run workflow*.

Para activarla, cargá un secret en GitHub → **Settings → Secrets and variables
→ Actions → New repository secret**:

| Secret | Qué poner |
|---|---|
| `CLASPRC_JSON` | El contenido completo del archivo `.clasprc.json` que dejó `npx clasp login` en tu carpeta de usuario (`C:\Users\<vos>\.clasprc.json`). |

El ID de la implementación web (`DEPLOYMENT_ID`) está escrito en el workflow: no
es secreto, es parte de la URL del tablero. Si alguna vez creás otra
implementación, actualizalo ahí (sale de `npx clasp deployments`, el que no dice `@HEAD`).

`CLASPRC_JSON` es una credencial de tu cuenta de Google: pegala solo en el secret
de GitHub (queda cifrada), nunca en el código ni en un chat. Si alguna vez querés
revocarla, andá a <https://myaccount.google.com/permissions> y quitá el acceso de *clasp*.

Si la Action falla con un error de autorización, volvé a correr `npx clasp login`
y actualizá el secret `CLASPRC_JSON` con el archivo nuevo.

### Publicar a mano (sin GitHub)

```bash
npm run push
```

Con `push` alcanza para la URL de prueba (`/dev`). Para la URL publicada (`/exec`):

```bash
npx clasp deploy --deploymentId <DEPLOYMENT_ID>
```

## Vista previa local (sin desplegar)

Con una copia `.xlsx` de la planilla (Archivo → Descargar → Microsoft Excel):

```bash
python dev/dump_xlsx.py ruta/a/Facturacion.xlsx
npm run preview
```

y abrí `dev/index.html` en el navegador. `raw.json` e `index.html` quedan fuera de git.

## Cómo se calcula cada cosa

| Indicador | Cálculo |
|---|---|
| Venta | Columna *Venta* de **Ventas Mensuales** (suma de las ventas diarias del mes) |
| % Meta / % Súper | Venta ÷ Meta, Venta ÷ Súper |
| Días transcurridos | Último día del mes con ventas cargadas (sin pasar la fecha de A1) |
| Proyección de cierre | Venta ÷ días transcurridos × días del mes (días corridos) |
| Necesario/día p/ Súper | (Súper − Venta) ÷ días restantes del mes |
| Participación | Venta del local ÷ venta del grupo |
| Mapa semáforo | Venta ÷ Súper; en el mes en curso, proyección ÷ Súper. Verde ≥ 100%, amarillo 97–100%, naranja 90–97%, rojo < 90% |

Todos los importes son nominales (sin ajustar por inflación).

## Si cambia la planilla

El parser busca los datos por su contenido, no por celdas fijas: encabezados
`Local`, bloque `FRANQUICIAS`, fila `Días remanentes`, fechas del mes en la fila 1,
y en las pestañas mensuales cada local con su fila `Ventas` debajo. Agregar
locales o meses nuevos no requiere tocar el código. Si se renombra la pestaña
**Ventas Mensuales** o cambia el ID del archivo, ajustá las constantes al
principio de `src/Code.js`.
