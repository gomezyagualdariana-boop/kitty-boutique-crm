# Kitty Boutique CRM

CRM web hecho con HTML/CSS/JS puro (sin backend, sin Node) con persistencia
en `localStorage` del navegador, basado en la estructura real del archivo
`Kitty_Boutique_100_Clientes.xlsx`:

```
Fecha, Cliente, Empresa, Producto, Valor USD, Estado, Fuente, Prioridad,
Mes, Notas, Fecha Seguimiento
```

## Cómo abrir la app

No requiere instalación. Necesitas servir los archivos con un servidor local
(no abrir el `.html` directamente con doble clic, porque algunos navegadores
bloquean `fetch`/módulos en `file://`):

```bash
cd "CRM KITTY"
python -m http.server 8000
```

Luego abre `http://localhost:8000/login.html`.

**Usuario demo:** `admin` — **Contraseña demo:** `kitty2026`

## ⚠️ Seguridad del login (léelo)

El login usa SHA-256 (Web Crypto API, `crypto.subtle`) + `sessionStorage`.
Esto es **protección de demostración para uso local en un solo navegador**,
NO seguridad real para un sitio multiusuario expuesto a internet: el hash
de la contraseña vive en `js/auth.js` y cualquiera con las herramientas de
desarrollador del navegador puede leerlo. Para producción real se necesita
un backend con autenticación de servidor.

## Estructura

- `login.html` — login (sin sidebar)
- `dashboard.html` — Resumen: 4 KPIs + 4 gráficos (Chart.js vía CDN)
- `pipeline.html` — Kanban por Estado
- `fuente.html` — Tarjetas + gráfico comparativo por Fuente/Canal
- `leads.html` — Tabla con búsqueda, filtros, paginación, CRUD completo
- `seguimientos.html` — Seguimientos pendientes/vencidos
- `js/data.js` — capa de datos (localStorage) + generador de datos simulados
- `js/auth.js` — autenticación demo
- `js/sidebar.js` — barra lateral común + modal de sincronización
- `css/styles.css` — variables CSS de la paleta (blanco / rosado pastel / rosado fuxia)
- `assets/logo.png` — logo real de Kitty Boutique (con fallback SVG automático si no existe)

Los montos están en **USD**, igual que en el Excel original.

## Sincronización con Google Sheets

Botón "🔗 Sincronizar" en la barra lateral (todas las páginas excepto login).
Es sincronización **manual, de reemplazo completo, en una sola dirección por clic**
(no es fusión automática en tiempo real — no es viable sin un servidor 24/7
detrás de una página HTML estática).

### Paso a paso para desplegar tu Apps Script

1. Sube `sheets/Kitty_Boutique_Sheets.xlsx` a Google Drive y ábrelo con Google Sheets
   (o copia los datos a una hoja nueva tuya, respetando el orden de columnas).
2. En Google Sheets: **Extensiones → Apps Script**.
3. Borra el contenido por defecto y pega el contenido de `sheets/Code.gs`.
4. Guarda (icono de disquete).
5. **Implementar → Nueva implementación**:
   - Tipo: **Aplicación web**
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
6. Autoriza los permisos que te pida Google (es tu propia cuenta — Claude nunca
   tiene acceso a ella, ni te pedirá tu contraseña).
7. Copia la URL que termina en `/exec`.
8. En el CRM, abre "🔗 Sincronizar", pega la URL y presiona "Guardar URL".
9. Usa "⬇ Traer de Sheets" o "⬆ Subir a Sheets" según lo que necesites.

**Importante:** si Claude te ayuda a verificar la URL, leerá la respuesta del
`doGet` para confirmar que el orden de columnas calza con lo que la app espera
(ID, Fecha, Cliente, Empresa, Producto, Valor USD, Estado, Fuente, Prioridad,
Mes, Notas, Fecha Seguimiento) — nunca iniciará sesión en tu cuenta de Google.

## Generador de la versión para Google Sheets

```bash
cd sheets
pip install openpyxl
python generar_kitty_sheets.py
```

Genera `Kitty_Boutique_Sheets.xlsx` con:
- Hoja **Leads**: columna ID + estructura real, listas desplegables y
  formato condicional por Estado.
- Hoja **Dashboard**: KPIs y gráficos con fórmulas (`SUMIFS`/`COUNTIFS`),
  nunca valores fijos. Las tablas auxiliares están alejadas (columnas N–R)
  para no sobreponerse con los gráficos.

> Nota: este archivo se generó sin LibreOffice/Excel instalado en el entorno
> donde se construyó, así que el espaciado se hizo de forma conservadora pero
> no se verificó visualmente al 100%. Revisa el resultado al abrirlo y avisa
> si algún gráfico se sobrepone con una tabla.

## Si cambias la estructura de columnas

Si agregas/quitas/reordenas columnas del CRM después de haber sincronizado
con Google Sheets, el orden esperado por `Code.gs` (mapeo por posición) deja
de calzar. Si esto pasa, hay que: (1) actualizar el arreglo `COLS` en
`Code.gs`, y (2) reimportar el Excel actualizado a tu hoja antes de volver a
sincronizar.
