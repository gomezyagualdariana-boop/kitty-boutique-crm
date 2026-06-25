/**
 * KITTY BOUTIQUE CRM — Apps Script de sincronización con Google Sheets
 *
 * Espera una hoja llamada "Leads" con columnas EN ESTE ORDEN (por POSICIÓN,
 * no por nombre de encabezado — si cambias el orden de columnas en Sheets,
 * actualiza también las constantes de índice abajo):
 *
 *   A: ID
 *   B: Fecha
 *   C: Cliente
 *   D: Empresa
 *   E: Producto
 *   F: Valor USD
 *   G: Estado
 *   H: Fuente
 *   I: Prioridad
 *   J: Mes
 *   K: Notas
 *   L: Fecha Seguimiento
 *
 * doGet  -> devuelve todos los leads de la hoja como JSON (array de objetos).
 * doPost -> accion "reemplazarTodo": borra todas las filas y reescribe con
 *           el array de leads recibido en el body (JSON).
 *
 * Despliegue (lo hace el propio usuario, Claude no tiene acceso a la cuenta):
 *   1. Abre tu Google Sheet -> Extensiones -> Apps Script.
 *   2. Borra el contenido por defecto y pega este código completo.
 *   3. Guarda (icono de disquete).
 *   4. Implementar -> Nueva implementación -> tipo "Aplicación web".
 *      - Ejecutar como: Yo
 *      - Quién tiene acceso: Cualquier usuario
 *   5. Autoriza los permisos que pida Google (es tu propia cuenta).
 *   6. Copia la URL "/exec" que te entrega y pégala en el botón
 *      "🔗 Sincronizar" del CRM web.
 */

const SHEET_NAME = 'Leads';
const COLS = ['id', 'fecha', 'cliente', 'empresa', 'producto', 'valor', 'estado',
              'fuente', 'prioridad', 'mes', 'notas', 'fechaSeguimiento'];

function _sheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function doGet(e) {
  try {
    const sheet = _sheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return _json([]);
    }
    const range = sheet.getRange(2, 1, lastRow - 1, COLS.length);
    const values = range.getValues();
    const leads = values
      .filter(row => row[0] !== '' && row[0] !== null)
      .map(row => {
        const obj = {};
        COLS.forEach((key, i) => {
          let v = row[i];
          if (v instanceof Date) {
            v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
          }
          obj[key] = v;
        });
        return obj;
      });
    return _json(leads);
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.accion !== 'reemplazarTodo') {
      return _json({ ok: false, error: 'Acción no soportada: ' + body.accion });
    }
    const leads = body.leads || [];
    const sheet = _sheet();

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, COLS.length).clearContent();
    }

    if (leads.length > 0) {
      const rows = leads.map(lead => COLS.map(key => lead[key] !== undefined ? lead[key] : ''));
      sheet.getRange(2, 1, rows.length, COLS.length).setValues(rows);
    }

    return _json({ ok: true, filas: leads.length });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
