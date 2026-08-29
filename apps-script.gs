/**
 * Conecta la planilla de gastos con la página web.
 *
 * Cómo instalarlo, una sola vez:
 *   1. Abrí la planilla en Google Sheets.
 *   2. Menú Extensiones → Apps Script.
 *   3. Borrá lo que haya y pegá todo este archivo.
 *   4. Guardá (el ícono del disquete).
 *   5. Botón azul "Implementar" → "Nueva implementación".
 *   6. Tipo: "Aplicación web".
 *        Ejecutar como:  Yo
 *        Quién tiene acceso:  Cualquier usuario
 *   7. "Implementar". Aceptá los permisos que pide.
 *   8. Copiá la URL que termina en /exec y pegala en gastos.html.
 */

const HOJA = 'Gastos';
const PRIMERA_FILA = 8;   // la primera fila de datos, debajo de los encabezados
// Columnas: A Fecha · B Hora · C Pagó · D Concepto · E Importe · F Dividir entre

function hoja_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA);
}

function responder_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Devuelve todos los gastos cargados. */
function doGet() {
  try {
    const h = hoja_();
    const ultima = h.getLastRow();
    const gastos = [];

    if (ultima >= PRIMERA_FILA) {
      const filas = h.getRange(PRIMERA_FILA, 1, ultima - PRIMERA_FILA + 1, 6).getValues();
      filas.forEach(function (f, i) {
        if (f[2] === '' || f[4] === '' || f[4] === null) return;   // fila vacía
        gastos.push({
          fila:     PRIMERA_FILA + i,
          fecha:    f[0] instanceof Date ? Utilities.formatDate(f[0], 'Europe/Madrid', 'yyyy-MM-dd')
                                         : String(f[0]),
          hora:     f[1] instanceof Date ? Utilities.formatDate(f[1], 'Europe/Madrid', 'HH:mm')
                                         : String(f[1]),
          pago:     String(f[2]),
          concepto: String(f[3]),
          importe:  Number(f[4]),
          entre:    Number(f[5]) === 2 ? 2 : 3
        });
      });
    }
    return responder_({ ok: true, gastos: gastos });
  } catch (err) {
    return responder_({ ok: false, error: String(err) });
  }
}

/** Agrega o borra un gasto. */
function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
    const h = hoja_();

    if (datos.accion === 'borrar') {
      h.deleteRow(Number(datos.fila));
      return responder_({ ok: true });
    }

    // agregar
    const importe = Number(datos.importe);
    if (!datos.pago || !importe) {
      return responder_({ ok: false, error: 'Faltan el pagador o el importe.' });
    }

    // primera fila libre a partir de PRIMERA_FILA
    let fila = Math.max(h.getLastRow() + 1, PRIMERA_FILA);
    const col_c = h.getRange(PRIMERA_FILA, 3, Math.max(1, h.getLastRow() - PRIMERA_FILA + 1), 1)
                   .getValues();
    for (let i = 0; i < col_c.length; i++) {
      if (col_c[i][0] === '') { fila = PRIMERA_FILA + i; break; }
    }

    const ahora = new Date();
    h.getRange(fila, 1, 1, 6).setValues([[
      datos.fecha || Utilities.formatDate(ahora, 'Europe/Madrid', 'yyyy-MM-dd'),
      datos.hora  || Utilities.formatDate(ahora, 'Europe/Madrid', 'HH:mm'),
      datos.pago,
      datos.concepto || '',
      importe,
      Number(datos.entre) === 2 ? 2 : 3
    ]]);
    h.getRange(fila, 5).setNumberFormat('#,##0.00 "€"');

    return responder_({ ok: true, fila: fila });
  } catch (err) {
    return responder_({ ok: false, error: String(err) });
  }
}
