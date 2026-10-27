/**
 * LEGADO DE SALTA — Registro de pedidos en Google Sheets
 *
 * Cómo instalarlo (una sola vez):
 *  1. Creá una planilla nueva en Google Sheets (por ejemplo "Pedidos Legado de Salta").
 *  2. Menú Extensiones → Apps Script. Borrá lo que aparece y pegá todo este archivo.
 *  3. Cambiá CLAVE_PANEL por una clave tuya (es la que vas a usar para entrar a pedidos.html).
 *  4. Guardá. Botón "Implementar" → "Nueva implementación" → tipo "Aplicación web".
 *       - Ejecutar como: Yo
 *       - Quién tiene acceso: Cualquier persona
 *     Autorizá los permisos que pide Google.
 *  5. Copiá la "URL de la aplicación web" (termina en /exec) y pegala en config.js → pedidosUrl.
 *
 * Si más adelante cambiás este código: Implementar → Administrar implementaciones →
 * editar (lápiz) → Versión: "Nueva versión". Así la URL sigue siendo la misma.
 */

const CLAVE_PANEL = 'CAMBIAR-ESTA-CLAVE';
const HOJA = 'Pedidos';

const COLUMNAS = [
  'codigo', 'fecha', 'estado', 'nombre', 'destino', 'pago', 'regalo',
  'items', 'piezas', 'subtotal', 'descuento', 'total', 'nota', 'notaInterna', 'actualizado'
];

const ENCABEZADOS = [
  'Código', 'Fecha', 'Estado', 'Nombre', 'Envío a', 'Pago', 'Regalo',
  'Piezas pedidas', 'Cant.', 'Subtotal', 'Descuento', 'Total', 'Nota del cliente', 'Nota interna', 'Actualizado'
];

const ESTADOS = ['Nuevo', 'Confirmado', 'Pagado', 'Enviado', 'Entregado', 'Cancelado'];

function hoja_() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(HOJA);
  if (!hoja) hoja = libro.insertSheet(HOJA);
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(ENCABEZADOS);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, ENCABEZADOS.length).setFontWeight('bold').setBackground('#1e1612').setFontColor('#ffffff');
  }
  return hoja;
}

function respuesta_(datos) {
  return ContentService.createTextOutput(JSON.stringify(datos)).setMimeType(ContentService.MimeType.JSON);
}

function texto_(valor, largo) {
  const t = String(valor == null ? '' : valor).slice(0, largo || 500);
  // Evita que un texto que empieza con "=" se tome como fórmula en la planilla
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function numero_(valor) {
  const n = Number(valor);
  return isFinite(n) ? n : 0;
}

/** Lista de pedidos para el panel (requiere clave). */
function doGet(e) {
  if ((e.parameter.clave || '') !== CLAVE_PANEL) return respuesta_({ ok: false, error: 'clave' });
  const hoja = hoja_();
  const filas = hoja.getLastRow() > 1
    ? hoja.getRange(2, 1, hoja.getLastRow() - 1, COLUMNAS.length).getValues()
    : [];
  const pedidos = filas
    .filter((f) => f[0])
    .map((f) => {
      const p = {};
      COLUMNAS.forEach((col, i) => {
        p[col] = f[i] instanceof Date ? f[i].toISOString() : f[i];
      });
      return p;
    })
    .reverse();
  return respuesta_({ ok: true, pedidos: pedidos, estados: ESTADOS });
}

/** Alta de pedido (desde la tienda) o cambio de estado (desde el panel). */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const datos = JSON.parse(e.postData.contents || '{}');

    if (datos.action === 'nuevo') {
      const p = datos.pedido || {};
      if (!p.codigo || !p.items) return respuesta_({ ok: false, error: 'datos' });
      hoja_().appendRow([
        texto_(p.codigo, 20),
        new Date(),
        'Nuevo',
        texto_(p.nombre, 120),
        texto_(p.destino, 160),
        texto_(p.pago, 80),
        texto_(p.regalo, 10),
        texto_(p.items, 2000),
        numero_(p.piezas),
        numero_(p.subtotal),
        numero_(p.descuento),
        numero_(p.total),
        texto_(p.nota, 600),
        '',
        new Date()
      ]);
      return respuesta_({ ok: true });
    }

    if (datos.action === 'actualizar') {
      if (datos.clave !== CLAVE_PANEL) return respuesta_({ ok: false, error: 'clave' });
      const hoja = hoja_();
      const codigos = hoja.getRange(2, 1, Math.max(hoja.getLastRow() - 1, 1), 1).getValues();
      const idx = codigos.findIndex((c) => c[0] === datos.codigo);
      if (idx === -1) return respuesta_({ ok: false, error: 'no-existe' });
      const fila = idx + 2;
      if (datos.estado && ESTADOS.indexOf(datos.estado) !== -1) {
        hoja.getRange(fila, COLUMNAS.indexOf('estado') + 1).setValue(datos.estado);
      }
      if (typeof datos.notaInterna === 'string') {
        hoja.getRange(fila, COLUMNAS.indexOf('notaInterna') + 1).setValue(texto_(datos.notaInterna, 600));
      }
      hoja.getRange(fila, COLUMNAS.indexOf('actualizado') + 1).setValue(new Date());
      return respuesta_({ ok: true });
    }

    return respuesta_({ ok: false, error: 'accion' });
  } catch (err) {
    return respuesta_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
