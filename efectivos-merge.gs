/**
 * Efectivos Junta de Andalucía – Fusionador de CSV a Google Sheets
 * ================================================================
 * Apps Script VINCULADO a una hoja de cálculo de Google.
 *
 * Qué hace:
 *   - Lee todos los archivos .csv de una carpeta de Google Drive
 *     (los que descarga el bookmarklet "Efectivos → CSV").
 *   - Los fusiona en una única pestaña "TODO".
 *   - Elimina duplicados por Especialidad + Orden + NIF (idempotente:
 *     puedes volver a ejecutarlo tras re-descargar sin duplicar nada).
 *   - Respeta el formato exacto del CSV (ISO-8859-1, separador ';') y
 *     mantiene todo como TEXTO para no estropear NIFs, "00-00-00",
 *     notas con coma decimal, códigos con ceros a la izquierda, etc.
 *
 * Cómo instalarlo:
 *   1. Abre (o crea) la hoja de cálculo de Google donde quieres los datos.
 *   2. Menú Extensiones → Apps Script.
 *   3. Pega este archivo completo, guarda.
 *   4. Recarga la hoja de cálculo. Aparecerá el menú "📥 Efectivos".
 *   5. Sube tus CSV a una carpeta de Google Drive.
 *   6. Menú "📥 Efectivos" → "Elegir carpeta de CSV..." (pega el enlace
 *      o el ID de la carpeta). Solo hace falta una vez; queda guardado.
 *   7. Menú "📥 Efectivos" → "Fusionar CSV de la carpeta". Autoriza los
 *      permisos la primera vez y listo.
 */

/*** CONFIGURACIÓN ***/
var CONFIG = {
  HOJA_DESTINO: 'TODO',           // nombre de la pestaña donde se vuelca todo
  SEPARADOR: ';',                 // separador del CSV que genera el extractor
  CHARSET: 'ISO-8859-1',          // codificación del CSV (windows-1252)
  PROP_CARPETA: 'CARPETA_CSV_ID', // clave donde se guarda el ID de la carpeta
  // Columnas del CSV (deben coincidir con las del extractor):
  CABECERAS: ['Especialidad','Orden','NIF','Nombre','Colectivo','Tiempo servicio',
              'Año ingreso','Nota','Centro código','Centro nombre','Centro localidad','Centro provincia'],
  // Índices que forman la clave de deduplicación: Especialidad, Orden, NIF
  CLAVE_DEDUP: [0, 1, 2]
};

/*** MENÚ ***/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📥 Efectivos')
    .addItem('Fusionar CSV de la carpeta', 'fusionarCsv')
    .addSeparator()
    .addItem('Elegir carpeta de CSV...', 'configurarCarpeta')
    .addItem('Ver carpeta configurada', 'verCarpeta')
    .addToUi();
}

/*** CONFIGURAR CARPETA DE DRIVE ***/
function configurarCarpeta() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    'Carpeta de Drive con los CSV',
    'Pega el ENLACE o el ID de la carpeta de Google Drive donde tienes los CSV:',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var id = extraerIdCarpeta(resp.getResponseText());
  if (!id) { ui.alert('No he podido reconocer el ID de la carpeta. Revisa el enlace.'); return; }

  // Validar que la carpeta existe y es accesible.
  try {
    var carpeta = DriveApp.getFolderById(id);
    PropertiesService.getDocumentProperties().setProperty(CONFIG.PROP_CARPETA, id);
    ui.alert('✅ Carpeta configurada:\n' + carpeta.getName());
  } catch (e) {
    ui.alert('No puedo abrir esa carpeta. Comprueba el ID/enlace y que tengas acceso.\n\n' + e.message);
  }
}

function verCarpeta() {
  var ui = SpreadsheetApp.getUi();
  var id = PropertiesService.getDocumentProperties().getProperty(CONFIG.PROP_CARPETA);
  if (!id) { ui.alert('Aún no has configurado ninguna carpeta.'); return; }
  try {
    var carpeta = DriveApp.getFolderById(id);
    ui.alert('Carpeta configurada:\n' + carpeta.getName() + '\n\nID: ' + id + '\n' + carpeta.getUrl());
  } catch (e) {
    ui.alert('La carpeta guardada ya no es accesible. Vuelve a configurarla.\n\n' + e.message);
  }
}

// Acepta un ID pelado o cualquier URL de carpeta de Drive.
function extraerIdCarpeta(texto) {
  texto = (texto || '').trim();
  if (!texto) return '';
  var m = texto.match(/[-\w]{25,}/);   // los IDs de Drive tienen 25+ caracteres
  return m ? m[0] : '';
}

/*** PROCESO PRINCIPAL: leer todos los CSV y fusionar ***/
function fusionarCsv() {
  var ui = SpreadsheetApp.getUi();
  var id = PropertiesService.getDocumentProperties().getProperty(CONFIG.PROP_CARPETA);
  if (!id) {
    ui.alert('Primero configura la carpeta:\n"📥 Efectivos" → "Elegir carpeta de CSV..."');
    return;
  }

  var carpeta;
  try { carpeta = DriveApp.getFolderById(id); }
  catch (e) { ui.alert('No puedo abrir la carpeta configurada. Vuelve a elegirla.\n\n' + e.message); return; }

  var filas = [];          // todas las filas de datos (sin cabecera)
  var mapa = {};           // clave -> índice en "filas" (para deduplicar)
  var archivosLeidos = 0;
  var duplicados = 0;

  var it = carpeta.getFiles();
  while (it.hasNext()) {
    var file = it.next();
    if (!/\.csv$/i.test(file.getName())) continue;   // solo .csv

    var texto;
    try {
      texto = file.getBlob().getDataAsString(CONFIG.CHARSET);
    } catch (e) {
      // Si por lo que sea no es ISO-8859-1, probamos UTF-8 como respaldo.
      texto = file.getBlob().getDataAsString('UTF-8');
    }

    var tabla = Utilities.parseCsv(texto, CONFIG.SEPARADOR);
    if (!tabla || tabla.length === 0) continue;
    archivosLeidos++;

    // Saltamos la cabecera de cada archivo (primera fila).
    for (var r = 1; r < tabla.length; r++) {
      var fila = tabla[r].map(limpiarCelda);
      if (fila.join('').trim() === '') continue;      // fila vacía
      // Normalizamos longitud a la cabecera esperada.
      while (fila.length < CONFIG.CABECERAS.length) fila.push('');
      if (fila.length > CONFIG.CABECERAS.length) fila = fila.slice(0, CONFIG.CABECERAS.length);

      var clave = CONFIG.CLAVE_DEDUP.map(function(i) { return fila[i]; }).join('|');
      if (mapa.hasOwnProperty(clave)) {
        filas[mapa[clave]] = fila;   // reemplaza por la versión más reciente
        duplicados++;
      } else {
        mapa[clave] = filas.length;
        filas.push(fila);
      }
    }
  }

  if (archivosLeidos === 0) { ui.alert('No he encontrado ningún archivo .csv en la carpeta.'); return; }
  if (filas.length === 0)   { ui.alert('He leído ' + archivosLeidos + ' CSV pero no había filas de datos.'); return; }

  // Orden: por Especialidad (asc) y luego por Orden (numérico asc).
  filas.sort(function(a, b) {
    if (a[0] !== b[0]) return a[0] < b[0] ? -1 : 1;
    return (parseInt(a[1], 10) || 0) - (parseInt(b[1], 10) || 0);
  });

  volcarEnHoja(filas);

  ui.alert('✅ Fusión completada\n\n' +
           'Archivos CSV leídos: ' + archivosLeidos + '\n' +
           'Registros únicos: ' + filas.length + '\n' +
           'Duplicados fusionados: ' + duplicados);
}

// Quita comillas envolventes residuales y un apóstrofo inicial (el que el
// extractor añade para forzar texto), dejando el valor limpio.
function limpiarCelda(v) {
  v = (v == null ? '' : String(v)).trim();
  if (v.charAt(0) === "'") v = v.slice(1);
  return v;
}

// Escribe cabecera + filas en la hoja destino, forzando formato TEXTO para
// no perder ceros a la izquierda, "00-00-00", notas con coma decimal, etc.
function volcarEnHoja(filas) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(CONFIG.HOJA_DESTINO);
  if (!hoja) hoja = ss.insertSheet(CONFIG.HOJA_DESTINO);
  hoja.clear();

  var datos = [CONFIG.CABECERAS].concat(filas);
  var nFilas = datos.length;
  var nCols = CONFIG.CABECERAS.length;

  var rango = hoja.getRange(1, 1, nFilas, nCols);
  rango.setNumberFormat('@');   // TEXTO: primero el formato, luego los valores
  rango.setValues(datos);

  // Cabecera en negrita y congelada.
  hoja.getRange(1, 1, 1, nCols).setFontWeight('bold');
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, nCols);
}
