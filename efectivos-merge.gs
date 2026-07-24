/**
 * Efectivos Junta de Andalucía – Fusionador de CSV a Google Sheets
 * ================================================================
 * Apps Script VINCULADO a una hoja de cálculo de Google.
 *
 * Flujo previsto:
 *   - Vas soltando los CSV que te van pasando en una carpeta de Drive.
 *   - Hay DOS fases: resolución PROVISIONAL y resolución DEFINITIVA.
 *     Cada fase tiene su propia carpeta de Drive y su propia pestaña.
 *   - Cuando quieras, pulsas "Fusionar": el script relee TODA la carpeta
 *     de esa fase y reconstruye su pestaña (idempotente, sin duplicados).
 *     Así puedes ir añadiendo archivos poco a poco y refusionar cuando toque.
 *   - Al llegar la definitiva, "Comparar" genera una pestaña con quién ha
 *     cambiado de destino entre la provisional y la definitiva.
 *
 * Cómo instalarlo:
 *   1. Abre (o crea) la hoja de cálculo de Google donde quieres los datos.
 *   2. Menú Extensiones → Apps Script. Pega este archivo completo y guarda.
 *   3. Recarga la hoja. Aparecerá el menú "📥 Efectivos".
 *   4. Crea en Drive DOS carpetas (una para provisional, otra para definitiva).
 *   5. Menú "📥 Efectivos" → "Elegir carpeta provisional..." (y la definitiva
 *      cuando la tengas). Solo hace falta una vez; queda guardado.
 *   6. Suelta tus CSV en la carpeta y pulsa "Fusionar provisional".
 *      Autoriza los permisos la primera vez y listo.
 */

/*** CONFIGURACIÓN ***/
var CONFIG = {
  SEPARADOR: ';',                 // separador del CSV que genera el extractor
  CHARSET: 'ISO-8859-1',          // codificación del CSV (windows-1252)
  // Columnas del CSV (deben coincidir con las del extractor):
  CABECERAS: ['Especialidad','Orden','NIF','Nombre','Colectivo','Tiempo servicio',
              'Año ingreso','Nota','Centro código','Centro nombre','Centro localidad','Centro provincia'],
  // Índices de columna útiles (0-based, según CABECERAS):
  COL: { ESP:0, ORDEN:1, NIF:2, NOMBRE:3, CENTRO_COD:8, CENTRO_NOM:9, CENTRO_LOC:10, CENTRO_PROV:11 },
  // Clave para deduplicar filas dentro de una fase: Especialidad+Orden+NIF.
  CLAVE_DEDUP: [0, 1, 2],
  HOJA_COMPARATIVA: 'COMPARATIVA',
  // Las dos fases del proceso. Cada una: propiedad donde se guarda el ID de su
  // carpeta de Drive, y nombre de la pestaña destino.
  FASES: {
    provisional: { prop: 'CARPETA_CSV_PROVISIONAL', hoja: 'PROVISIONAL', etiqueta: 'provisional' },
    definitiva:  { prop: 'CARPETA_CSV_DEFINITIVA',  hoja: 'DEFINITIVA',  etiqueta: 'definitiva' }
  }
};

/*** MENÚ ***/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📥 Efectivos')
    .addItem('Fusionar provisional', 'fusionarProvisional')
    .addItem('Fusionar definitiva', 'fusionarDefinitiva')
    .addSeparator()
    .addItem('Comparar provisional ↔ definitiva', 'compararFases')
    .addSeparator()
    .addItem('Elegir carpeta provisional...', 'configurarCarpetaProvisional')
    .addItem('Elegir carpeta definitiva...', 'configurarCarpetaDefinitiva')
    .addItem('Ver carpetas configuradas', 'verCarpetas')
    .addToUi();
}

/*** ATAJOS DE MENÚ (Apps Script no pasa argumentos a los items) ***/
function fusionarProvisional() { fusionar('provisional'); }
function fusionarDefinitiva()  { fusionar('definitiva'); }
function configurarCarpetaProvisional() { configurarCarpeta('provisional'); }
function configurarCarpetaDefinitiva()  { configurarCarpeta('definitiva'); }

/*** CONFIGURAR CARPETA DE DRIVE ***/
function configurarCarpeta(faseId) {
  var fase = CONFIG.FASES[faseId];
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    'Carpeta de Drive – resolución ' + fase.etiqueta,
    'Pega el ENLACE o el ID de la carpeta de Google Drive con los CSV ' + fase.etiqueta + 's:',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var id = extraerIdCarpeta(resp.getResponseText());
  if (!id) { ui.alert('No he podido reconocer el ID de la carpeta. Revisa el enlace.'); return; }

  try {
    var carpeta = DriveApp.getFolderById(id);
    PropertiesService.getDocumentProperties().setProperty(fase.prop, id);
    ui.alert('✅ Carpeta ' + fase.etiqueta + ' configurada:\n' + carpeta.getName());
  } catch (e) {
    ui.alert('No puedo abrir esa carpeta. Comprueba el ID/enlace y que tengas acceso.\n\n' + e.message);
  }
}

function verCarpetas() {
  var props = PropertiesService.getDocumentProperties();
  var lineas = [];
  Object.keys(CONFIG.FASES).forEach(function(faseId) {
    var fase = CONFIG.FASES[faseId];
    var id = props.getProperty(fase.prop);
    if (!id) { lineas.push('• ' + fase.etiqueta + ': (sin configurar)'); return; }
    try {
      var c = DriveApp.getFolderById(id);
      lineas.push('• ' + fase.etiqueta + ': ' + c.getName() + '\n  ' + c.getUrl());
    } catch (e) {
      lineas.push('• ' + fase.etiqueta + ': (guardada pero inaccesible, reconfigúrala)');
    }
  });
  SpreadsheetApp.getUi().alert('Carpetas configuradas:\n\n' + lineas.join('\n'));
}

// Acepta un ID pelado o cualquier URL de carpeta de Drive.
function extraerIdCarpeta(texto) {
  texto = (texto || '').trim();
  if (!texto) return '';
  var m = texto.match(/[-\w]{25,}/);   // los IDs de Drive tienen 25+ caracteres
  return m ? m[0] : '';
}

/*** FUSIONAR: leer todos los CSV de una fase y volcarlos en su pestaña ***/
function fusionar(faseId) {
  var fase = CONFIG.FASES[faseId];
  var ui = SpreadsheetApp.getUi();
  var id = PropertiesService.getDocumentProperties().getProperty(fase.prop);
  if (!id) {
    ui.alert('Primero configura la carpeta ' + fase.etiqueta + ':\n' +
             '"📥 Efectivos" → "Elegir carpeta ' + fase.etiqueta + '..."');
    return;
  }

  var carpeta;
  try { carpeta = DriveApp.getFolderById(id); }
  catch (e) { ui.alert('No puedo abrir la carpeta ' + fase.etiqueta + '. Vuelve a elegirla.\n\n' + e.message); return; }

  var res = leerFilasDeCarpeta(carpeta);
  if (res.archivos === 0) { ui.alert('No he encontrado ningún archivo .csv en la carpeta ' + fase.etiqueta + '.'); return; }
  if (res.filas.length === 0) { ui.alert('He leído ' + res.archivos + ' CSV pero no había filas de datos.'); return; }

  // Orden: por Especialidad (asc) y luego por Orden (numérico asc).
  res.filas.sort(function(a, b) {
    if (a[CONFIG.COL.ESP] !== b[CONFIG.COL.ESP]) return a[CONFIG.COL.ESP] < b[CONFIG.COL.ESP] ? -1 : 1;
    return (parseInt(a[CONFIG.COL.ORDEN], 10) || 0) - (parseInt(b[CONFIG.COL.ORDEN], 10) || 0);
  });

  volcarEnHoja(fase.hoja, CONFIG.CABECERAS, res.filas);

  ui.alert('✅ Fusión ' + fase.etiqueta + ' completada\n\n' +
           'Archivos CSV leídos: ' + res.archivos + '\n' +
           'Registros únicos: ' + res.filas.length + '\n' +
           'Duplicados fusionados: ' + res.duplicados);
}

// Lee todos los .csv de una carpeta y devuelve filas deduplicadas.
function leerFilasDeCarpeta(carpeta) {
  var filas = [], mapa = {}, archivos = 0, duplicados = 0;
  var it = carpeta.getFiles();
  while (it.hasNext()) {
    var file = it.next();
    if (!/\.csv$/i.test(file.getName())) continue;

    var texto;
    try { texto = file.getBlob().getDataAsString(CONFIG.CHARSET); }
    catch (e) { texto = file.getBlob().getDataAsString('UTF-8'); }

    var tabla = Utilities.parseCsv(texto, CONFIG.SEPARADOR);
    if (!tabla || tabla.length === 0) continue;
    archivos++;

    for (var r = 1; r < tabla.length; r++) {   // r=1: saltamos la cabecera
      var fila = tabla[r].map(limpiarCelda);
      if (fila.join('').trim() === '') continue;
      while (fila.length < CONFIG.CABECERAS.length) fila.push('');
      if (fila.length > CONFIG.CABECERAS.length) fila = fila.slice(0, CONFIG.CABECERAS.length);

      var clave = CONFIG.CLAVE_DEDUP.map(function(i) { return fila[i]; }).join('|');
      if (mapa.hasOwnProperty(clave)) { filas[mapa[clave]] = fila; duplicados++; }
      else { mapa[clave] = filas.length; filas.push(fila); }
    }
  }
  return { filas: filas, archivos: archivos, duplicados: duplicados };
}

/*** COMPARAR PROVISIONAL ↔ DEFINITIVA ***/
function compararFases() {
  var ui = SpreadsheetApp.getUi();
  var prov = leerHoja(CONFIG.FASES.provisional.hoja);
  var defi = leerHoja(CONFIG.FASES.definitiva.hoja);
  if (!prov) { ui.alert('No existe la pestaña "' + CONFIG.FASES.provisional.hoja + '". Fusiona antes la provisional.'); return; }
  if (!defi) { ui.alert('No existe la pestaña "' + CONFIG.FASES.definitiva.hoja + '". Fusiona antes la definitiva.'); return; }

  // Indexamos cada fase por persona (NIF + Nombre, que va sin enmascarar).
  var mapaProv = indexarPorPersona(prov);
  var mapaDefi = indexarPorPersona(defi);

  var claves = {};
  Object.keys(mapaProv).forEach(function(k) { claves[k] = true; });
  Object.keys(mapaDefi).forEach(function(k) { claves[k] = true; });

  var C = CONFIG.COL;
  var salida = [];
  var contadores = { igual: 0, cambio: 0, soloProv: 0, soloDefi: 0 };

  Object.keys(claves).forEach(function(k) {
    var p = mapaProv[k], d = mapaDefi[k];
    var base = (p || d);
    var nif = base[C.NIF], nombre = base[C.NOMBRE];
    var estado;

    if (p && d) {
      var mismoCentro = (p[C.CENTRO_COD] === d[C.CENTRO_COD]) && (p[C.ESP] === d[C.ESP]);
      estado = mismoCentro ? 'Igual' : 'Cambio de destino';
      mismoCentro ? contadores.igual++ : contadores.cambio++;
    } else if (p) { estado = 'Solo provisional'; contadores.soloProv++; }
    else          { estado = 'Solo definitiva';  contadores.soloDefi++; }

    salida.push([
      nif, nombre, estado,
      p ? p[C.ESP] : '', p ? p[C.ORDEN] : '', p ? p[C.CENTRO_COD] : '',
      p ? p[C.CENTRO_NOM] : '', p ? p[C.CENTRO_LOC] : '', p ? p[C.CENTRO_PROV] : '',
      d ? d[C.ESP] : '', d ? d[C.ORDEN] : '', d ? d[C.CENTRO_COD] : '',
      d ? d[C.CENTRO_NOM] : '', d ? d[C.CENTRO_LOC] : '', d ? d[C.CENTRO_PROV] : ''
    ]);
  });

  // Orden: primero los cambios, luego solo-una-fase, luego iguales; y por nombre.
  var prioridad = { 'Cambio de destino': 0, 'Solo definitiva': 1, 'Solo provisional': 2, 'Igual': 3 };
  salida.sort(function(a, b) {
    if (prioridad[a[2]] !== prioridad[b[2]]) return prioridad[a[2]] - prioridad[b[2]];
    return a[1] < b[1] ? -1 : (a[1] > b[1] ? 1 : 0);
  });

  var cabeceras = ['NIF','Nombre','Estado',
    'Esp. prov.','Orden prov.','Centro cód. prov.','Centro prov.','Localidad prov.','Provincia prov.',
    'Esp. def.','Orden def.','Centro cód. def.','Centro def.','Localidad def.','Provincia def.'];
  volcarEnHoja(CONFIG.HOJA_COMPARATIVA, cabeceras, salida);

  SpreadsheetApp.getUi().alert('✅ Comparativa generada\n\n' +
    'Cambios de destino: ' + contadores.cambio + '\n' +
    'Solo en definitiva (nuevos): ' + contadores.soloDefi + '\n' +
    'Solo en provisional (ya no están): ' + contadores.soloProv + '\n' +
    'Sin cambios: ' + contadores.igual);
}

// Indexa las filas de una hoja por persona: NIF + '|' + Nombre.
function indexarPorPersona(filas) {
  var C = CONFIG.COL, mapa = {};
  filas.forEach(function(fila) {
    var clave = (fila[C.NIF] || '') + '|' + (fila[C.NOMBRE] || '');
    mapa[clave] = fila;
  });
  return mapa;
}

/*** UTILIDADES DE HOJA ***/
// Devuelve las filas de datos (sin cabecera) de una pestaña, o null si no existe.
function leerHoja(nombre) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!hoja) return null;
  var valores = hoja.getDataRange().getValues();
  if (valores.length <= 1) return [];
  return valores.slice(1).map(function(fila) {
    return fila.map(function(v) { return v == null ? '' : String(v); });
  });
}

// Quita comillas residuales y un apóstrofo inicial (el que el extractor añade
// para forzar texto), dejando el valor limpio.
function limpiarCelda(v) {
  v = (v == null ? '' : String(v)).trim();
  if (v.charAt(0) === "'") v = v.slice(1);
  return v;
}

// Escribe cabecera + filas en una pestaña, forzando formato TEXTO para no
// perder ceros a la izquierda, "00-00-00", notas con coma decimal, etc.
function volcarEnHoja(nombreHoja, cabeceras, filas) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) hoja = ss.insertSheet(nombreHoja);
  hoja.clear();

  var datos = [cabeceras].concat(filas);
  var nCols = cabeceras.length;
  var rango = hoja.getRange(1, 1, datos.length, nCols);
  rango.setNumberFormat('@');   // TEXTO: primero el formato, luego los valores
  rango.setValues(datos);

  hoja.getRange(1, 1, 1, nCols).setFontWeight('bold');
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, nCols);
}
