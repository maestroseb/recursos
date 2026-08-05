// ============================================================
//  COLOCACIÓN DE EFECTIVOS — Apps Script (hoja de efectivos)
// ============================================================
//  ⚠️ Este es el script COMPLETO del proyecto de la HOJA DE EFECTIVOS.
//     Reemplaza el script actual (el que tenía mergeSheets / onOpen
//     "🛠️ Herramientas personalizadas") por este archivo.
//     NO mezclar con el proyecto del Concurso de Traslados.
//
//  Nomenclatura (espejo del CGT, que usa _Prov / _Def):
//    · Las pestañas de cada especialidad terminan en "_Prov" o "_Def".
//        Ej.: "Inf_Prov", "EF_Prov", ...  y  "Inf_Def", "EF_Def", ...
//    · Consolidar Provisional  →  fusiona  *_Prov           en  "Global_Prov"
//    · Consolidar Definitiva   →  fusiona  *_Def + SIN sufijo en  "Global_Def"
//      (las pestañas que NO terminen en "_Prov" se consideran definitivas)
//    · Quedan SIEMPRE fuera: Especialidades, Centros, BÚSQUEDA y los destinos.
//    · Exportar genera efectivos_AAAA.json con { datos, datosDef, hayDef }.
//
//  Columnas esperadas en cada pestaña (fila 1 = cabecera), A–L:
//    A Especialidad (código)   B Orden           C NIF            D Nombre
//    E Colectivo               F Tiempo_servicio G Año_ingreso    H Nota
//    I Centro_código           J Centro_nombre   K Centro_localidad L Centro_provincia
//
//  Los códigos de especialidad y centro se interpretan igual que en
//  diccionarios.json (el mismo fichero que el Concurso de Traslados).
// ============================================================

// Destinos de la consolidación
const EFECTIVOS_DEST_PROV = 'Global_Prov';
const EFECTIVOS_DEST_DEF  = 'Global_Def';

// Pestañas que SIEMPRE quedan fuera de cualquier consolidación
const EFECTIVOS_EXCLUIR = ['Especialidades', 'Centros', 'BÚSQUEDA',
  'MergedData', 'Merged_Data', EFECTIVOS_DEST_PROV, EFECTIVOS_DEST_DEF];

// --- Columnas de las pestañas de diccionario (1-based) ---
// Pestaña "Especialidades": A = "(código) Nombre", B = código exacto, C = nombre
const ESP_COL_CODIGO = 2;   // B (código que coincide con los datos)
const ESP_COL_NOMBRE = 3;   // C (nombre largo)
// Pestaña "Centros":
const CENTROS_COL_CODIGO    = 1;   // A
const CENTROS_COL_NOMBRE    = 5;   // E
const CENTROS_COL_LOCALIDAD = 7;   // G
const CENTROS_COL_PROVINCIA = 11;  // K
const CENTROS_COL_LAT       = 16;  // P (N_LATITUD)
const CENTROS_COL_LON       = 17;  // Q (N_LONGITUD)

// ===================== MENÚ =====================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📂 Colocación de Efectivos')
    .addItem('🔄 Importar y consolidar PROVISIONAL', 'importarYConsolidarProv')
    .addItem('🔄 Importar y consolidar DEFINITIVA', 'importarYConsolidarDef')
    .addSeparator()
    .addItem('📥 Solo importar CSV provisional (carpeta → *_Prov)', 'importarProv')
    .addItem('📥 Solo importar CSV definitiva (carpeta → *_Def)', 'importarDef')
    .addItem('📋 Solo consolidar Provisional (*_Prov → Global_Prov)', 'consolidarProv')
    .addItem('📋 Solo consolidar Definitiva (resto → Global_Def)', 'consolidarDef')
    .addSeparator()
    .addItem('📤 Exportar JSON Efectivos (Vercel)', 'exportarEfectivosJSON')
    .addItem('📤 Exportar diccionarios.json (esp. y centros)', 'exportarDiccionariosJSON')
    .addSeparator()
    .addItem('📁 Elegir carpeta CSV provisional...', 'elegirCarpetaProv')
    .addItem('📁 Elegir carpeta CSV definitiva...', 'elegirCarpetaDef')
    .addItem('⏱️ Activar auto-import cada hora', 'activarAutoImport')
    .addItem('⏹️ Desactivar auto-import', 'desactivarAutoImport')
    .addToUi();
}

// NOTA: no usamos un trigger onChange para reconsolidar automáticamente.
// La consolidación crea/recrea las hojas Global_*, lo que a su vez genera
// eventos INSERT_GRID y provocaría una recursión ("Ya existe una hoja...").
// La consolidación se lanza manualmente desde el menú.

// ===================== CONSOLIDACIÓN =====================

// Provisional: solo las pestañas que terminan en "_Prov".
function consolidarProv() {
  return _consolidar(function (nombre) {
    return nombre.slice(-5) === '_Prov';
  }, EFECTIVOS_DEST_PROV);
}

// Definitiva: todo lo que NO sea "_Prov" ni esté excluido (incluye sin sufijo y "_Def").
function consolidarDef() {
  return _consolidar(function (nombre) {
    return nombre.slice(-5) !== '_Prov';
  }, EFECTIVOS_DEST_DEF);
}

// Nº de columnas relevantes (A–L). Solo se leen/escriben estas, para no
// arrastrar columnas auxiliares de las hojas de origen y acelerar el proceso.
const EFECTIVOS_NUM_COLS = 12;

// Fusiona en `destino` las pestañas que cumplan `incluir(nombre)` (y no estén excluidas).
// Mantiene una sola fila de cabecera (la de la primera hoja) y omite filas vacías.
function _consolidar(incluir, destino) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const NC = EFECTIVOS_NUM_COLS;
  let merged = [];

  sheets.forEach(function (sheet) {
    const nombre = sheet.getName();
    if (EFECTIVOS_EXCLUIR.indexOf(nombre) >= 0) return; // fuera siempre
    if (!incluir(nombre)) return;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return;

    // Leer solo las columnas A–L (o las que haya si son menos).
    // getDisplayValues() = el texto tal como se ve en la celda: así el Tiempo de
    // servicio (AA-MM-DD) no se convierte en fecha ni se arrastran formatos numéricos.
    const cols = Math.min(lastCol, NC);
    const data = sheet.getRange(1, 1, lastRow, cols).getDisplayValues();

    data.forEach(function (row, index) {
      // Omitir la cabecera de las hojas siguientes
      if (index === 0 && merged.length > 0) return;
      if (row.join('').trim() !== '') {
        // Normalizar a NC columnas (rellenar si la hoja tiene menos)
        while (row.length < NC) row.push('');
        merged.push(row);
      }
    });
  });

  // Reutiliza la hoja destino si ya existe (evita "Ya existe una hoja...").
  let hoja = ss.getSheetByName(destino);
  if (!hoja) {
    try {
      hoja = ss.insertSheet(destino);
    } catch (e) {
      // Si otra ejecución la creó entre medias, recupérala en lugar de fallar.
      hoja = ss.getSheetByName(destino);
      if (!hoja) throw e;
    }
  }
  hoja.clear();
  if (merged.length > 0) {
    const rng = hoja.getRange(1, 1, merged.length, NC);
    rng.setNumberFormat('@'); // texto sin formato: evita que AA-MM-DD se vuelva fecha
    rng.setValues(merged);
  }
  return merged.length;
}

// ===================== LECTURA PARA EL JSON =====================

// Enmascara el NIF dejando visibles solo los dígitos centrales (***1234**),
// igual que en el Concurso de Traslados, para no exponer el dato completo.
function _maskNif(nif) {
  nif = String(nif || '').trim();
  if (nif.length < 5) return nif ? '***' : '';
  const start = 3;
  const end = nif.length - 2;
  let out = '';
  for (let i = 0; i < nif.length; i++) {
    out += (i >= start && i < end) ? nif.charAt(i) : '*';
  }
  return out;
}

// Lee una pestaña consolidada y devuelve el array de registros (columnas A–L).
function _leerMerged(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja || hoja.getLastRow() < 2) return [];

  const rango = hoja.getDataRange();
  const datos = rango.getValues();
  // Valores tal como se muestran en la celda: evita que el Tiempo de servicio
  // (AA-MM-DD) salga como objeto Date ("Sun Apr 23 2000...").
  const disp = rango.getDisplayValues();
  const filas = [];
  for (let i = 1; i < datos.length; i++) {
    const f = datos[i];
    if (!f[0] && !f[3]) continue; // sin especialidad y sin nombre → fila vacía
    filas.push({
      esp: String(f[0] || '').trim(),               // A Especialidad (código)
      orden: String(f[1] || '').trim(),             // B Orden
      nif: _maskNif(f[2]),                           // C NIF (enmascarado)
      nombre: String(f[3] || '').trim(),            // D Nombre
      colectivo: String(f[4] || '').trim(),         // E Colectivo
      motivo: '',                                    // ya no hay columna Motivo; se mantiene vacío en el JSON
      tiempoServicio: String(disp[i][5] || '').trim(), // F Tiempo_servicio (valor mostrado)
      anioIngreso: String(f[6] || '').trim(),       // G Año_ingreso
      nota: String(f[7] || '').trim(),              // H Nota
      centro: String(f[8] || '').trim(),            // I Centro_código
      // El centro ahora viene en columnas separadas (J nombre, K localidad,
      // L provincia). Mantenemos el JSON igual que antes: prov = provincia y
      // centroLoc = "Nombre, Localidad" (la web resuelve el resto por el código).
      prov: String(f[11] || '').trim(),             // L Centro_provincia
      centroLoc: [String(f[9] || '').trim(), String(f[10] || '').trim()].filter(String).join(', '), // J + K
      _idx: i
    });
  }
  return filas;
}

// ===================== EXPORTAR JSON =====================

// Exporta efectivos_AAAA.json para subir a vercel-efectivos/data/
function exportarEfectivosJSON() {
  const ui = SpreadsheetApp.getUi();

  const datosProv = _leerMerged(EFECTIVOS_DEST_PROV);
  const datosDef  = _leerMerged(EFECTIVOS_DEST_DEF);

  if (datosProv.length === 0 && datosDef.length === 0) {
    ui.alert('⚠️ No hay datos en "' + EFECTIVOS_DEST_PROV + '" ni en "' + EFECTIVOS_DEST_DEF +
      '". Ejecuta antes "Consolidar Provisional" o "Consolidar Definitiva".');
    return;
  }

  // Año del listado: se pregunta (cada año generas un efectivos_AAAA.json distinto).
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matchAnio = ss.getName().match(/(\d{4})\s*[\/\-]\s*(\d{4})/);
  const anioSugerido = matchAnio ? parseInt(matchAnio[2], 10) : new Date().getFullYear();

  const resp = ui.prompt('Año del listado',
    'Genera "efectivos_AAAA.json". ¿De qué año son estos datos?\n' +
    '(p.ej. 2025 para el curso 2024/2025)',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const textoAnio = resp.getResponseText().trim();
  const anio = /^\d{4}$/.test(textoAnio) ? parseInt(textoAnio, 10) : anioSugerido;
  const curso = (anio - 1) + '/' + anio;
  const nombreArchivo = 'efectivos_' + anio;

  const payload = {
    datos: datosProv,
    datosDef: datosDef,
    hayDef: datosDef.length > 0
  };
  const json = JSON.stringify(payload);

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: -apple-system, sans-serif; padding: 20px; color: #333; }
      h3 { color: #1a5c2e; margin: 0 0 12px; }
      .stats { font-size: 13px; margin-bottom: 16px; }
      .stats b { color: #1a5c2e; }
      .file-row { display: flex; align-items: center; gap: 10px; margin: 10px 0; padding: 12px; background: #f5f7fa; border-radius: 8px; }
      .file-row code { font-size: 13px; font-weight: 600; flex: 1; }
      .file-row .size { font-size: 12px; color: #666; }
      .btn { padding: 8px 16px; border: none; border-radius: 6px; font-size: 13px; font-weight: bold; cursor: pointer; }
      .btn-download { background: #1a5c2e; color: #fff; }
      .info { font-size: 12px; color: #888; margin-top: 16px; line-height: 1.5; }
      #msg { margin-top: 8px; font-size: 13px; color: #28a745; }
    </style>
    <h3>📤 Exportar Efectivos — Curso ${curso}</h3>
    <div class="stats">
      <b>${datosProv.length}</b> provisionales · <b>${datosDef.length}</b> definitivos ·
      <b>${Math.round(json.length / 1024)} KB</b>
    </div>
    <div class="file-row">
      <code>${nombreArchivo}.json</code>
      <span class="size">${Math.round(json.length / 1024)} KB</span>
      <button class="btn btn-download" onclick="descargar()">⬇ Descargar</button>
    </div>
    <div id="msg"></div>
    <div class="info">
      <b>${nombreArchivo}.json</b> → <code>vercel-efectivos/data/</code> (datos privados, no se sirven al navegador).<br>
      <code>diccionarios.json</code> es el mismo que el del Concurso de Traslados; ya está en <code>vercel-efectivos/public/</code>.
    </div>
    <script>
      var json = ${JSON.stringify(json)};
      function descargar() {
        var blob = new Blob([json], {type: 'application/json'});
        var a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = '${nombreArchivo}.json';
        a.click();
        document.getElementById('msg').textContent = '✅ ${nombreArchivo}.json descargado';
      }
    </script>
  `).setWidth(520).setHeight(360);
  ui.showModalDialog(html, 'Exportar JSON Efectivos para Vercel');
}

// ===================== EXPORTAR DICCIONARIOS =====================

// Lee las pestañas "Especialidades" y "Centros" y devuelve el objeto diccionarios.
// Formato: { especialidades: {codigo: nombre}, centros: [{codigo,nombre,localidad,provincia[,lat,lon]}] }
// lat/lon (columnas P/Q de "Centros") se añaden solo si la fila trae coordenadas válidas.
function _construirDiccionarios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Especialidades: A = código, C = resumen
  const especialidades = {};
  const espSheet = ss.getSheetByName('Especialidades');
  if (espSheet && espSheet.getLastRow() > 1) {
    const espData = espSheet.getDataRange().getValues();
    for (let i = 1; i < espData.length; i++) {
      const codigo = String(espData[i][ESP_COL_CODIGO - 1] || '').trim();
      const nombre = String(espData[i][ESP_COL_NOMBRE - 1] || '').trim();
      if (codigo) especialidades[codigo] = nombre;
    }
  }

  // Centros: código, nombre, localidad, provincia, lat, lon
  const centros = [];
  const centrosSheet = ss.getSheetByName('Centros');
  if (centrosSheet && centrosSheet.getLastRow() > 1) {
    const cd = centrosSheet.getDataRange().getValues();
    for (let i = 1; i < cd.length; i++) {
      const codigo = String(cd[i][CENTROS_COL_CODIGO - 1] || '').trim();
      const nombre = String(cd[i][CENTROS_COL_NOMBRE - 1] || '').trim();
      const localidad = String(cd[i][CENTROS_COL_LOCALIDAD - 1] || '').trim();
      const provincia = String(cd[i][CENTROS_COL_PROVINCIA - 1] || '').trim();
      if (!codigo) continue;
      const c = { codigo: codigo, nombre: nombre, localidad: localidad, provincia: provincia };
      // Coordenadas (P/Q). Acepta número o texto con coma decimal. Solo se añaden si son válidas.
      const lat = _aNumeroCoord(cd[i][CENTROS_COL_LAT - 1]);
      const lon = _aNumeroCoord(cd[i][CENTROS_COL_LON - 1]);
      if (lat !== null && lon !== null && lat >= 30 && lat <= 45 && lon >= -10 && lon <= 5) {
        c.lat = lat;
        c.lon = lon;
      }
      centros.push(c);
    }
  }

  return { especialidades: especialidades, centros: centros };
}

// Convierte un valor de celda (número o "36,9608") a Number, o null si no es válido.
function _aNumeroCoord(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  const s = String(v).trim().replace(/\s/g, '').replace(',', '.');
  if (!s) return null;
  const n = parseFloat(s);
  return isFinite(n) ? n : null;
}

// Exporta diccionarios.json para subir a vercel-efectivos/public/
function exportarDiccionariosJSON() {
  const ui = SpreadsheetApp.getUi();
  const dicc = _construirDiccionarios();
  const nEsp = Object.keys(dicc.especialidades).length;
  const nCentros = dicc.centros.length;

  if (nEsp === 0 && nCentros === 0) {
    ui.alert('⚠️ No se encontraron datos en las pestañas "Especialidades" ni "Centros". ' +
      'Comprueba que existen y que las columnas coinciden (Esp: A código / C nombre; ' +
      'Centros: A código / E nombre / G localidad / K provincia / P lat / Q lon).');
    return;
  }

  const json = JSON.stringify(dicc);
  const nCoords = dicc.centros.filter(function (c) { return typeof c.lat === 'number' && typeof c.lon === 'number'; }).length;

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: -apple-system, sans-serif; padding: 20px; color: #333; }
      h3 { color: #1a5c2e; margin: 0 0 12px; }
      .stats { font-size: 13px; margin-bottom: 16px; }
      .stats b { color: #1a5c2e; }
      .file-row { display: flex; align-items: center; gap: 10px; margin: 10px 0; padding: 12px; background: #f5f7fa; border-radius: 8px; }
      .file-row code { font-size: 13px; font-weight: 600; flex: 1; }
      .file-row .size { font-size: 12px; color: #666; }
      .btn { padding: 8px 16px; border: none; border-radius: 6px; font-size: 13px; font-weight: bold; cursor: pointer; }
      .btn-download { background: #1a5c2e; color: #fff; }
      .info { font-size: 12px; color: #888; margin-top: 16px; line-height: 1.5; }
      #msg { margin-top: 8px; font-size: 13px; color: #28a745; }
    </style>
    <h3>📤 Exportar diccionarios.json</h3>
    <div class="stats"><b>${nEsp}</b> especialidades · <b>${nCentros}</b> centros (<b>${nCoords}</b> con coordenadas) · <b>${Math.round(json.length / 1024)} KB</b></div>
    <div class="file-row">
      <code>diccionarios.json</code>
      <span class="size">${Math.round(json.length / 1024)} KB</span>
      <button class="btn btn-download" onclick="descargar()">⬇ Descargar</button>
    </div>
    <div id="msg"></div>
    <div class="info">
      <b>diccionarios.json</b> → <code>vercel-efectivos/public/</code> (datos públicos: especialidades y centros).
    </div>
    <script>
      var json = ${JSON.stringify(json)};
      function descargar() {
        var blob = new Blob([json], {type: 'application/json'});
        var a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'diccionarios.json';
        a.click();
        document.getElementById('msg').textContent = '✅ diccionarios.json descargado';
      }
    </script>
  `).setWidth(520).setHeight(340);
  ui.showModalDialog(html, 'Exportar diccionarios.json para Vercel');
}

// ============================================================
//  IMPORTADOR DE CSV DESDE CARPETAS DE DRIVE
// ============================================================
//  Sueltas los CSV del bookmarklet "Efectivos → CSV" en una carpeta de Drive
//  y desde el menú se importan automáticamente, creando/actualizando una
//  pestaña por especialidad con el sufijo de fase:
//     · Carpeta provisional  →  pestañas "<código>_Prov"
//     · Carpeta definitiva    →  pestañas "<código>_Def"
//  Reimportar reprocesa TODA la carpeta (idempotente): cada CSV sobrescribe
//  su pestaña, así puedes ir soltando archivos poco a poco y reimportar.
//  El CSV ya viene en formato A–L (12 columnas), listo para consolidar.
// ============================================================

// Claves donde se guardan los IDs de las carpetas de Drive (por hoja).
const PROP_CARPETA_PROV = 'CARPETA_CSV_PROV';
const PROP_CARPETA_DEF  = 'CARPETA_CSV_DEF';

const IMPORT_CHARSET = 'UTF-8';   // el CSV del extractor se genera en UTF-8 (con BOM)
const IMPORT_SEP     = ';';       // separador del CSV del extractor

// ---------- Configurar carpetas ----------
function elegirCarpetaProv() { _pedirCarpeta(PROP_CARPETA_PROV, 'provisional'); }
function elegirCarpetaDef()  { _pedirCarpeta(PROP_CARPETA_DEF, 'definitiva'); }

function _idCarpetaGuardada(prop) {
  return PropertiesService.getDocumentProperties().getProperty(prop);
}

function _pedirCarpeta(prop, etiqueta) {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Carpeta de Drive – CSV ' + etiqueta,
    'Pega el ENLACE o el ID de la carpeta de Drive con los CSV ' + etiqueta + 's:',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const m = (resp.getResponseText() || '').match(/[-\w]{25,}/);  // ID de Drive: 25+ chars
  if (!m) { ui.alert('No he reconocido el ID de la carpeta. Revisa el enlace.'); return; }
  try {
    const c = DriveApp.getFolderById(m[0]);
    PropertiesService.getDocumentProperties().setProperty(prop, m[0]);
    ui.alert('✅ Carpeta ' + etiqueta + ' configurada:\n' + c.getName());
  } catch (e) {
    ui.alert('No puedo abrir esa carpeta. Comprueba el enlace y que tengas acceso.\n\n' + e.message);
  }
}

// ---------- Importar (núcleo sin interfaz, reutilizable por el trigger) ----------
// Devuelve { error } o { archivos, hojas, filas }.
function _importarCarpeta(prop, sufijo) {
  const id = _idCarpetaGuardada(prop);
  if (!id) return { error: 'sin_carpeta' };

  let carpeta;
  try { carpeta = DriveApp.getFolderById(id); }
  catch (e) { return { error: 'inaccesible', mensaje: e.message }; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const it = carpeta.getFiles();
  let nArch = 0, nHojas = 0, nFilas = 0;

  while (it.hasNext()) {
    const file = it.next();
    if (!/\.csv$/i.test(file.getName())) continue;

    let texto = file.getBlob().getDataAsString(IMPORT_CHARSET);
    if (texto.charCodeAt(0) === 0xFEFF) texto = texto.slice(1);   // quita el BOM inicial

    const tabla = Utilities.parseCsv(texto, IMPORT_SEP);
    if (!tabla || tabla.length < 2) continue;   // sin filas de datos
    nArch++;

    // Limpiar celdas (quita el apóstrofo inicial que fuerza texto en el CSV).
    const limpio = tabla.map(function (fila) { return fila.map(_limpiarCeldaCsv); });

    // Nombre de la pestaña = código de especialidad (columna A, 1ª fila de datos)
    // + sufijo de fase. Si no hay código, cae al nombre del archivo.
    let codigo = (limpio[1][0] || '').trim();
    if (!codigo) codigo = file.getName().replace(/\.csv$/i, '');
    const nombreHoja = _nombreHojaValido(codigo + sufijo);

    // Normalizar todas las filas a la misma anchura.
    let nCols = 1;
    limpio.forEach(function (f) { if (f.length > nCols) nCols = f.length; });
    const datos = limpio.map(function (f) {
      const c = f.slice();
      while (c.length < nCols) c.push('');
      return c;
    });

    // Crear/limpiar la pestaña y volcar como TEXTO (formato @) para no romper
    // NIFs, "00-00-00", notas con coma decimal ni códigos con ceros a la izq.
    let hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) hoja = ss.insertSheet(nombreHoja);
    hoja.clear();
    const rng = hoja.getRange(1, 1, datos.length, nCols);
    rng.setNumberFormat('@');
    rng.setValues(datos);
    hoja.setFrozenRows(1);

    nHojas++;
    nFilas += (datos.length - 1);
  }

  return { archivos: nArch, hojas: nHojas, filas: nFilas };
}

function _limpiarCeldaCsv(v) {
  v = (v == null ? '' : String(v)).trim();
  if (v.charAt(0) === "'") v = v.slice(1);
  return v;
}

// Un nombre de pestaña válido: sin caracteres prohibidos ni apóstrofos en los
// extremos, y recortado a un máximo razonable.
function _nombreHojaValido(n) {
  n = String(n).replace(/[\[\]\*\?\/\\:]/g, '_').replace(/^'+|'+$/g, '').trim().slice(0, 95);
  return n || 'SIN_COD';
}

// ---------- Acciones de menú ----------
function importarProv() { if (_avisoImport(_importarCarpeta(PROP_CARPETA_PROV, '_Prov'), 'provisional')) {} }
function importarDef()  { if (_avisoImport(_importarCarpeta(PROP_CARPETA_DEF, '_Def'), 'definitiva')) {} }

// Muestra el resultado de una importación y devuelve true si fue bien.
function _avisoImport(r, etiqueta) {
  const ui = SpreadsheetApp.getUi();
  if (r.error === 'sin_carpeta') {
    ui.alert('Primero configura la carpeta ' + etiqueta + ':\n' +
      '"📂 Colocación de Efectivos" → "Elegir carpeta CSV ' + etiqueta + '..."');
    return false;
  }
  if (r.error === 'inaccesible') {
    ui.alert('No puedo abrir la carpeta ' + etiqueta + '. Vuelve a elegirla.\n\n' + r.mensaje);
    return false;
  }
  if (r.archivos === 0) {
    ui.alert('No he encontrado ningún .csv en la carpeta ' + etiqueta + '.');
    return false;
  }
  ui.alert('✅ Importación ' + etiqueta + '\n\n' +
    'CSV leídos: ' + r.archivos + '\n' +
    'Pestañas creadas/actualizadas: ' + r.hojas + '\n' +
    'Filas de datos: ' + r.filas);
  return true;
}

// Combos "todo en uno": importa la carpeta y consolida en Global_*.
function importarYConsolidarProv() {
  const r = _importarCarpeta(PROP_CARPETA_PROV, '_Prov');
  if (!_avisoImportBreve(r, 'provisional')) return;
  const n = consolidarProv();
  SpreadsheetApp.getUi().alert('✅ Provisional lista\n\n' +
    'CSV importados: ' + r.archivos + '  ·  Pestañas *_Prov: ' + r.hojas + '\n' +
    'Filas en ' + EFECTIVOS_DEST_PROV + ': ' + n);
}

function importarYConsolidarDef() {
  const r = _importarCarpeta(PROP_CARPETA_DEF, '_Def');
  if (!_avisoImportBreve(r, 'definitiva')) return;
  const n = consolidarDef();
  SpreadsheetApp.getUi().alert('✅ Definitiva lista\n\n' +
    'CSV importados: ' + r.archivos + '  ·  Pestañas *_Def: ' + r.hojas + '\n' +
    'Filas en ' + EFECTIVOS_DEST_DEF + ': ' + n);
}

// Como _avisoImport pero solo avisa en caso de error (para los combos).
function _avisoImportBreve(r, etiqueta) {
  const ui = SpreadsheetApp.getUi();
  if (r.error === 'sin_carpeta') {
    ui.alert('Primero configura la carpeta ' + etiqueta + ' (menú → "Elegir carpeta CSV ' + etiqueta + '...").');
    return false;
  }
  if (r.error === 'inaccesible') { ui.alert('No puedo abrir la carpeta ' + etiqueta + '. Reconfigúrala.\n\n' + r.mensaje); return false; }
  if (r.archivos === 0) { ui.alert('No hay ningún .csv en la carpeta ' + etiqueta + '.'); return false; }
  return true;
}

// ---------- Auto-import programado (opcional) ----------
// Instala un disparador horario que importa ambas carpetas y reconsolida.
function activarAutoImport() {
  const ui = SpreadsheetApp.getUi();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'autoImportProgramado') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoImportProgramado').timeBased().everyHours(1).create();
  ui.alert('✅ Auto-import activado\n\n' +
    'Cada hora se importarán los CSV de las carpetas configuradas y se\n' +
    'reconsolidarán Global_Prov / Global_Def.\n\n' +
    '⚠️ No edites a mano las pestañas *_Prov / *_Def: se regeneran en cada import.');
}

function desactivarAutoImport() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'autoImportProgramado') { ScriptApp.deleteTrigger(t); n++; }
  });
  SpreadsheetApp.getUi().alert(n ? '⏹️ Auto-import desactivado.' : 'No había ningún auto-import activo.');
}

// Ejecutada por el disparador (SIN interfaz). Importa y consolida cada fase
// solo si su carpeta está configurada. Los errores se registran en el log.
function autoImportProgramado() {
  if (_idCarpetaGuardada(PROP_CARPETA_PROV)) {
    const rp = _importarCarpeta(PROP_CARPETA_PROV, '_Prov');
    if (!rp.error) consolidarProv();
    else console.warn('Auto-import provisional:', rp.error, rp.mensaje || '');
  }
  if (_idCarpetaGuardada(PROP_CARPETA_DEF)) {
    const rd = _importarCarpeta(PROP_CARPETA_DEF, '_Def');
    if (!rd.error) consolidarDef();
    else console.warn('Auto-import definitiva:', rd.error, rd.mensaje || '');
  }
}
