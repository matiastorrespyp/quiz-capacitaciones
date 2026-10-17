/**
 * QUIZ DE CAPACITACIONES — PyP Logística
 * Web App de Google Apps Script vinculada a una Google Sheet.
 *
 *  - Cada escaneo del QR sortea preguntas distintas (y mezcla las opciones).
 *  - Las respuestas correctas NUNCA viajan al celular: se corrigen en el servidor.
 *  - Todo queda en la planilla (Resultados + Detalle).
 *  - Panel del capacitador:  <URL>?admin=TU_CLAVE  → QR, ranking en vivo y respuestas correctas.
 *
 * Para sumar otra capacitación: agregá filas en la hoja "Preguntas" con otro nombre en la
 * columna "Capacitacion" y abrí el panel con &cap=Nombre de la capacitación (o cambiala en Config).
 */

const HOJA_PREG = 'Preguntas';
const HOJA_RES  = 'Resultados';
const HOJA_DET  = 'Detalle';
const HOJA_CFG  = 'Config';

/* ============================== WEB ============================== */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const cfg = getConfig_();
  const appUrl = ScriptApp.getService().getUrl();  // dentro de la web app sí es el link real
  guardarUrlApp_(appUrl);
  const capa = (p.cap || cfg.capacitacion || '').trim();  // 'c' está reservado por Apps Script

  // API JSON para la pantalla publicada en GitHub Pages
  if (p.api === 'info') return json_({ ok: true, data: { capa: capa, empresa: cfg.empresa || '' } });
  if (p.api === 'banco') {   // preguntas para otras planillas conectadas (protegido con clave)
    if (!cfg.claveBanco || String(p.k) !== String(cfg.claveBanco)) return json_({ ok: false, error: 'Clave de preguntas compartidas incorrecta.' });
    return json_({ ok: true, data: filasLocales_() });
  }
  if (p.api === 'cuestionario') {
    try { return json_({ ok: true, data: obtenerCuestionario(capa) }); }
    catch (err) { return json_({ ok: false, error: err.message }); }
  }

  if (p.admin !== undefined) {
    if (String(p.admin) !== String(cfg.clave)) {
      return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px">Clave incorrecta.</p>');
    }
    const t = HtmlService.createTemplateFromFile('Admin');
    t.clave = String(p.admin);
    t.capa = capa;
    const base = urlCuestionario_(cfg, appUrl);
    t.urlJuego = base + (p.cap ? (base.indexOf('?') < 0 ? '?' : '&') + 'c=' + encodeURIComponent(capa) : '');
    t.capacitaciones = JSON.stringify(listarCapacitaciones_());
    t.urlBase = urlCuestionario_(cfg, appUrl);
    t.capaPedida = p.cap ? 'SI' : '';
    return t.evaluate()
      .setTitle('Panel · ' + capa)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const t = HtmlService.createTemplateFromFile('Quiz');
  t.capa = capa;
  return t.evaluate()
    .setTitle(capa || 'Cuestionario')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Envío de respuestas desde la pantalla de GitHub Pages (body JSON como text/plain). */
function doPost(e) {
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    return json_({ ok: true, data: enviarRespuestas(b.token, b.nombre, b.respuestas) });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

/** Link que va en el QR: la pantalla pública de GitHub apuntando al motor de ESTA planilla. */
const PANTALLA_PUBLICA = 'https://matiastorrespyp.github.io/quiz-capacitaciones/';
function urlCuestionario_(cfg, app) {
  if (cfg.urlCuestionario) return cfg.urlCuestionario;
  app = app || '';
  const m = app.match(/\/s\/([A-Za-z0-9_-]+)\/(exec|dev)/);
  return m ? PANTALLA_PUBLICA + '?s=' + m[1] : app;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ===================== MENÚ EN LA PLANILLA ===================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Quiz')
    .addItem('Limpiar resultados', 'limpiarResultados')
    .addItem('Ver links del cuestionario y del panel', 'verLinks')
    .addItem('Conectar a preguntas compartidas…', 'conectarPreguntas')
    .addSeparator()
    .addItem('Preparar planilla (setup)', 'setup')
    .addToUi();
}

function limpiarResultados() {
  const ui = SpreadsheetApp.getUi();
  const capas = listarCapacitaciones_();
  const r = ui.prompt('Limpiar resultados',
    'Escribí el nombre de la capacitación a borrar (o TODAS). Las otras no se tocan.\n\nCapacitaciones: ' + capas.join('  ·  '),
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const elegida = r.getResponseText().trim();
  const todas = elegida.toUpperCase() === 'TODAS';
  const match = capas.find(c => c.toLowerCase() === elegida.toLowerCase());
  if (!todas && !match) { ui.alert('No encontré "' + elegida + '". No se borró nada.'); return; }
  const ss = SpreadsheetApp.getActive();
  let borrados = 0;
  [HOJA_RES, HOJA_DET].forEach(nombre => {
    const sh = ss.getSheetByName(nombre);
    if (!sh || sh.getLastRow() < 2) return;
    const rango = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn());
    const filas = rango.getValues();
    const quedan = filas.filter(f => !(todas || String(f[1]).trim() === match));
    if (nombre === HOJA_RES) borrados = filas.length - quedan.length;
    rango.clearContent();
    if (quedan.length) sh.getRange(2, 1, quedan.length, quedan[0].length).setValues(quedan);
  });
  ss.toast(borrados + ' participante(s) borrados de ' + (todas ? 'todas las capacitaciones' : match) + '.', 'Quiz', 6);
}

/** Guarda el link real de la implementación. Desde el menú, getUrl() devuelve un link que no funciona (falla de Google). */
function guardarUrlApp_(url) {
  if (!url || !/\/exec$/.test(url)) return;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('URL_APP') !== url) props.setProperty('URL_APP', url);
}

/** Conecta esta planilla a la planilla central: sus capacitaciones aparecen solas en el panel. */
function conectarPreguntas() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Conectar a preguntas compartidas',
    'Pegá el link de preguntas compartidas que te pasaron (empieza con https://script.google.com/macros/s/...).\n' +
    'Para desconectar, dejalo vacío.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const url = r.getResponseText().trim();
  if (url && !/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec\?/.test(url)) { ui.alert('Ese link no parece correcto. No se cambió nada.'); return; }
  let capas = [];
  if (url) {
    try {
      const j = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true }).getContentText());
      if (!j.ok) { ui.alert('La planilla central respondió: ' + j.error); return; }
      const set = {}; j.data.forEach(f => set[f[0]] = 1); capas = Object.keys(set);
    } catch (err) { ui.alert('No pude leer ese link: ' + err.message); return; }
  }
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA_CFG);
  const claves = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(x => String(x[0]).trim());
  const i = claves.indexOf('Preguntas compartidas desde');
  if (i >= 0) sh.getRange(i + 1, 2).setValue(url); else sh.appendRow(['Preguntas compartidas desde', url]);
  CacheService.getScriptCache().removeAll(['banco_remoto', 'banco_remoto_respaldo']);
  ui.alert(url ? 'Conectado. Capacitaciones disponibles: ' + capas.join(' · ') + '\n\nAparecen como botones en tu panel.' : 'Desconectado.');
}

function verLinks() {
  const cfg = getConfig_();
  const app = PropertiesService.getScriptProperties().getProperty('URL_APP');
  if (!app) {
    SpreadsheetApp.getUi().alert('Links del Quiz',
      'Todavía no se abrió nunca el link de la implementación.\n\n' +
      '1. Extensiones → Apps Script → Implementar → Administrar implementaciones → botón Copiar (debajo de la URL).\n' +
      '2. Pegalo en una ventana de incógnito y agregá al final: ?admin=' + cfg.clave + '\n\n' +
      'Ese es tu panel. Después de abrirlo una vez, este menú ya muestra los links.',
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const juego = urlCuestionario_(cfg, app);
  const html = '<div style="font-family:sans-serif;font-size:14px;line-height:1.5">' +
    '<p><b>Cuestionario (QR):</b><br><a href="' + juego + '" target="_blank">' + juego + '</a></p>' +
    '<p><b>Panel del capacitador:</b><br><a href="' + app + '?admin=' + encodeURIComponent(cfg.clave) + '" target="_blank">' +
    app + '?admin=' + cfg.clave + '</a></p></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(520).setHeight(220), 'Links del Quiz');
}

/* ===================== LLAMADAS DESDE EL CELULAR ===================== */

/** Sortea el cuestionario para un participante y lo guarda en caché con un token. */
function obtenerCuestionario(capa) {
  const cfg = getConfig_();
  const banco = leerBanco_(capa);
  if (!banco.length) throw new Error('No hay preguntas activas para "' + capa + '".');

  const total = Math.min(Number(cfg.preguntasPorPersona) || 6, banco.length);
  const capciosasPedidas = Number(cfg.capciosasPorPersona) || 0;

  const capciosas = mezclar_(banco.filter(q => q.capciosa));
  const normales  = mezclar_(banco.filter(q => !q.capciosa));
  let elegidas = capciosas.slice(0, Math.min(capciosasPedidas, total));
  elegidas = elegidas.concat(normales.slice(0, total - elegidas.length));
  if (elegidas.length < total) {           // por si faltan normales
    const resto = capciosas.slice(elegidas.filter(q => q.capciosa).length);
    elegidas = elegidas.concat(resto.slice(0, total - elegidas.length));
  }
  elegidas = mezclar_(elegidas);

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(token, JSON.stringify({
    capa: capa,
    ids: elegidas.map(q => q.id),
    inicio: Date.now()
  }), 21600); // 6 horas

  return {
    token: token,
    capa: capa,
    preguntas: elegidas.map(q => ({
      id: q.id,
      texto: q.pregunta,
      opciones: mezclar_([q.correcta].concat(q.incorrectas))
    }))
  };
}

/** Recibe las respuestas, corrige y guarda. Devuelve solo el puntaje (no las correctas). */
function enviarRespuestas(token, nombre, respuestas) {
  nombre = String(nombre || '').trim().replace(/\s+/g, ' ').slice(0, 60);
  if (!nombre) throw new Error('Falta el nombre.');

  const cache = CacheService.getScriptCache();
  const raw = cache.get(token);
  if (!raw) throw new Error('Este cuestionario ya fue enviado o venció. Escaneá el QR de nuevo.');
  const sesion = JSON.parse(raw);

  const banco = leerBanco_(sesion.capa);
  const porId = {};
  banco.forEach(q => porId[q.id] = q);

  const ahora = new Date();
  const segundos = Math.round((Date.now() - sesion.inicio) / 1000);
  let aciertos = 0;
  const filasDet = [];

  sesion.ids.forEach(id => {
    const q = porId[id];
    if (!q) return;
    const elegida = String((respuestas || {})[id] || '');
    const ok = norm_(elegida) === norm_(q.correcta);
    if (ok) aciertos++;
    filasDet.push([ahora, sesion.capa, nombre, id, q.pregunta, elegida, q.correcta, ok ? 'SI' : 'NO']);
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!cache.get(token)) throw new Error('Este cuestionario ya fue enviado.');
    cache.remove(token);
    const ss = SpreadsheetApp.getActive();
    ss.getSheetByName(HOJA_RES).appendRow([ahora, sesion.capa, nombre, aciertos, sesion.ids.length, segundos, token]);
    const det = ss.getSheetByName(HOJA_DET);
    if (filasDet.length) det.getRange(det.getLastRow() + 1, 1, filasDet.length, filasDet[0].length).setValues(filasDet);
  } finally {
    lock.releaseLock();
  }

  return { aciertos: aciertos, total: sesion.ids.length, segundos: segundos };
}

/* ======================= LLAMADAS DEL PANEL ======================= */

function datosPanel(clave, capa) {
  const cfg = getConfig_();
  if (String(clave) !== String(cfg.clave)) throw new Error('Clave incorrecta.');
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();

  const res = valores_(ss.getSheetByName(HOJA_RES))
    .filter(r => r[1] === capa)
    .map(r => ({
      fecha: Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd'),
      hora: Utilities.formatDate(new Date(r[0]), tz, 'HH:mm'),
      nombre: r[2], aciertos: Number(r[3]), total: Number(r[4]), segundos: Number(r[5])
    }));

  const det = valores_(ss.getSheetByName(HOJA_DET))
    .filter(r => r[1] === capa)
    .map(r => ({ fecha: Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd'), id: String(r[3]), ok: r[7] === 'SI' }));

  const preguntas = leerBanco_(capa).map(q => ({
    id: q.id, pregunta: q.pregunta, correcta: q.correcta,
    incorrectas: q.incorrectas, capciosa: q.capciosa, explicacion: q.explicacion
  }));

  return { resultados: res, detalle: det, preguntas: preguntas,
           hoy: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd') };
}

/* ============================ HELPERS ============================ */

function getConfig_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA_CFG);
  const cfg = {};
  if (!sh) return cfg;
  sh.getDataRange().getValues().forEach(r => {
    const k = String(r[0] || '').trim();
    if (!k) return;
    const map = {
      'Capacitacion activa': 'capacitacion',
      'Preguntas por persona': 'preguntasPorPersona',
      'Capciosas por persona': 'capciosasPorPersona',
      'Clave del panel': 'clave',
      'URL del cuestionario': 'urlCuestionario',
      'Nombre de la empresa': 'empresa',
      'Clave para compartir preguntas': 'claveBanco',
      'Preguntas compartidas desde': 'fuente'
    };
    if (map[k]) cfg[map[k]] = String(r[1]).trim();
  });
  return cfg;
}

/* ---- Preguntas: las de esta planilla + las compartidas desde la planilla central ---- */

function filasLocales_() {
  return valores_(SpreadsheetApp.getActive().getSheetByName(HOJA_PREG))
    .filter(r => String(r[0]).trim())
    .map(r => r.slice(0, 10).map(v => v instanceof Date ? v.toISOString() : v));
}

/** Filas de la planilla central (si esta planilla está conectada). Se guardan 5 min en caché. */
function filasRemotas_(cfg) {
  cfg = cfg || getConfig_();
  if (!cfg.fuente) return [];
  const cache = CacheService.getScriptCache();
  const enCache = cache.get('banco_remoto');
  if (enCache) return JSON.parse(enCache);
  try {
    const r = UrlFetchApp.fetch(cfg.fuente, { muteHttpExceptions: true, followRedirects: true });
    const j = JSON.parse(r.getContentText());
    if (!j.ok) throw new Error(j.error);
    const txt = JSON.stringify(j.data);
    if (txt.length < 95000) { cache.put('banco_remoto', txt, 300); cache.put('banco_remoto_respaldo', txt, 21600); }
    return j.data;
  } catch (err) {
    const resp = cache.get('banco_remoto_respaldo');   // si la central no responde, usa la última copia
    return resp ? JSON.parse(resp) : [];
  }
}

/** Todas las filas: si una capacitación existe en la central, se usa la versión central. */
function todasLasFilas_() {
  const remotas = filasRemotas_();
  const capasRemotas = {};
  remotas.forEach(r => capasRemotas[String(r[0]).trim()] = 1);
  return remotas.concat(filasLocales_().filter(r => !capasRemotas[String(r[0]).trim()]));
}

function leerBanco_(capa) {
  return todasLasFilas_()
    .filter(r => String(r[0]).trim() === capa && String(r[9]).toUpperCase() !== 'NO' && r[2] && r[3])
    .map(r => ({
      id: String(r[1]),
      pregunta: String(r[2]),
      correcta: String(r[3]),
      incorrectas: [r[4], r[5], r[6]].filter(x => String(x).trim() !== '').map(String),
      capciosa: String(r[7]).toUpperCase() === 'SI',
      explicacion: String(r[8] || '')
    }));
}

function listarCapacitaciones_() {
  const set = {};
  todasLasFilas_().forEach(r => { if (r[0]) set[String(r[0]).trim()] = 1; });
  return Object.keys(set);
}

function valores_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}

function mezclar_(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function norm_(s) { return String(s).trim().toLowerCase(); }

function include(nombre) { return HtmlService.createHtmlOutputFromFile(nombre).getContent(); }

/* ======================= INSTALACIÓN (correr 1 vez) ======================= */

function setup() {
  const ss = SpreadsheetApp.getActive();

  const crear = (nombre, encabezados) => {
    let sh = ss.getSheetByName(nombre);
    if (!sh) sh = ss.insertSheet(nombre);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, encabezados.length).setValues([encabezados])
        .setFontWeight('bold').setBackground('#0b5d3b').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    return sh;
  };

  const cfg = crear(HOJA_CFG, ['Parametro', 'Valor']);
  if (cfg.getLastRow() < 2) {
    const clave = Math.random().toString(36).slice(2, 8);
    cfg.getRange(2, 1, 4, 2).setValues([
      ['Capacitacion activa', 'Pasos para la venta'],
      ['Preguntas por persona', 6],
      ['Capciosas por persona', 2],
      ['Clave del panel', clave]
    ]);
    cfg.autoResizeColumns(1, 2);
  }
  const claves = cfg.getRange(1, 1, cfg.getLastRow(), 1).getValues().map(r => String(r[0]).trim());
  if (claves.indexOf('URL del cuestionario') < 0) cfg.appendRow(['URL del cuestionario', '']);
  if (claves.indexOf('Nombre de la empresa') < 0) cfg.appendRow(['Nombre de la empresa', '']);
  crearLeeme_(ss);

  const preg = crear(HOJA_PREG, ['Capacitacion', 'ID', 'Pregunta', 'Correcta', 'Incorrecta 1', 'Incorrecta 2', 'Incorrecta 3', 'Capciosa (SI/NO)', 'Explicacion', 'Activa (SI/NO)']);
  if (preg.getLastRow() < 2) {
    const filas = BANCO_PASOS_VENTA.map((q, i) =>
      ['Pasos para la venta', 'PV' + String(i + 1).padStart(2, '0'), q[0], q[1], q[2], q[3], q[4], q[5] ? 'SI' : 'NO', q[6], 'SI']);
    preg.getRange(2, 1, filas.length, filas[0].length).setValues(filas);
    preg.setColumnWidth(3, 380); preg.setColumnWidths(4, 4, 240); preg.setColumnWidth(9, 320);
    preg.getRange(2, 1, filas.length, 10).setWrap(true).setVerticalAlignment('top');
  }

  crear(HOJA_RES, ['Fecha', 'Capacitacion', 'Nombre', 'Aciertos', 'Total', 'Segundos', 'Token']);
  crear(HOJA_DET, ['Fecha', 'Capacitacion', 'Nombre', 'ID pregunta', 'Pregunta', 'Respuesta elegida', 'Respuesta correcta', 'Correcta']);

  const hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && hoja1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(hoja1);

  Logger.log('Listo. Clave del panel: ' + getConfig_().clave);
}

function crearLeeme_(ss) {
  if (ss.getSheetByName('LEEME')) return;
  const sh = ss.insertSheet('LEEME', 0);
  const filas = [
    ['QUIZ DE CAPACITACIONES POR QR — cómo ponerlo en marcha (una sola vez, 5 minutos)'],
    [''],
    ['1. Menú Quiz → "Preparar planilla (setup)". Google pide autorización: elegí tu cuenta → Configuración avanzada → Ir a ... (no seguro) → Permitir. Es tu propia copia, por eso aparece ese aviso.'],
    ['2. Hoja Config, columna B: en "Nombre de la empresa" escribí el de tu distribuidora, y en "Clave del panel" borrá "cambiar" e inventá una clave tuya (protege tu panel con las respuestas). "URL del cuestionario" dejala vacía.'],
    ['3. Menú Extensiones → Apps Script → botón azul Implementar → Nueva implementación → engranaje → "App web".'],
    ['    Ejecutar como: Yo   ·   Quién tiene acceso: Cualquiera   →   Implementar.'],
    ['4. En la ventana que aparece al implementar, tocá "Copiar" (debajo de la URL). Abrí ese link en una ventana de incógnito agregándole al final ?admin=TU_CLAVE  →  ese es tu PANEL: QR para tus vendedores, ranking en vivo y respuestas correctas. Guardalo en favoritos (después también aparece en Quiz → Ver links).'],
    [''],
    ['USO DIARIO'],
    ['• Preguntas: hoja Preguntas. Columna Capciosa = SI marca las "trampa". Activa = NO la saca del sorteo.'],
    ['• Config: cuántas preguntas recibe cada persona y cuántas capciosas.'],
    ['• Otra capacitación: cargá filas con otro nombre en la columna Capacitacion y ponelo en Config → Capacitacion activa.'],
    ['• Antes de cada capacitación: menú Quiz → "Limpiar resultados" para arrancar el ranking de cero.'],
    ['• Si cambiás el código de Apps Script: Implementar → Administrar implementaciones → lápiz → Versión nueva (el link y el QR no cambian).'],
    [''],
    ['No hace falta tocar GitHub ni ningún servidor: el QR del panel ya apunta a la pantalla pública conectada a ESTA planilla.']
  ];
  sh.getRange(1, 1, filas.length, 1).setValues(filas).setWrap(true).setVerticalAlignment('top');
  sh.setColumnWidth(1, 900);
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(13).setBackground('#0b5d3b').setFontColor('#ffffff');
  sh.getRange(9, 1).setFontWeight('bold');
}

/* ===== Banco inicial: "Pasos para la venta" =====
   [pregunta, correcta, incorrecta1, incorrecta2, incorrecta3, capciosa, explicación] */
const BANCO_PASOS_VENTA = [
  ['¿Cuál de estas tareas se hace ANTES de salir a ruta?',
   'Verificar que cada PDV a visitar haya recibido su pedido anterior',
   'Realizar el check in en el sistema', 'Acomodar la mercadería según el planograma', 'Cargar los motivos de No Venta', false,
   'Es el paso 3 de "Antes de salir a ruta". El check in y el acomodo se hacen en ruta; los motivos de No Venta, al final del día.'],
  ['La ruta del día se planifica en…',
   'Volumen e IP', 'Solo volumen', 'Cantidad de clientes a visitar', 'Kilómetros a recorrer', false,
   'Paso 2: planificar la ruta del día en volumen e IP.'],
  ['Antes de salir a ruta, ¿de qué hay que informarse?',
   'Objetivos, avance de ventas, presupuesto de combos y focos de la jornada',
   'Solo los objetivos del mes', 'Los precios de la competencia', 'La ruta del supervisor', false,
   'Paso 1: objetivos, avance de ventas, presupuesto de combos y focos para la jornada.'],
  ['¿Qué productos foco hay que conocer antes de salir a ruta?',
   'Los de frescura e innovación', 'Los de mayor precio', 'Los que están en liquidación', 'Los que el cliente pidió la vez anterior', false,
   'Paso 4: conocer faltantes de stock y productos foco a comercializar (frescura / innovación).'],
  ['Al llegar al PDV, ¿qué es lo PRIMERO que se hace?',
   'Realizar el check in en el sistema', 'Saludar al cliente por su nombre', 'Revisar si hay producto vencido', 'Tomar el pedido', true,
   'El paso 5 (check in) va antes que el 6 (saludar). Muchos responden "saludar" por costumbre.'],
  ['Sobre el saludo al cliente, el paso dice que hay que…',
   'Conocer al cliente por su nombre y que el cliente conozca el tuyo',
   'Saludar cordialmente, sin necesidad de nombres', 'Presentarse siempre con el nombre de la empresa', 'Saludar solo al dueño del comercio', true,
   'Paso 6: conocer al cliente por su nombre "y viceversa". La relación va en los dos sentidos.'],
  ['Si encontrás producto vencido o próximo a vencer en el PDV, ¿qué hacés?',
   'Verificarlo y gestionar su reemplazo', 'Retirarlo sin avisar al cliente', 'Ofrecerle al cliente un descuento para que lo venda', 'Informarlo recién en la reunión vespertina', false,
   'Paso 7: verificar la existencia de producto vencido o próximo a vencer y gestionar su reemplazo.'],
  ['La ejecución del punto de venta se trabaja de acuerdo con…',
   'La foto de éxito', 'Lo que pida el cliente', 'El censo anual de exhibidores', 'El criterio de cada vendedor', false,
   'Paso 7: trabajar la ejecución del PDV de acuerdo con la foto de éxito.'],
  ['Con el POP liviano, la tarea del vendedor es…',
   'Garantizar que esté vigente y en buenas condiciones', 'Cambiarlo por uno nuevo en cada visita', 'Revisarlo solo en el censo anual', 'Colocarlo únicamente si el cliente lo pide', false,
   'Paso 8: garantizar que el POP liviano se encuentre vigente y en buenas condiciones.'],
  ['Al acomodar la mercadería hay que respetar…',
   'El planograma y la rotación de los productos', 'Solo el planograma', 'Solo la rotación de los productos', 'El orden que prefiera el cliente', true,
   'Paso 9: planograma Y rotación. Las opciones de "solo uno de los dos" están incompletas.'],
  ['¿Qué información de la competencia hay que obtener?',
   'Ofertas, productos y promociones', 'Solo los precios', 'Solo el stock que tienen en el PDV', 'Los nombres de sus vendedores', false,
   'Paso: obtener la información de la competencia (ofertas, productos, promociones).'],
  ['La práctica de precios y el lugar de los exhibidores se orientan según…',
   'La foto de éxito', 'El censo de precios del mes', 'Los precios de la competencia', 'Lo que decida el cliente', true,
   'Paso 11: orientar precios y lugar de exhibidores de acuerdo con la foto de éxito, no con el censo.'],
  ['¿Qué dice el proceso sobre las exhibiciones adicionales?',
   'Identificar y negociar zonas oportunas para colocarlas', 'Colocarlas donde haya espacio libre', 'Solo se colocan durante el censo de exhibidores', 'No forman parte de la tarea del vendedor', false,
   'Paso 12: identificar y negociar zonas oportunas para colocar exhibiciones adicionales.'],
  ['Al efectuar el cierre de la venta, además hay que…',
   'Comunicarle al PDV el importe del pedido', 'Cobrar el pedido en el momento', 'Hacer el check out antes de cerrar la venta', 'Mandarle una foto del pedido al supervisor', true,
   'Paso 13: cerrar la venta, tomar recaudos para facilitar la entrega y comunicar el importe. El check out viene después.'],
  ['¿Cuál es el último paso dentro del PDV?',
   'Realizar el check out', 'Realizar el pedido', 'Comunicar el importe del pedido', 'Saludar al cliente', false,
   'Paso 14: realizar el check out del PDV.'],
  ['En la reunión vespertina, el vendedor informa…',
   'El cierre de venta en volumen e IP', 'Solo el volumen vendido', 'Solo los clientes que no pudo visitar', 'Los precios relevados de la competencia', false,
   'Paso 15: participar de la reunión vespertina informando el cierre de venta (volumen e IP).'],
  ['¿Con qué frecuencia es la reunión vespertina?',
   'Diaria, al terminar el programa de visitas', 'Semanal', 'Mensual', 'Solo cuando no se cumple el objetivo', true,
   'La vespertina es diaria. La que es mensual es la reunión de ventas (tarea general).'],
  ['¿Qué hay que validar con los clientes donde no se logró una visita exitosa?',
   'Haberles cargado el motivo de "No Venta"', 'Volver a visitarlos el mismo día', 'Avisarles por WhatsApp', 'Nada, se retoman en la próxima visita', false,
   'Paso 16: validar haber cargado motivos de "No Venta".'],
  ['¿Cada cuánto se completa la encuesta de Tienda Perfecta?',
   'Una vez al mes', 'Una vez por semana', 'Una vez al año', 'En cada visita', false,
   'Tareas generales: una vez al mes completar la encuesta de Tienda Perfecta.'],
  ['¿Cada cuánto se hace el censo de precios?',
   'Mensual', 'Anual', 'Semanal', 'Trimestral', true,
   'El censo de precios es mensual. El que es anual es el censo de exhibidores y tipología.'],
  ['¿Cada cuánto se hace el censo de exhibidores y tipología?',
   'Anual', 'Mensual', 'Semestral', 'En cada visita', true,
   'Es anual. Se confunde con el censo de precios, que es mensual.'],
  ['¿Cuál de estas NO es una tarea en ruta (dentro del PDV)?',
   'Planificar la ruta del día', 'Realizar el check in', 'Negociar zonas para exhibiciones adicionales', 'Obtener información de la competencia', true,
   'Planificar la ruta se hace antes de salir. Las otras tres se hacen dentro del PDV.'],
  ['¿Cuál de estas es una tarea general (no diaria)?',
   'Participar de la reunión mensual de ventas', 'Participar de la reunión vespertina', 'Realizar el check out', 'Cargar los motivos de No Venta', false,
   'Tareas generales: reunión mensual de ventas, encuesta Tienda Perfecta, censos y carpeta de ventas.'],
  ['Además de las reuniones y censos, ¿qué otra tarea general tiene el vendedor?',
   'Actualizar su carpeta de ventas y materiales de trabajo', 'Armar el planograma de cada PDV', 'Definir los precios de venta al público', 'Planificar las rutas de todo el equipo', false,
   'Tareas generales: actualizar la información de su carpeta de ventas y materiales de trabajo.']
];
