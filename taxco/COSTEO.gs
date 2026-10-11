/**
 * COSTEO TAXCO + BIBLIAS — v2.2
 * =============================
 * v2.2: botón "Subir Saldos de bodega" en la pestaña Costeo (subirSaldosBodega): el archivo se
 *       guarda en la carpeta con la fecha de hoy y los costos se actualizan de inmediato.
 * Costeo de recetas, margen (costo vs. venta) y biblias de cocina y de barra
 * (ficha técnica de cada plato y trago, con formulario para llenarla desde el iPad).
 * Archivo NUEVO dentro del proyecto VENTAS TAXCO. No toca Código.gs.
 * Usa SHEET_ID, ZONA y HOJA_PRODUCTOS que ya están definidos en Código.gs.
 *
 * CÓMO FUNCIONA
 *  - En Drive existe la carpeta "COSTEO TAXCO". Ahí se suben los reportes del MIG:
 *      Recetas*.xls       -> solo cuando cambie una receta
 *      Productos*.xlsx    -> solo cuando cambie el menú o un precio
 *      SaldosBodega*.xlsx -> cada mes (de aquí salen los costos)
 *  - cargarCosteoDrive() lee los archivos nuevos y los mueve a PROCESADOS.
 *  - Cada carga de Saldos guarda una foto de costos con su fecha (HIST_COSTOS).
 *  - Las pestañas Biblia cocina y Biblia barra del tablero guardan cada ficha
 *    en la pestaña FICHAS (y las fotos en la subcarpeta FOTOS FICHAS de Drive).
 *
 * PESTAÑAS DE LA HOJA
 *  INSUMOS          costo actual de cada insumo (de Saldos de bodega)
 *  HIST_COSTOS      historial de costos por fecha (FECHA, ID, COSTO)
 *  MAESTRO          productos con precio, impoconsumo y estado (de Productos)
 *  RECETAS          ingredientes de cada producto (de Recetas)
 *  COSTOS_MANUALES  costo de insumos que el MIG no trae (se llena a mano)
 *  FICHAS           biblias: se llena desde el formulario de la web
 *
 * ACTUALIZAR DE v1 A v2 (en este orden)
 *  1. Costeo.gs: borra todo el contenido y pega este archivo completo. Guarda.
 *  2. Costeo.html: borra todo el contenido y pega el nuevo. Guarda.
 *  3. Archivos (+) > HTML > "Biblia" (exactamente así): pega Biblia.html. Guarda.
 *  4. Ejecuta "configurarBiblia" (prepara la pestaña FICHAS y la carpeta de fotos).
 *  5. Implementar > Gestionar implementaciones > lápiz > Nueva versión > Implementar.
 */

const COSTEO_CARPETA = 'COSTEO TAXCO';
const COSTEO_FC_VERDE = 0.30;   // food cost hasta 30%  -> verde
const COSTEO_FC_AMBAR = 0.38;   // hasta 38% -> ámbar; más -> rojo
const COSTEO_SEMANAS = 13;      // ventana de rotación (≈ 3 meses)
const COSTEO_PRECIO_MIN = 1500; // precios menores se consideran de relleno (ej. adiciones a $100) y no entran en márgenes

const H_INSUMOS = 'INSUMOS';
const H_HIST = 'HIST_COSTOS';
const H_MAESTRO = 'MAESTRO';
const H_RECETAS = 'RECETAS';
const H_MANUALES = 'COSTOS_MANUALES';
const H_FICHAS = 'FICHAS';

const ENC_INSUMOS = ['ID', 'FAMILIA', 'INSUMO', 'SALDO', 'COSTO', 'FECHA'];
const ENC_HIST = ['FECHA', 'ID', 'COSTO'];
const ENC_MAESTRO = ['ID', 'PRODUCTO', 'ID_GRUPO', 'ACTIVO', 'PRECIO', 'IMPOCONSUMO', 'COSTO_MIG'];
const ENC_RECETAS = ['ID_PRODUCTO', 'PRODUCTO', 'GRUPO', 'ID_INSUMO', 'INSUMO', 'CANTIDAD', 'UNIDAD'];
const ENC_MANUALES = ['ID', 'NOMBRE', 'COSTO_LOTE', 'RENDIMIENTO', 'COSTO_UNIT'];
const ENC_FICHAS = ['ID', 'PRODUCTO', 'AREA', 'PORCIONES', 'PREPARACION', 'EMPLATADO', 'ALERGENOS', 'MERMA', 'VIDA_UTIL',
                    'CRISTALERIA', 'METODO', 'HIELO', 'GUARNICION', 'NOTAS', 'FOTO_ID', 'ACTUALIZADO',
                    'ORIGEN', 'TIPO', 'GRADOS', 'NOTAS_CATA', 'SERVICIO', 'MARIDAJE'];
// Campos que cuentan para que una ficha esté "completa" (la foto y las notas son opcionales)
const FICHA_REQ = {
  Cocina: ['porciones', 'preparacion', 'emplatado', 'alergenos'],
  Barra: ['cristaleria', 'metodo', 'hielo', 'preparacion', 'alergenos'],
  Botella: ['origen', 'tipo', 'notacata', 'servicio']   // licores, vinos y cervezas (productos de un solo ingrediente)
};


/* ======================================================================
 * 0. PÁGINA WEB (reemplaza a la doGet de Dashboard.gs)
 * ====================================================================== */
/** Sirve el tablero de siempre (Index) y le agrega las pestañas Costeo, Biblias y Compras al final de la página. */
function doGet() {
  let html = HtmlService.createHtmlOutputFromFile('Index').getContent();
  let extra = HtmlService.createHtmlOutputFromFile('Costeo').getContent();
  try { extra += HtmlService.createHtmlOutputFromFile('Biblia').getContent(); } catch (e) { /* si falta Biblia.html, el tablero sigue funcionando */ }
  try { extra += HtmlService.createHtmlOutputFromFile('Compras').getContent(); } catch (e) { /* si falta Compras.html, el tablero sigue funcionando */ }
  const i = html.lastIndexOf('</body>');
  html = i >= 0 ? html.slice(0, i) + extra + html.slice(i) : html + extra;
  return HtmlService.createHtmlOutput(html)
    .setTitle('Taxco Ventas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/* ======================================================================
 * 1. CONFIGURAR (ejecutar una vez)
 * ====================================================================== */
function configurarCosteo() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  cHoja(ss, H_INSUMOS, ENC_INSUMOS);
  cHoja(ss, H_HIST, ENC_HIST);
  cHoja(ss, H_MAESTRO, ENC_MAESTRO);
  cHoja(ss, H_RECETAS, ENC_RECETAS);
  cHoja(ss, H_FICHAS, ENC_FICHAS);

  const m = cHoja(ss, H_MANUALES, ENC_MANUALES);
  m.getRange('A1').setNote('ID del insumo en el MIG (columna ID_INSUMO de RECETAS).');
  m.getRange('C1').setNote('Costo del lote o del paquete completo.');
  m.getRange('D1').setNote('Cuánto rinde el lote (en la misma unidad de la receta: g, ml o und). Si lo dejas vacío, COSTO_LOTE se toma como costo por unidad.');
  m.getRange('E1').setNote('Se calcula solo: COSTO_LOTE ÷ RENDIMIENTO.');
  const formulas = [];
  for (let i = 2; i <= 301; i++) formulas.push(['=IF(A' + i + '="","",IF(D' + i + '="",C' + i + ',C' + i + '/D' + i + '))']);
  m.getRange(2, 5, 300, 1).setFormulas(formulas);
  m.getRange(2, 1, 300, 1).setNumberFormat('@');

  const carpeta = cCarpeta();
  cProcesados(carpeta);
  Logger.log('✅ Pestañas listas. Sube los reportes del MIG a esta carpeta de Drive: ' + carpeta.getUrl());
}

/** Prepara la pestaña FICHAS (v2) y la carpeta de fotos. Ejecutar una vez al pasar a la v2. */
function configurarBiblia() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  cHoja(ss, H_FICHAS, ENC_FICHAS);
  const fotos = cFotos();
  Logger.log('✅ Biblias listas. Las fotos se guardan en: ' + fotos.getUrl());
}

function crearDisparadorCosteo() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'cargarCosteoDrive') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('cargarCosteoDrive').timeBased().everyDays(1).atHour(7).inTimezone(ZONA).create();
  Logger.log('✅ Disparador creado: revisa la carpeta COSTEO TAXCO todos los días a las 7 AM');
}


/* ======================================================================
 * 2. CARGA DESDE DRIVE
 * ====================================================================== */

/** Lee los reportes nuevos de la carpeta y actualiza las pestañas. También la llama el botón del tablero. */
function cargarCosteoDrive() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false, mensajes: ['Ya hay una carga en curso. Espera un minuto e inténtalo de nuevo.'] };
  const mensajes = [];
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const carpeta = cCarpeta();
    const procesados = cProcesados(carpeta);
    const grupos = { recetas: [], productos: [], saldos: [] };

    const it = carpeta.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      const n = cNorm(f.getName());
      if (/^saldos/.test(n)) grupos.saldos.push(f);
      else if (/^recetas/.test(n)) grupos.recetas.push(f);
      else if (/^productos/.test(n)) grupos.productos.push(f);
    }

    const proc = [
      ['recetas', 'Recetas', cProcRecetas],
      ['productos', 'Productos', cProcProductos],
      ['saldos', 'Saldos de bodega', cProcSaldos]
    ];
    proc.forEach(function (p) {
      const lista = grupos[p[0]];
      if (!lista.length) return;
      lista.sort(function (a, b) { return b.getLastUpdated() - a.getLastUpdated(); });
      try {
        mensajes.push('✅ ' + p[1] + ': ' + p[2](ss, lista[0]));
        lista.forEach(function (f) { f.moveTo(procesados); });
      } catch (e) {
        mensajes.push('⚠️ ' + p[1] + ' (' + lista[0].getName() + '): ' + e.message);
      }
    });

    if (!mensajes.length) mensajes.push('No hay archivos nuevos en la carpeta COSTEO TAXCO. Suban Recetas, Productos o SaldosBodega del MIG.');
    return { ok: true, mensajes: mensajes };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Recibe desde la página el reporte Saldos de bodega del MIG (.xls o .xlsx, en base64),
 * lo guarda en la carpeta COSTEO TAXCO con la fecha de hoy en el nombre (así la foto de costos
 * queda al día de hoy) y lo procesa enseguida con cargarCosteoDrive.
 * Si el archivo no es un reporte de Saldos válido, se quita de la carpeta para no trabar otras cargas.
 */
function subirSaldosBodega(nombre, base64) {
  nombre = String(nombre || '');
  const ext = (nombre.match(/\.xlsx?$/i) || [''])[0].toLowerCase();
  if (!ext) throw new Error('El archivo debe ser Excel (.xls o .xlsx), como lo exporta el MIG.');
  if (!base64) throw new Error('No llegó el archivo. Intenta de nuevo.');
  if (String(base64).length > 30000000) throw new Error('El archivo es muy pesado (más de 20 MB).');
  const tipo = ext === '.xlsx' ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/vnd.ms-excel';
  const carpeta = cCarpeta();
  const archivo = carpeta.createFile(Utilities.newBlob(Utilities.base64Decode(base64), tipo, 'SaldosBodega ' + cHoy() + ext));
  const r = cargarCosteoDrive();
  // Si quedó en la carpeta (no pasó a PROCESADOS) es porque no se pudo procesar
  const sigue = r.ok && archivo.getParents().hasNext() && archivo.getParents().next().getId() === carpeta.getId();
  if (sigue) {
    archivo.setTrashed(true);
    r.mensajes.push('El archivo subido no se pudo usar y se quitó de la carpeta. Revisa que sea el reporte Saldos de bodega del MIG.');
  }
  r.subido = !sigue && r.ok;
  return r;
}

function cProcRecetas(ss, file) {
  const t = cTabla(cLeerArchivo(file), 'Grupo');
  const filas = [];
  t.forEach(function (r) {
    const idp = cId(r['id producto']), ins = cId(r['id receta']);
    if (!idp || !ins) return;
    filas.push([idp, cTxt(r['producto']), cTxt(r['grupo']), ins, cTxt(r['receta']), cNum(r['cantidad']), cTxt(r['id unidad'])]);
  });
  if (filas.length < 10) throw new Error('No parece el reporte de Recetas (encabezados: Grupo, Id Producto, Producto, Id Receta, Receta, Cantidad, Id Unidad).');
  cEscribir(cHoja(ss, H_RECETAS, ENC_RECETAS), ENC_RECETAS.length, filas, [1, 4]);
  const n = Object.keys(filas.reduce(function (a, f) { a[f[0]] = 1; return a; }, {})).length;
  PropertiesService.getScriptProperties().setProperty('COSTEO_RECETAS', cHoy());
  return n + ' productos con receta (' + filas.length + ' líneas).';
}

function cProcProductos(ss, file) {
  const t = cTabla(cLeerArchivo(file), 'Id Referencia');
  const filas = [];
  t.forEach(function (r) {
    const id = cId(r['id referencia']);
    if (!id) return;
    const est = r['estado'] === true || cNorm(r['estado']) === 'true' || +r['estado'] === 1;
    filas.push([id, cTxt(r['referencia']), cId(r['id grupo']), est ? 1 : 0, cNum(r['precio venta1']), cNum(r['impoconsumo']), cNum(r['costo'])]);
  });
  if (filas.length < 10) throw new Error('No parece el reporte de Productos (falta la columna Id Referencia).');
  cEscribir(cHoja(ss, H_MAESTRO, ENC_MAESTRO), ENC_MAESTRO.length, filas, [1, 3]);
  PropertiesService.getScriptProperties().setProperty('COSTEO_PRODUCTOS', cHoy());
  return filas.length + ' productos (' + filas.filter(function (f) { return f[3] === 1; }).length + ' activos).';
}

function cProcSaldos(ss, file) {
  const filas = cLeerArchivo(file);
  const t = cTabla(filas, 'Bodega');
  const fecha = cFechaSaldos(file.getName(), filas);
  const porId = {};
  t.forEach(function (r) {
    const id = cId(r['idreferencia']);
    if (!id) return;
    const costo = cNum(r['costo']);
    const prev = porId[id];
    if (!prev || (prev.costo <= 0 && costo > 0)) {
      porId[id] = { id: id, fam: cTxt(r['familia']), nombre: cTxt(r['referencia']), saldo: cNum(r['saldo']), costo: costo };
    }
  });
  const lista = Object.keys(porId).map(function (k) { return porId[k]; });
  if (lista.length < 10) throw new Error('No parece el reporte de Saldos de bodega (falta la columna idreferencia).');

  // Historial: reemplaza la foto de esa fecha
  const hh = cHoja(ss, H_HIST, ENC_HIST);
  let previas = hh.getLastRow() > 1 ? hh.getRange(2, 1, hh.getLastRow() - 1, 3).getValues() : [];
  previas = previas.filter(function (r) { return cFechaTxt(r[0]) !== fecha; });
  const nuevas = lista.filter(function (x) { return x.costo > 0; }).map(function (x) { return [fecha, x.id, x.costo]; });
  const todas = previas.map(function (r) { return [cFechaTxt(r[0]), cId(r[1]), r[2]]; }).concat(nuevas);
  cEscribir(hh, 3, todas, [1, 2]);

  // Costos actuales: solo si esta foto es la más reciente
  const maxPrevia = previas.reduce(function (m, r) { const f = cFechaTxt(r[0]); return f > m ? f : m; }, '');
  let nota = '';
  if (fecha >= maxPrevia) {
    const filasIns = lista.map(function (x) { return [x.id, x.fam, x.nombre, x.saldo, x.costo, fecha]; });
    cEscribir(cHoja(ss, H_INSUMOS, ENC_INSUMOS), ENC_INSUMOS.length, filasIns, [1, 6]);
    PropertiesService.getScriptProperties().setProperty('COSTEO_SALDOS', fecha);
  } else {
    nota = ' (es más antiguo que el último cargado: solo se guardó en el historial)';
  }
  return lista.length + ' insumos con costos al ' + fecha + '.' + nota;
}

function cFechaSaldos(nombre, filas) {
  let m = String(nombre).match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  const banner = String((filas[0] && filas[0][0]) || '');
  m = banner.match(/Fecha:\s*(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) return m[3] + '-' + m[2] + '-' + m[1];
  return Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd');
}


/* ======================================================================
 * 3. LECTURA PARA EL TABLERO
 * ====================================================================== */

/** Lee las hojas y calcula costos y márgenes. Devuelve { vacio: true, falta: [...] } si aún no hay datos. */
function cModelo(ss) {
  const base = cBase(ss, false);
  const falta = [];
  if (!Object.keys(base.insumos).length) falta.push('Saldos de bodega');
  if (!Object.keys(base.recetas).length) falta.push('Recetas');
  if (!Object.keys(base.maestro).length) falta.push('Productos');
  if (falta.length) return { vacio: true, falta: falta };
  return cCalcular(base, cRotacion(ss));
}

/** Todo lo que necesita la pestaña Costeo (una sola llamada). */
function getCosteo() {
  const r = cModelo(SpreadsheetApp.openById(SHEET_ID));
  if (r.vacio) return r;
  const p = PropertiesService.getScriptProperties();
  r.cargas = { recetas: p.getProperty('COSTEO_RECETAS') || '', productos: p.getProperty('COSTEO_PRODUCTOS') || '' };
  return r;
}

/** Lista de productos para las biblias, con el estado de su ficha. */
function getBiblia() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const r = cModelo(ss);
  if (r.vacio) return r;
  const productos = r.productos.filter(function (x) {
    return !x.at && (x.a === 'Cocina' || x.a === 'Postres' || x.a === 'Barra');
  }).map(function (x) {
    return { i: x.i, n: x.n, g: x.g, a: x.a, p: x.p, c: x.c, ok: x.ok, f: x.f, m: x.m, u: x.u, ni: x.ni };
  });
  const fichas = cLeerFichas(ss);
  const publicas = {};
  Object.keys(fichas).forEach(function (id) { publicas[id] = cFichaPublica(fichas[id]); });
  return { productos: productos, fichas: publicas, req: FICHA_REQ, objetivo: r.objetivo, fechas: r.fechas };
}

/** Guarda (o actualiza) la ficha de un producto. La llama el formulario de la web. */
function guardarFicha(f) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('El sistema está ocupado. Intenta de nuevo en unos segundos.');
  try {
    const id = cId(f && f.id);
    if (!id) throw new Error('Falta el ID del producto.');
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const h = ss.getSheetByName(H_FICHAS) || cHoja(ss, H_FICHAS, ENC_FICHAS);
    const nCols = ENC_FICHAS.length;
    const n = h.getLastRow();
    let fila = n + 1, previa = null;
    if (n > 1) {
      const vals = h.getRange(2, 1, n - 1, nCols).getValues();
      for (let i = 0; i < vals.length; i++) {
        if (cId(vals[i][0]) === id) { fila = i + 2; previa = vals[i]; break; }
      }
    }
    if (h.getMaxRows() < fila) h.insertRowsAfter(h.getMaxRows(), fila - h.getMaxRows());

    let fotoId = previa ? cTxt(previa[14]) : '';
    if (f.quitarFoto && fotoId) { cBorrarFoto(fotoId); fotoId = ''; }
    if (f.foto) {
      const b64 = String(f.foto).replace(/^data:image\/[a-zA-Z+]+;base64,/, '');
      if (b64.length > 7000000) throw new Error('La foto es muy pesada. Prueba con otra.');
      const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', 'ficha_' + id + '_' + Date.now() + '.jpg');
      const nuevo = cFotos().createFile(blob);
      if (fotoId) cBorrarFoto(fotoId);
      fotoId = nuevo.getId();
    }

    const area = ['Barra', 'Botella'].indexOf(cTxt(f.area)) >= 0 ? cTxt(f.area) : 'Cocina';
    const corto = function (v) { return cTxt(v).slice(0, 200); };
    const ahora = Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd HH:mm');
    const filaNueva = [id, corto(f.nombre), area, corto(f.porciones), cTxtLargo(f.preparacion, 6000), cTxtLargo(f.emplatado, 3000),
                       corto(f.alergenos), corto(f.merma), corto(f.vida), corto(f.cristaleria), corto(f.metodo), corto(f.hielo),
                       corto(f.guarnicion), cTxtLargo(f.notas, 3000), fotoId, ahora,
                       corto(f.origen), corto(f.tipo), corto(f.grados), cTxtLargo(f.notacata, 3000), corto(f.servicio), corto(f.maridaje)];
    const rango = h.getRange(fila, 1, 1, nCols);
    rango.setNumberFormat('@');   // texto plano: así nada que se escriba se vuelve fórmula
    rango.setValues([filaNueva]);
    return { ok: true, actualizado: ahora, tieneFoto: !!fotoId };
  } finally {
    lock.releaseLock();
  }
}

/** Foto de la ficha como imagen incrustada (data URI). Cadena vacía si no hay. */
function getFotoFicha(id) {
  id = cId(id);
  const fi = cLeerFichas(SpreadsheetApp.openById(SHEET_ID))[id];
  if (!fi || !fi.fotoId) return '';
  try {
    const blob = DriveApp.getFileById(fi.fotoId).getBlob();
    return 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    return '';
  }
}

/** Detalle de un plato: ingredientes con costo y su ficha técnica. */
function getFichaPlato(id) {
  id = cId(id);
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const base = cBase(ss, true);
  const rc = base.recetas[id];
  const mp = base.maestro[id];
  if (!rc) return null;
  const costos = cMapaCostos(cCostosActuales(base.insumos), base.manuales);
  let total = 0, faltan = 0;
  const lineas = rc.lineas.map(function (l) {
    const cu = costos[l[0]];
    const ok = cu !== undefined && (cu > 0 || base.manuales[l[0]]);
    const s = ok ? l[2] * cu : null;
    if (ok) total += s; else faltan++;
    return { i: l[0], n: l[1], c: l[2], u: l[3], cu: ok ? cu : null, s: s };
  }).sort(function (a, b) { return (b.s || 0) - (a.s || 0); });
  const pv = mp ? mp.precio : 0;
  const p = mp ? pv / (1 + mp.imp / 100) : 0;
  const ok = pv >= COSTEO_PRECIO_MIN && faltan === 0 && total > 0 && p > 0;
  const fi = cLeerFichas(ss)[id];
  return { i: id, n: rc.nombre || (mp && mp.nombre) || id, g: cNombreGrupo(rc.grupoTxt, mp && mp.grupo),
           a: cArea(cNumGrupo(rc.grupoTxt, mp && mp.grupo)), pv: pv, p: p, c: total, x: faltan,
           f: ok ? total / p : null, m: ok ? p - total : null, lineas: lineas,
           ficha: fi && cFichaConDatos(fi) ? cFichaPublica(fi) : null };
}

/** Fichas guardadas en la pestaña FICHAS, por ID de producto. */
function cLeerFichas(ss) {
  const out = {};
  cLeerHoja(ss, H_FICHAS, ENC_FICHAS.length).forEach(function (r) {
    const id = cId(r[0]);
    if (!id) return;
    out[id] = { id: id, nombre: cTxt(r[1]), area: cTxt(r[2]), porciones: cTxt(r[3]), preparacion: cTxtLargo(r[4], 6000),
                emplatado: cTxtLargo(r[5], 3000), alergenos: cTxt(r[6]), merma: cTxt(r[7]), vida: cTxt(r[8]),
                cristaleria: cTxt(r[9]), metodo: cTxt(r[10]), hielo: cTxt(r[11]), guarnicion: cTxt(r[12]),
                notas: cTxtLargo(r[13], 3000), fotoId: cTxt(r[14]), actualizado: cTxt(r[15]),
                origen: cTxt(r[16]), tipo: cTxt(r[17]), grados: cTxt(r[18]), notacata: cTxtLargo(r[19], 3000),
                servicio: cTxt(r[20]), maridaje: cTxt(r[21]) };
  });
  return out;
}

function cFichaPublica(f) {
  const o = {};
  Object.keys(f).forEach(function (k) { if (k !== 'fotoId') o[k] = f[k]; });
  o.tieneFoto = !!f.fotoId;
  return o;
}

function cFichaConDatos(f) {
  return ['porciones', 'preparacion', 'emplatado', 'alergenos', 'merma', 'vida', 'cristaleria', 'metodo', 'hielo', 'guarnicion', 'notas', 'fotoId',
          'origen', 'tipo', 'grados', 'notacata', 'servicio', 'maridaje']
    .some(function (k) { return !!f[k]; });
}

/** Lee las pestañas y arma el modelo en memoria. sinHist = true omite el historial (más rápido). */
function cBase(ss, sinHist) {
  const insumos = {};
  cLeerHoja(ss, H_INSUMOS, 6).forEach(function (r) {
    const id = cId(r[0]);
    if (id) insumos[id] = { fam: cTxt(r[1]), nombre: cTxt(r[2]), saldo: +r[3] || 0, costo: +r[4] || 0, fecha: cFechaTxt(r[5]) };
  });
  const manuales = {};
  cLeerHoja(ss, H_MANUALES, 4).forEach(function (r) {
    const id = cId(r[0]);
    if (!id || r[2] === '' || r[2] === null) return;
    let c = cNum(r[2]);
    const rend = cNum(r[3]);
    if (rend > 0) c = c / rend;
    manuales[id] = { nombre: cTxt(r[1]), costo: c };
  });
  const maestro = {};
  cLeerHoja(ss, H_MAESTRO, 7).forEach(function (r) {
    const id = cId(r[0]);
    if (id) maestro[id] = { nombre: cTxt(r[1]), grupo: cId(r[2]), activo: +r[3] === 1, precio: +r[4] || 0, imp: +r[5] || 0, costoMig: +r[6] || 0 };
  });
  const recetas = {};
  cLeerHoja(ss, H_RECETAS, 7).forEach(function (r) {
    const id = cId(r[0]);
    if (!id) return;
    if (!recetas[id]) recetas[id] = { nombre: cTxt(r[1]), grupoTxt: cTxt(r[2]), lineas: [] };
    recetas[id].lineas.push([cId(r[3]), cTxt(r[4]), +r[5] || 0, cTxt(r[6])]);
  });
  const hist = { fechas: [], mapas: {} };
  if (!sinHist) {
    cLeerHoja(ss, H_HIST, 3).forEach(function (r) {
      const f = cFechaTxt(r[0]);
      if (!f) return;
      if (!hist.mapas[f]) { hist.mapas[f] = {}; hist.fechas.push(f); }
      hist.mapas[f][cId(r[1])] = +r[2] || 0;
    });
    hist.fechas.sort().reverse();
  }
  const primero = Object.keys(insumos)[0];
  const fechaSaldos = (primero && insumos[primero].fecha) || hist.fechas[0] || '';
  return { insumos: insumos, manuales: manuales, maestro: maestro, recetas: recetas, hist: hist, fechaSaldos: fechaSaldos };
}

/** Unidades vendidas por semana de cada producto (últimas 13 semanas de la pestaña PRODUCTOS). */
function cRotacion(ss) {
  const out = { porId: {}, hasta: '' };
  const h = ss.getSheetByName(HOJA_PRODUCTOS);
  if (!h || h.getLastRow() < 2) return out;
  const filas = h.getRange(2, 1, h.getLastRow() - 1, 6).getValues();
  let max = '';
  filas.forEach(function (r) { const f = cFechaTxt(r[0]); if (f > max) max = f; });
  if (!max) return out;
  const desde = cSumarDias(max, -(COSTEO_SEMANAS * 7 - 1));
  filas.forEach(function (r) {
    const f = cFechaTxt(r[0]);
    if (f < desde || f > max) return;
    const id = cId(r[3]);
    if (!id) return;
    out.porId[id] = (out.porId[id] || 0) + (+r[5] || 0);
  });
  Object.keys(out.porId).forEach(function (k) { out.porId[k] = out.porId[k] / COSTEO_SEMANAS; });
  out.hasta = max;
  return out;
}


/* ======================================================================
 * 4. MOTOR DE COSTOS (funciones puras: no leen hojas)
 * ====================================================================== */

function cCostosActuales(insumos) {
  const m = {};
  Object.keys(insumos).forEach(function (id) { m[id] = insumos[id].costo; });
  return m;
}

/** Costo por unidad de cada insumo: el de bodega si es mayor a 0; si no, el de COSTOS_MANUALES. */
function cMapaCostos(mapaBase, manuales) {
  const m = {};
  Object.keys(mapaBase).forEach(function (id) { if (mapaBase[id] > 0) m[id] = mapaBase[id]; });
  Object.keys(manuales).forEach(function (id) { if (!(id in m)) m[id] = manuales[id].costo; });
  return m;
}

/** Suma cantidad × costo. Una línea sin costo cuenta como faltante (el plato queda incompleto). */
function cCostoReceta(lineas, costos, manuales) {
  let total = 0, faltan = 0;
  lineas.forEach(function (l) {
    const c = costos[l[0]];
    const conocido = c !== undefined && (c > 0 || (manuales && manuales[l[0]]));
    if (conocido) total += l[2] * c; else faltan++;
  });
  return { total: total, faltan: faltan };
}

function cArea(g) {
  const n = +g;
  if ([36, 38, 39, 42, 43, 45, 54].indexOf(n) >= 0) return 'Cocina';
  if (n === 40) return 'Postres';
  if ((n >= 10 && n <= 35) || [46, 51, 52, 53].indexOf(n) >= 0) return 'Barra';
  return 'Otros';
}

function cNombreGrupo(grupoTxt, grupoId) {
  const t = String(grupoTxt || '');
  const i = t.indexOf(' - ');
  return i >= 0 ? t.slice(i + 3).trim() : (grupoId ? 'Grupo ' + grupoId : '');
}

function cNumGrupo(grupoTxt, grupoId) {
  const m = String(grupoTxt || '').match(/^\s*(\d+)/);
  return m ? m[1] : cId(grupoId);
}

function cCalcular(base, rot) {
  const hist = base.hist || { fechas: [], mapas: {} };
  const fActual = base.fechaSaldos || hist.fechas[0] || '';
  const fPrev = hist.fechas.filter(function (f) { return f < fActual; })[0] || '';
  const costos = cMapaCostos(cCostosActuales(base.insumos), base.manuales);
  const costosPrev = fPrev ? cMapaCostos(hist.mapas[fPrev], base.manuales) : null;

  const productos = [];
  const usoInsumo = {};     // id -> { n, usos }
  const sinCostoUso = {};   // id -> { n, usos }
  const difMig = [];

  Object.keys(base.recetas).forEach(function (id) {
    const mp = base.maestro[id];
    if (!mp || !mp.activo || mp.precio <= 0) return;
    const rc = base.recetas[id];
    const cc = cCostoReceta(rc.lineas, costos, base.manuales);
    const pvNeto = mp.precio / (1 + mp.imp / 100);
    const atip = mp.precio < COSTEO_PRECIO_MIN;
    const ok = !atip && cc.faltan === 0 && cc.total > 0 && pvNeto > 0;
    const num = cNumGrupo(rc.grupoTxt, mp.grupo);
    const u = (rot.porId && rot.porId[id]) || 0;

    let d = null;
    if (ok && costosPrev) {
      const cp = cCostoReceta(rc.lineas, costosPrev, base.manuales);
      if (cp.faltan === 0 && cp.total > 0) d = cc.total / cp.total - 1;
    }

    rc.lineas.forEach(function (l) {
      const o = usoInsumo[l[0]] || (usoInsumo[l[0]] = { n: l[1], usos: 0 });
      o.usos++;
      const c = costos[l[0]];
      const conocido = c !== undefined && (c > 0 || base.manuales[l[0]]);
      if (!conocido) {
        const s = sinCostoUso[l[0]] || (sinCostoUso[l[0]] = { i: l[0], n: l[1], usos: 0 });
        s.usos++;
      }
    });

    const x = { i: id, n: rc.nombre || mp.nombre, g: cNombreGrupo(rc.grupoTxt, mp.grupo), a: cArea(num),
                pv: mp.precio, p: pvNeto, c: cc.total, x: cc.faltan, ok: ok, at: atip, ni: rc.lineas.length,
                f: ok ? cc.total / pvNeto : null, m: ok ? pvNeto - cc.total : null,
                u: u, k: ok ? (pvNeto - cc.total) * u : 0, d: d, cl: '' };
    productos.push(x);

    if (ok && mp.costoMig > 0) {
      const dm = cc.total / mp.costoMig - 1;
      if (Math.abs(dm) > 0.05) difMig.push({ i: id, n: x.n, c: cc.total, mig: mp.costoMig, d: dm });
    }
  });

  // Ingeniería de menú por área (Kasavana–Smith): popularidad vs. margen en pesos
  ['Cocina', 'Barra', 'Postres'].forEach(function (area) {
    const it = productos.filter(function (x) { return x.a === area && x.ok && x.u > 0; });
    if (!it.length) return;
    const totU = it.reduce(function (s, x) { return s + x.u; }, 0);
    const umbralPop = 0.7 * totU / it.length;
    const margenProm = it.reduce(function (s, x) { return s + x.m * x.u; }, 0) / totU;
    it.forEach(function (x) {
      const pop = x.u >= umbralPop, alto = x.m >= margenProm;
      x.cl = pop ? (alto ? 'E' : 'C') : (alto ? 'R' : 'P');
    });
  });
  productos.forEach(function (x) { if (!x.cl && x.ok && x.u === 0 && x.a !== 'Otros') x.cl = 'S'; });

  // Indicadores
  const vend = productos.filter(function (x) { return x.ok && x.u > 0; });
  const fcDe = function (arr) {
    const v = arr.reduce(function (s, x) { return s + x.p * x.u; }, 0);
    const c = arr.reduce(function (s, x) { return s + x.c * x.u; }, 0);
    return v > 0 ? c / v : null;
  };
  const kpi = {
    fcPond: fcDe(vend),
    fcArea: { Cocina: fcDe(vend.filter(function (x) { return x.a === 'Cocina'; })),
              Barra: fcDe(vend.filter(function (x) { return x.a === 'Barra'; })),
              Postres: fcDe(vend.filter(function (x) { return x.a === 'Postres'; })) },
    margenSemana: vend.reduce(function (s, x) { return s + x.m * x.u; }, 0),
    conVentas: vend.length,
    costeados: productos.filter(function (x) { return x.ok; }).length,
    total: productos.length,
    incompletos: productos.filter(function (x) { return !x.ok; }).length,
    altos: vend.filter(function (x) { return x.f > COSTEO_FC_AMBAR; }).length
  };

  // Alertas
  const sinCosto = Object.keys(sinCostoUso).map(function (k) { return sinCostoUso[k]; })
    .sort(function (a, b) { return b.usos - a.usos; });

  const sinReceta = Object.keys(base.maestro).filter(function (id) {
    const mp = base.maestro[id];
    return mp.activo && mp.precio > 0 && !base.recetas[id];
  }).map(function (id) {
    const mp = base.maestro[id];
    return { i: id, n: mp.nombre, pv: mp.precio };
  }).sort(function (a, b) { return b.pv - a.pv; });

  const precioBajo = productos.filter(function (x) { return x.at; })
    .map(function (x) { return { i: x.i, n: x.n, pv: x.pv }; });

  const cambios = [];
  if (costosPrev) {
    Object.keys(usoInsumo).forEach(function (id) {
      const a = costosPrev[id], b = costos[id];
      if (!(a > 0) || !(b > 0)) return;
      const d = b / a - 1;
      if (Math.abs(d) >= 0.10) {
        const ins = base.insumos[id];
        cambios.push({ i: id, n: (ins && ins.nombre) || usoInsumo[id].n, a: a, b: b, d: d, usos: usoInsumo[id].usos });
      }
    });
    cambios.sort(function (p, q) { return Math.abs(q.d) * q.usos - Math.abs(p.d) * p.usos; });
  }

  const negativos = Object.keys(base.insumos).filter(function (id) { return base.insumos[id].saldo < 0; }).map(function (id) {
    const s = base.insumos[id];
    return { i: id, n: s.nombre, s: s.saldo, v: s.saldo * s.costo };
  }).sort(function (a, b) { return a.v - b.v; });

  difMig.sort(function (a, b) { return Math.abs(b.d) - Math.abs(a.d); });

  return {
    productos: productos,
    kpi: kpi,
    objetivo: { verde: COSTEO_FC_VERDE, ambar: COSTEO_FC_AMBAR },
    fechas: { saldos: fActual, previa: fPrev, ventasHasta: rot.hasta || '' },
    alertas: {
      sinCosto: { total: sinCosto.length, lista: sinCosto.slice(0, 40) },
      sinReceta: { total: sinReceta.length, lista: sinReceta.slice(0, 40) },
      precioBajo: { total: precioBajo.length, lista: precioBajo.slice(0, 40) },
      cambios: { total: cambios.length, lista: cambios.slice(0, 30) },
      difMig: { total: difMig.length, lista: difMig.slice(0, 15) },
      negativos: { total: negativos.length, lista: negativos.slice(0, 15) }
    }
  };
}


/* ======================================================================
 * 5. UTILIDADES
 * ====================================================================== */

function cCarpeta() {
  const props = PropertiesService.getScriptProperties();
  const guardada = props.getProperty('COSTEO_CARPETA_ID');
  if (guardada) { try { return DriveApp.getFolderById(guardada); } catch (e) { /* se vuelve a buscar */ } }
  const it = DriveApp.getFoldersByName(COSTEO_CARPETA);
  const carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(COSTEO_CARPETA);
  props.setProperty('COSTEO_CARPETA_ID', carpeta.getId());
  return carpeta;
}

function cFotos() {
  const carpeta = cCarpeta();
  const it = carpeta.getFoldersByName('FOTOS FICHAS');
  return it.hasNext() ? it.next() : carpeta.createFolder('FOTOS FICHAS');
}

function cBorrarFoto(fileId) {
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* ya no existe */ }
}

/** Texto largo: conserva los saltos de línea (pasos de preparación). */
function cTxtLargo(v, max) {
  return String(v === null || v === undefined ? '' : v)
    .replace(/\r/g, '').replace(/[ \t]+/g, ' ').replace(/ ?\n ?/g, '\n').replace(/\n{3,}/g, '\n\n').trim().slice(0, max || 3000);
}

function cProcesados(carpeta) {
  const it = carpeta.getFoldersByName('PROCESADOS');
  return it.hasNext() ? it.next() : carpeta.createFolder('PROCESADOS');
}

function cHoja(ss, nombre, enc) {
  let h = ss.getSheetByName(nombre);
  if (!h) h = ss.insertSheet(nombre);
  h.getRange(1, 1, 1, enc.length).setValues([enc])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F3864').setHorizontalAlignment('center');
  h.setFrozenRows(1);
  return h;
}

function cEscribir(h, nCols, filas, colsTexto) {
  const ult = h.getLastRow();
  if (ult > 1) h.getRange(2, 1, ult - 1, nCols).clearContent();
  if (!filas.length) return;
  const necesarias = filas.length + 1;
  if (h.getMaxRows() < necesarias) h.insertRowsAfter(h.getMaxRows(), necesarias - h.getMaxRows());
  (colsTexto || []).forEach(function (c) { h.getRange(2, c, filas.length, 1).setNumberFormat('@'); });
  h.getRange(2, 1, filas.length, nCols).setValues(filas);
}

function cLeerHoja(ss, nombre, nCols) {
  const h = ss.getSheetByName(nombre);
  if (!h || h.getLastRow() < 2) return [];
  return h.getRange(2, 1, h.getLastRow() - 1, nCols).getValues();
}

/** Convierte el Excel del MIG (.xls o .xlsx) a una hoja temporal, lee los valores y la borra. */
function cLeerArchivo(file) {
  if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
    return SpreadsheetApp.openById(file.getId()).getSheets()[0].getDataRange().getValues();
  }
  const blob = file.getBlob();
  const b = blob.getBytes();
  // El MIG llama .xls a archivos que en realidad son .xlsx (empiezan con "PK")
  const esZip = b.length > 3 && b[0] === 80 && b[1] === 75;
  blob.setContentType(esZip ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/vnd.ms-excel');
  const id = cReintento(function () {
    if (Drive.Files.create) return Drive.Files.create({ name: 'tmp_costeo', mimeType: MimeType.GOOGLE_SHEETS }, blob).id;
    return Drive.Files.insert({ title: 'tmp_costeo', mimeType: MimeType.GOOGLE_SHEETS }, blob, { convert: true }).id;
  });
  try {
    return SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getValues();
  } finally {
    DriveApp.getFileById(id).setTrashed(true);
  }
}

function cReintento(fn) {
  for (let intento = 0; ; intento++) {
    try {
      return fn();
    } catch (e) {
      const limite = /rate limit|User Rate Limit|quota/i.test(String(e));
      if (!limite || intento >= 3) throw e;
      Utilities.sleep(5000 * Math.pow(2, intento));
    }
  }
}

/** Busca la fila de encabezados (primera columna = primeraColumna) y devuelve objetos por nombre de columna. */
function cTabla(filas, primeraColumna) {
  const objetivo = cNorm(primeraColumna);
  let h = -1;
  for (let i = 0; i < Math.min(filas.length, 30); i++) {
    if (cNorm(filas[i][0]) === objetivo) { h = i; break; }
  }
  if (h < 0) throw new Error('No encontré la columna "' + primeraColumna + '". ¿Es el reporte correcto?');
  const cols = filas[h].map(cNorm);
  const out = [];
  for (let i = h + 1; i < filas.length; i++) {
    const f = filas[i];
    if (f.every(function (v) { return v === '' || v === null; })) continue;
    const o = {};
    cols.forEach(function (c, j) { if (c) o[c] = f[j]; });
    out.push(o);
  }
  return out;
}

function cNorm(s) {
  return String(s === null || s === undefined ? '' : s).toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim();
}

function cId(v) {
  if (v === null || v === undefined || v === '') return '';
  return String(v).trim().replace(/\.0+$/, '');
}

function cTxt(v) {
  return String(v === null || v === undefined ? '' : v).replace(/\s+/g, ' ').trim();
}

function cNum(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/\s/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

function cFechaTxt(v) {
  if (v instanceof Date) return Utilities.formatDate(v, ZONA, 'yyyy-MM-dd');
  return String(v === null || v === undefined ? '' : v).slice(0, 10);
}

function cHoy() {
  return Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd');
}

function cSumarDias(fecha, dias) {
  const p = fecha.split('-');
  const d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2]) + dias * 86400000);
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}
