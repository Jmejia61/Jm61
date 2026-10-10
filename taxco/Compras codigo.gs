/**
 * COMPRAS DINÁMICAS TAXCO — v1.4.1
 * ================================
 * v1.4.1: si un ID está repetido en COMPRAS_PARAM (ej. 63007 Gordon's 700 ml y 1,5 L),
 *         el pedido guardado usa la fila ACTIVA, no la última.
 * Cruza las VENTAS por producto (pestaña PRODUCTOS) con las RECETAS del MIG para saber
 * cuánto de cada insumo se consume por semana, y con eso calcula mínimos, máximos y pedido
 * sugerido que se mueven solos con la venta. Barra y cocina separadas. Pedido por proveedor.
 *
 * Archivo NUEVO dentro del proyecto VENTAS TAXCO. No toca Código.gs ni Dashboard.gs.
 * Reusa las funciones de Costeo.gs: cBase, cLeerHoja, cHoja, cEscribir, cId, cTxt, cNum, cNorm, cFechaTxt, cSumarDias.
 * Los datos iniciales (proveedores, factores, máximos anteriores) están en ComprasDatos.gs.
 *
 * INVENTARIO MANUAL (v1.1)
 *  El stock sale del conteo que se digita en la pestaña Compras > Inventario (planilla por área).
 *  Cada conteo se guarda en la pestaña INVENTARIO (ID_INSUMO, AREA, CONTEO, FECHA). Se crea sola al guardar.
 *  El conteo se digita en la unidad de compra: botellas en barra, la misma unidad de la receta en cocina.
 *
 * DE DÓNDE SALE CADA DATO
 *  Ventas por producto   pestaña PRODUCTOS (ya se llena sola con los correos del MIG)
 *  Recetas               pestaña RECETAS    (Costeo.gs la llena desde Recetas*.xls en Drive)
 *  Costo (y saldo MIG)   pestaña INSUMOS    (Costeo.gs la llena desde SaldosBodega*.xlsx)
 *  Stock contado         pestaña INVENTARIO (se llena desde la planilla de la web)
 *  Proveedor, factor y   pestaña COMPRAS_PARAM (se crea con configurarCompras, la edita Julian)
 *  máximos del modelo anterior
 *
 * CÓMO SE CALCULA (por insumo)
 *  consumo diario = promedio de las últimas N semanas completas (lunes a domingo) ÷ 7
 *  mínimo  = consumo diario × (días de entrega + días de seguridad)
 *  máximo  = consumo diario × (días de entrega + días de cobertura)
 *  si el stock contado ≤ mínimo  →  pedir hasta el máximo   (en botellas o paquetes enteros)
 *
 * REGISTRO DE PEDIDOS (v1.4)
 *  Desde la pestaña Compras > Pedido se puede guardar el pedido de cada proveedor. Queda en la pestaña
 *  PEDIDOS (se crea sola): una fila por insumo, con cantidad, costo, valor, semana (lunes) y estado
 *  (Guardado / Enviado). Si guardas otra vez el mismo proveedor el mismo día, reemplaza el anterior.
 *  La pestaña Historial de la web muestra lo pedido semana a semana.
 *
 * CATEGORÍAS (v1.2)
 *  La planilla de inventario se puede ordenar por categoría (Proteínas, Lácteos, Destilados...).
 *  Las categorías están en ComprasCategorias.gs. Para poder cambiarlas desde la hoja, ejecuta una vez
 *  "categorizarCompras": agrega la columna CATEGORIA a COMPRAS_PARAM y la llena. Si no la ejecutas, igual funcionan.
 *
 * ACTUALIZAR DE v1.2/v1.3 A v1.4
 *  1. En el archivo "Compras codigo": borra todo el contenido y pega este archivo. Guarda.
 *  2. En Compras.html: borra todo el contenido y pega el nuevo. Guarda.
 *  3. Implementar > Gestionar implementaciones > lápiz > Nueva versión > Implementar.
 *  (La pestaña PEDIDOS se crea sola la primera vez que guardes un pedido.)
 */

const H_COMPRAS = 'COMPRAS_PARAM';
const ENC_COMPRAS = ['ID_INSUMO', 'INSUMO', 'AREA', 'PROVEEDOR', 'FACTOR', 'DIAS_ENTREGA', 'DIAS_COBERTURA', 'MAX_ACTUAL', 'ACTIVO', 'NOTA', 'CATEGORIA'];
const H_PED = 'PEDIDOS';
const ENC_PED = ['ID_PEDIDO', 'FECHA', 'SEMANA', 'PROVEEDOR', 'AREA', 'ID_INSUMO', 'INSUMO', 'CANTIDAD', 'UNIDAD', 'FACTOR', 'UNIDAD_BASE', 'COSTO_UNIT', 'VALOR', 'ESTADO'];
const H_INV = 'INVENTARIO';
const ENC_INV = ['ID_INSUMO', 'AREA', 'CONTEO', 'FECHA'];
const COMPRAS_SEMANAS = 8;      // semanas de historia que se leen de PRODUCTOS
const COMPRAS_DEF = { semanasBase: 4, entrega: 2, cobertura: 7, seguridad: 1 };

/** Orden en que se muestran las categorías en la planilla (las que no estén aquí salen al final). */
const COMPRAS_CAT_ORDEN = [
  'Proteínas (carnes y aves)', 'Pescados y mariscos', 'Lácteos y huevos', 'Frutas, verduras y tubérculos', 'Panadería, masas y granos',
  'Abarrotes y secos', 'Salsas y condimentos', 'Especias, sal y chiles', 'Aceites, vinagres y vinos de cocina', 'Postres, dulces y chocolates', 'Preparados de cocina',
  'Destilados', 'Licores dulces y aperitivos', 'Vinos y espumosos', 'Cervezas', 'Gaseosas, sodas y aguas', 'Jugos, pulpas y frutas',
  'Jarabes, mixers y complementos', 'Café', 'Preparados de barra', 'Otros'
];

/** Categoría de cada insumo según ComprasCategorias.gs: { idInsumo: 'Categoría' } */
function comprasCategorias() {
  const out = {};
  COMPRAS_CAT_TXT.split('\n').forEach(function (l) {
    const p = l.replace(/\u00a0/g, ' ').trim().split('|');
    if (p.length >= 2 && p[0].trim()) out[p[0].trim()] = p[1].trim();
  });
  return out;
}

/** ¿La fila de COMPRAS_PARAM está activa? (columna ACTIVO: 1, verdadero o "sí") */
function cpActivo(r) {
  const v = String(r[8]).trim();
  return r[8] === 1 || r[8] === true || v === '1' || /^s[ií]$/i.test(v);
}

/** Convierte los datos iniciales (ComprasDatos.gs) en filas. Cada línea: ID | insumo | área | proveedor | factor | máximo anterior | activo | nota */
function comprasSemilla() {
  return COMPRAS_TXT.split('\n').map(function (l) { return l.replace(/\u00a0/g, ' ').trim(); }).filter(function (l) { return l; }).map(function (l) {
    const p = l.split('|');
    return [p[0], p[1], p[2], p[3], +p[4] || 1, +p[5] || 0, +p[6] || 0, p[7] || ''];
  });
}


/* ======================================================================
 * 1. CONFIGURAR (ejecutar una vez)
 * ====================================================================== */
function configurarCompras() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const h = cHoja(ss, H_COMPRAS, ENC_COMPRAS);
  cHoja(ss, H_INV, ENC_INV);
  cHoja(ss, H_PED, ENC_PED);
  if (h.getLastRow() < 2) {
    const cats = comprasCategorias();
    const filas = comprasSemilla().map(function (s) {
      return [s[0], s[1], s[2], s[3], s[4], '', '', s[5], s[6], s[7], cats[s[0]] || ''];
    });
    h.getRange(2, 1, filas.length, 1).setNumberFormat('@');
    h.getRange(2, 1, filas.length, ENC_COMPRAS.length).setValues(filas);
    Logger.log('✅ COMPRAS_PARAM creada con ' + filas.length + ' insumos (' + filas.filter(function (f) { return +f[8] === 1; }).length + ' activos).');
  } else {
    Logger.log('COMPRAS_PARAM ya tiene datos: no se tocó.');
  }
  h.getRange('E1').setNote('Cuántas unidades base (ml, g) trae cada unidad de compra. Botella de 750 ml = 750. Si pides en la misma unidad de la receta, 1.');
  h.getRange('F1').setNote('Opcional. Días que tarda el proveedor en entregar. Vacío = usa el valor general de la pestaña.');
  h.getRange('G1').setNote('Opcional. Cada cuántos días pides a este proveedor. Vacío = usa el valor general de la pestaña.');
  h.getRange('H1').setNote('Máximo del modelo anterior, en unidades de compra (botellas en barra). Solo sirve para comparar.');
  h.getRange('I1').setNote('1 = entra al cálculo. 0 = en revisión (mira la columna NOTA).');
  h.getRange('K1').setNote('Categoría para ordenar la planilla de inventario. Puedes cambiarla: la de esta columna manda.');
  h.setColumnWidth(2, 300); h.setColumnWidth(4, 170); h.setColumnWidth(10, 360); h.setColumnWidth(11, 230);
}

/** Opcional: agrega la columna CATEGORIA a COMPRAS_PARAM y la llena donde esté vacía. No toca nada más. */
function categorizarCompras() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const h = cHoja(ss, H_COMPRAS, ENC_COMPRAS);
  const n = h.getLastRow();
  if (n < 2) { Logger.log('COMPRAS_PARAM está vacía. Ejecuta configurarCompras primero.'); return; }
  const cats = comprasCategorias();
  const rango = h.getRange(2, 11, n - 1, 1);
  const ids = h.getRange(2, 1, n - 1, 1).getValues();
  const actuales = rango.getValues();
  let nuevos = 0;
  const salida = actuales.map(function (r, i) {
    const previa = cTxt(r[0]);
    if (previa) return [previa];
    const c = cats[cId(ids[i][0])];
    if (c) { nuevos++; return [c]; }
    return [''];
  });
  rango.setValues(salida);
  h.getRange('K1').setNote('Categoría para ordenar la planilla de inventario. Puedes cambiarla: la de esta columna manda.');
  h.setColumnWidth(11, 230);
  Logger.log('✅ Categorías puestas en ' + nuevos + ' insumos. Puedes cambiar cualquiera en la columna CATEGORIA.');
}


/* ======================================================================
 * 2. DATOS PARA LA PESTAÑA
 * ====================================================================== */
function getCompras() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const filasParam = cLeerHoja(ss, H_COMPRAS, ENC_COMPRAS.length);
  if (!filasParam.length) return { vacio: true, falta: ['La pestaña COMPRAS_PARAM está vacía. Ejecuta configurarCompras.'] };
  const base = cBase(ss, true);
  if (!Object.keys(base.recetas).length) return { vacio: true, falta: ['La pestaña RECETAS está vacía. Sube Recetas*.xls a la carpeta COSTEO TAXCO.'] };
  const h = ss.getSheetByName(HOJA_PRODUCTOS);
  const ventas = (h && h.getLastRow() > 1) ? h.getRange(2, 1, h.getLastRow() - 1, 8).getValues() : [];
  const res = cpCalcular(base, filasParam, ventas, cpLeerConteos(ss));
  if (!res.vacio) { res.hoy = cpHoy(); res.pedidosHoy = cpPedidosHoy(ss); }
  return res;
}

/** Cálculo puro (no lee hojas): fácil de probar. */
function cpCalcular(base, filasParam, ventas, conteos) {
  conteos = conteos || {};
  const cats = comprasCategorias();
  const recetas = base.recetas;
  // --- ventas por producto y semana (lunes a domingo) ---
  // Algunas filas AGRUPADOS no traen código: se cruzan por nombre con la receta (si el nombre es único)
  const porNombre = {}, repetido = {};
  Object.keys(recetas).forEach(function (pid) {
    const k = cNorm(recetas[pid].nombre);
    if (porNombre[k] && porNombre[k] !== pid) repetido[k] = true; else porNombre[k] = pid;
  });
  const idVenta = function (r) {
    const id = cId(r[3]);
    if (id) return id;
    const k = cNorm(r[4]);
    return (porNombre[k] && !repetido[k]) ? porNombre[k] : (k ? 'N:' + k : '');
  };
  const fuentes = ventas.filter(function (r) { return (r[7] === 'INFORME' || r[7] === 'AGRUPADOS') && idVenta(r); });
  let max = '';
  fuentes.forEach(function (r) { const f = cFechaTxt(r[0]); if (f > max) max = f; });
  if (!max) return { vacio: true, falta: ['PRODUCTOS no tiene ventas por producto (INFORME o AGRUPADOS).'] };
  const fin = cpUltimoDomingo(max);                       // última semana completa termina en este domingo
  const inicio0 = cSumarDias(fin, -(COMPRAS_SEMANAS * 7 - 1));
  const semanas = [];
  for (let k = 0; k < COMPRAS_SEMANAS; k++) semanas.push({ i: cSumarDias(inicio0, k * 7), d: 0 });
  const fechasVistas = {};
  const vendidos = {};          // idProducto -> [unidades por semana]
  const nombres = {};           // idProducto -> [nombre, total $ últimas 4 semanas]
  const ini4 = cSumarDias(fin, -27);
  fuentes.forEach(function (r) {
    const f = cFechaTxt(r[0]);
    if (f < inicio0 || f > fin) return;
    const k = Math.floor(cpDias(inicio0, f) / 7);
    const id = idVenta(r);
    if (!vendidos[id]) vendidos[id] = new Array(COMPRAS_SEMANAS).fill(0);
    vendidos[id][k] += +r[5] || 0;
    if (!nombres[id]) nombres[id] = [cTxt(r[4]), 0];
    if (f >= ini4) nombres[id][1] += +r[6] || 0;
    const clave = f; if (!fechasVistas[clave]) { fechasVistas[clave] = true; semanas[k].d++; }
  });

  // --- explosión de recetas hasta insumos base (las subrecetas se abren) ---
  const memo = {};
  function aplanar(id, pila) {
    if (memo[id]) return memo[id];
    const out = {};
    ((recetas[id] && recetas[id].lineas) || []).forEach(function (l) {
      const ins = l[0], q = l[2];
      if (recetas[ins] && pila.indexOf(ins) < 0) {
        const sub = aplanar(ins, pila.concat(id));
        Object.keys(sub).forEach(function (b) { out[b] = (out[b] || 0) + q * sub[b]; });
      } else out[ins] = (out[ins] || 0) + q;
    });
    memo[id] = out;
    return out;
  }
  const consumo = {};            // idInsumoBase -> [consumo por semana en unidad base]
  let uTot = 0, uCon = 0;
  const sinReceta = [];
  Object.keys(vendidos).forEach(function (pid) {
    const v = vendidos[pid];
    const u4 = v.slice(-4).reduce(function (a, b) { return a + b; }, 0);
    uTot += u4;
    if (!recetas[pid]) { if (u4 > 0) sinReceta.push({ n: nombres[pid][0], u: u4, t: nombres[pid][1] }); return; }
    uCon += u4;
    const plano = aplanar(pid, []);
    Object.keys(plano).forEach(function (b) {
      if (!consumo[b]) consumo[b] = new Array(COMPRAS_SEMANAS).fill(0);
      for (let k = 0; k < COMPRAS_SEMANAS; k++) consumo[b][k] += plano[b] * v[k];
    });
  });
  sinReceta.sort(function (a, b) { return b.t - a.t; });

  // --- unidad base de cada insumo (la de la receta) ---
  const unidad = {};
  Object.keys(recetas).forEach(function (pid) {
    recetas[pid].lineas.forEach(function (l) { if (!unidad[l[0]] && l[3]) unidad[l[0]] = l[3].toLowerCase(); });
  });
  const costoDe = function (id) {
    const i = base.insumos[id];
    if (i && i.costo > 0) return i.costo;
    return (base.manuales[id] && base.manuales[id].costo) || 0;
  };

  // --- insumos del modelo de compras ---
  const items = [], activos = {}, enRevision = {};
  let conteoFecha = '';
  filasParam.forEach(function (r) {
    const id = cId(r[0]);
    if (!id) return;
    const on = cpActivo(r);
    if (!on) { enRevision[id] = true; return; }
    const factor = cNum(r[4]) > 0 ? cNum(r[4]) : 1;
    const ins = base.insumos[id];
    activos[id] = true;
    items.push({
      id: id,
      n: cTxt(r[1]) || (ins && ins.nombre) || id,
      a: cTxt(r[2]) || 'Cocina',
      k: cTxt(r[10]) || cats[id] || 'Otros',
      p: cTxt(r[3]) || 'Sin proveedor',
      f: factor,
      u: unidad[id] || '',
      c: costoDe(id),
      s: ins ? ins.saldo : null,
      w: (consumo[id] || new Array(COMPRAS_SEMANAS).fill(0)).map(function (x) { return Math.round(x * 100) / 100; }),
      de: r[5] === '' ? null : cNum(r[5]),
      dc: r[6] === '' ? null : cNum(r[6]),
      mx: cNum(r[7]) * factor,
      cv: conteos[id] ? conteos[id].v : null,                 // conteo tal como se digitó (unidad de compra)
      q: conteos[id] ? conteos[id].v * factor : null,         // conteo en unidad base (g, ml, und)
      cf: conteos[id] ? conteos[id].f : ''
    });
    if (conteos[id] && conteos[id].f > conteoFecha) conteoFecha = conteos[id].f;
  });

  // --- lo que se consume pero no está en el modelo de compras ---
  const sinParam = [];
  Object.keys(consumo).forEach(function (id) {
    if (activos[id]) return;
    const sem = consumo[id].slice(-4).reduce(function (a, b) { return a + b; }, 0) / 4;
    if (sem <= 0) return;
    const ins = base.insumos[id];
    const c = costoDe(id);
    sinParam.push({ id: id, n: (ins && ins.nombre) || cpNombreInsumo(recetas, id), fam: (ins && ins.fam) || '', u: unidad[id] || '', sem: Math.round(sem), c: c, v: sem * c, rev: !!enRevision[id] });
  });
  sinParam.sort(function (a, b) { return b.v - a.v; });

  return {
    vacio: false,
    fechas: { saldos: base.fechaSaldos || '', ventasHasta: max, semanaFin: fin, conteo: conteoFecha },
    semanas: semanas,
    def: COMPRAS_DEF,
    items: items,
    catOrden: COMPRAS_CAT_ORDEN,
    sinParam: sinParam.slice(0, 60),
    nSinParam: sinParam.length,
    cobertura: { pct: uTot > 0 ? uCon / uTot : 0, sinReceta: sinReceta.slice(0, 8), nSinReceta: sinReceta.length }
  };
}

/* ======================================================================
 * 3. INVENTARIO MANUAL (planilla de conteo de la web)
 * ====================================================================== */

/** Conteos guardados: { idInsumo: { v: cantidad contada, f: 'aaaa-mm-dd hh:mm' } } */
function cpLeerConteos(ss) {
  const out = {};
  cLeerHoja(ss, H_INV, ENC_INV.length).forEach(function (r) {
    const id = cId(r[0]);
    if (!id || r[2] === '' || r[2] === null) return;
    out[id] = { v: cNum(r[2]), f: cpFechaHora(r[3]) };
  });
  return out;
}

function cpFechaHora(v) {
  if (v instanceof Date) return Utilities.formatDate(v, ZONA, 'yyyy-MM-dd HH:mm');
  return String(v === null || v === undefined ? '' : v).slice(0, 16);
}

/**
 * Guarda los conteos que llegan de la planilla. cambios = [{ id, a, v }]
 * v = cantidad contada; v vacío = borrar el conteo de ese insumo.
 */
function guardarConteo(cambios) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('El sistema está ocupado. Intenta de nuevo en unos segundos.');
  try {
    if (!cambios || !cambios.length) return { ok: true, n: 0, fecha: '' };
    if (cambios.length > 1000) throw new Error('Demasiados cambios en un solo guardado. Guarda por partes.');
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const validos = {};
    cLeerHoja(ss, H_COMPRAS, 1).forEach(function (r) { const i = cId(r[0]); if (i) validos[i] = 1; });
    const lista = cambios.filter(function (c) { return validos[cId(c && c.id)]; });
    const h = ss.getSheetByName(H_INV) || cHoja(ss, H_INV, ENC_INV);
    const fecha = Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd HH:mm');
    const previas = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, ENC_INV.length).getValues() : [];
    const filas = cpAplicarConteo(previas, lista, fecha);
    cEscribir(h, ENC_INV.length, filas, [1, 4]);
    return { ok: true, n: lista.length, fecha: fecha, total: filas.length };
  } finally {
    lock.releaseLock();
  }
}

/** Mezcla los conteos nuevos con los guardados (función pura). Devuelve las filas completas de INVENTARIO. */
function cpAplicarConteo(previas, cambios, fecha) {
  const mapa = {};
  previas.forEach(function (r) {
    const id = cId(r[0]);
    if (id && r[2] !== '' && r[2] !== null) mapa[id] = [id, cTxt(r[1]), cNum(r[2]), cpFechaHora(r[3])];
  });
  cambios.forEach(function (c) {
    const id = cId(c.id);
    if (!id) return;
    const crudo = (c.v === null || c.v === undefined) ? '' : String(c.v).trim().replace(',', '.');
    if (crudo === '') { delete mapa[id]; return; }
    const n = Number(crudo);
    if (!isFinite(n) || n < 0) throw new Error('Conteo inválido para el insumo ' + id + ': "' + c.v + '".');
    mapa[id] = [id, cTxt(c.a), Math.round(n * 1000) / 1000, fecha];
  });
  return Object.keys(mapa).map(function (k) { return mapa[k]; });
}

/* ======================================================================
 * 4. REGISTRO DE PEDIDOS (guardar el pedido de cada proveedor y verlo por semana)
 * ====================================================================== */

function cpHoy() { return Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd'); }

/** Lunes de la semana de una fecha 'aaaa-mm-dd'. */
function cpLunes(f) {
  const p = f.split('-');
  const dow = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay();   // 0 = domingo
  const atras = (dow + 6) % 7;
  return atras ? cSumarDias(f, -atras) : f;
}

function cpUnidadCompra(f, u) { return f > 1 ? ((f >= 500 && u === 'ml') ? 'botellas' : 'paquetes') : (u || 'und'); }

function cpSlug(t) { return String(t).toUpperCase().replace(/[^A-Z0-9]+/g, '').slice(0, 14) || 'PROV'; }

/** Pedidos que ya se guardaron hoy: { proveedor: { id, f, e } } */
function cpPedidosHoy(ss) {
  const hoy = cpHoy(), out = {};
  cLeerHoja(ss, H_PED, ENC_PED.length).forEach(function (r) {
    const f = cpFechaHora(r[1]);
    if (f.slice(0, 10) !== hoy) return;
    out[cTxt(r[3])] = { id: cTxt(r[0]), f: f, e: cTxt(r[13]) };
  });
  return out;
}

/**
 * Arma las filas de PEDIDOS de un proveedor (función pura).
 * ped = { p, estado, lineas: [{ id, c }] } · ref = { param: {id:{n,a,f}}, unidad: {id:u}, costo: function(id) }
 */
function cpFilasPedido(ped, ref, ahora, hoy) {
  const prov = cTxt(ped && ped.p);
  if (!prov) throw new Error('Falta el proveedor.');
  const estado = ped.estado === 'Enviado' ? 'Enviado' : 'Guardado';
  const id = hoy.replace(/-/g, '') + '-' + ahora.slice(11, 13) + ahora.slice(14, 16) + '-' + cpSlug(prov);
  const sem = cpLunes(hoy);
  const filas = [];
  (ped.lineas || []).forEach(function (l) {
    const iid = cId(l && l.id);
    if (!iid) return;
    const c = Number(String(l.c).replace(',', '.'));
    if (!isFinite(c) || c < 0) throw new Error('Cantidad inválida en el insumo ' + iid + '.');
    if (c === 0) return;
    const pr = ref.param[iid];
    if (!pr) throw new Error('El insumo ' + iid + ' no está en COMPRAS_PARAM.');
    const u = ref.unidad[iid] || '';
    const costo = ref.costo(iid);
    filas.push([id, ahora, sem, prov, pr.a, iid, pr.n, Math.round(c * 1000) / 1000, cpUnidadCompra(pr.f, u), pr.f, u, costo, Math.round(c * pr.f * costo), estado]);
  });
  if (!filas.length) throw new Error('El pedido de ' + prov + ' no tiene cantidades.');
  return { id: id, prov: prov, estado: estado, filas: filas, total: filas.reduce(function (a, r) { return a + r[12]; }, 0) };
}

/** Quita de la tabla el pedido del mismo proveedor y día (para reemplazarlo). Devuelve { filas, quitadas }. */
function cpQuitarPedidoDelDia(tabla, prov, hoy) {
  const quedan = tabla.filter(function (r) { return !(cTxt(r[3]) === prov && cpFechaHora(r[1]).slice(0, 10) === hoy); });
  return { filas: quedan, quitadas: tabla.length - quedan.length };
}

/** Guarda uno o varios pedidos (uno por proveedor). lista = [{ p, estado, lineas: [{ id, c }] }] */
function guardarPedidos(lista) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('El sistema está ocupado. Intenta de nuevo en unos segundos.');
  try {
    if (!lista || !lista.length) return { ok: true, resultados: [], fecha: '' };
    if (lista.length > 80) throw new Error('Demasiados proveedores en un solo guardado.');
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const param = {};
    cLeerHoja(ss, H_COMPRAS, ENC_COMPRAS.length).forEach(function (r) {
      const i = cId(r[0]);
      if (!i) return;
      // Un ID repetido no debe pisar la fila activa con una inactiva (ej. 63007 Gordon's 700 ml y 1,5 L)
      if (param[i] && param[i].on && !cpActivo(r)) return;
      param[i] = { n: cTxt(r[1]), a: cTxt(r[2]) || 'Cocina', f: cNum(r[4]) > 0 ? cNum(r[4]) : 1, on: cpActivo(r) };
    });
    const base = cBase(ss, true);
    const unidad = {};
    Object.keys(base.recetas).forEach(function (pid) {
      base.recetas[pid].lineas.forEach(function (l) { if (!unidad[l[0]] && l[3]) unidad[l[0]] = l[3].toLowerCase(); });
    });
    const ref = {
      param: param, unidad: unidad,
      costo: function (id) {
        const i = base.insumos[id];
        if (i && i.costo > 0) return i.costo;
        return (base.manuales[id] && base.manuales[id].costo) || 0;
      }
    };
    const hoy = cpHoy(), ahora = Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd HH:mm');
    const h = ss.getSheetByName(H_PED) || cHoja(ss, H_PED, ENC_PED);
    let tabla = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, ENC_PED.length).getValues() : [];
    let hubo = false;
    const nuevas = [], resultados = [];
    lista.forEach(function (ped) {
      const x = cpFilasPedido(ped, ref, ahora, hoy);
      const q = cpQuitarPedidoDelDia(tabla, x.prov, hoy);
      tabla = q.filas;
      if (q.quitadas) hubo = true;
      nuevas.push.apply(nuevas, x.filas);
      resultados.push({ p: x.prov, id: x.id, estado: x.estado, total: x.total, n: x.filas.length, reemplazo: q.quitadas > 0 });
    });
    if (hubo) cEscribir(h, ENC_PED.length, tabla.concat(nuevas), [1, 2, 3, 6]);
    else {
      const r0 = h.getLastRow() + 1;
      if (h.getMaxRows() < r0 + nuevas.length) h.insertRowsAfter(h.getMaxRows(), r0 + nuevas.length - h.getMaxRows());
      [1, 2, 3, 6].forEach(function (c) { h.getRange(r0, c, nuevas.length, 1).setNumberFormat('@'); });
      h.getRange(r0, 1, nuevas.length, ENC_PED.length).setValues(nuevas);
    }
    return { ok: true, fecha: ahora, resultados: resultados };
  } finally {
    lock.releaseLock();
  }
}

/** Guarda el pedido de un solo proveedor. */
function guardarPedido(ped) {
  const r = guardarPedidos([ped]);
  return { ok: true, fecha: r.fecha, id: r.resultados[0].id, estado: r.resultados[0].estado, total: r.resultados[0].total, n: r.resultados[0].n, reemplazo: r.resultados[0].reemplazo };
}

/** Pedidos guardados de las últimas 16 semanas, agrupados por pedido (el más reciente primero). */
function getPedidos() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const h = ss.getSheetByName(H_PED);
  const filas = (h && h.getLastRow() > 1) ? h.getRange(2, 1, h.getLastRow() - 1, ENC_PED.length).getValues() : [];
  const hoy = cpHoy();
  return { hoy: hoy, pedidos: cpAgruparPedidos(filas, hoy) };
}

function cpAgruparPedidos(filas, hoy) {
  const desde = cSumarDias(hoy, -7 * 16);
  const mapa = {}, orden = [];
  filas.forEach(function (r) {
    const id = cTxt(r[0]);
    if (!id) return;
    const sem = cFechaTxt(r[2]);
    if (sem < desde) return;
    if (!mapa[id]) { mapa[id] = { id: id, f: cpFechaHora(r[1]), s: sem, p: cTxt(r[3]), e: cTxt(r[13]), t: 0, l: [] }; orden.push(id); }
    const v = cNum(r[12]);
    mapa[id].t += v;
    mapa[id].l.push({ i: cId(r[5]), n: cTxt(r[6]), a: cTxt(r[4]), c: cNum(r[7]), u: cTxt(r[8]), v: v });
  });
  return orden.map(function (k) { return mapa[k]; }).sort(function (a, b) { return a.f < b.f ? 1 : (a.f > b.f ? -1 : 0); });
}

/** Borra un pedido guardado (por su ID_PEDIDO). */
function eliminarPedido(id) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('El sistema está ocupado. Intenta de nuevo en unos segundos.');
  try {
    id = cTxt(id);
    if (!id) throw new Error('Falta el ID del pedido.');
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const h = ss.getSheetByName(H_PED);
    if (!h || h.getLastRow() < 2) return { ok: true, n: 0 };
    const tabla = h.getRange(2, 1, h.getLastRow() - 1, ENC_PED.length).getValues();
    const quedan = tabla.filter(function (r) { return cTxt(r[0]) !== id; });
    if (quedan.length !== tabla.length) cEscribir(h, ENC_PED.length, quedan, [1, 2, 3, 6]);
    return { ok: true, n: tabla.length - quedan.length };
  } finally {
    lock.releaseLock();
  }
}

function cpNombreInsumo(recetas, id) {
  const pids = Object.keys(recetas);
  for (let i = 0; i < pids.length; i++) {
    const ls = recetas[pids[i]].lineas;
    for (let j = 0; j < ls.length; j++) if (ls[j][0] === id) return ls[j][1];
  }
  return id;
}

/** Días entre dos fechas 'aaaa-mm-dd'. */
function cpDias(a, b) {
  const pa = a.split('-'), pb = b.split('-');
  return Math.round((Date.UTC(+pb[0], +pb[1] - 1, +pb[2]) - Date.UTC(+pa[0], +pa[1] - 1, +pa[2])) / 86400000);
}

/** Último domingo igual o anterior a la fecha (fin de la última semana completa). */
function cpUltimoDomingo(f) {
  const p = f.split('-');
  const dow = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay();   // 0 = domingo
  return dow === 0 ? f : cSumarDias(f, -dow);
}
