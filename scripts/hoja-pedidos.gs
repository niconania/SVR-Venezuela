/**
 * SVR Venezuela — recepción de pedidos de la tienda online
 *
 * Este archivo NO corre en el sitio web. Se pega dentro de Google Apps Script.
 *
 * ── QUÉ HACE ────────────────────────────────────────────────────────
 * Recibe cada pedido de svr-venezuela.com/tienda y lo escribe como una
 * fila en una hoja de cálculo. La hoja la crea él solo la primera vez
 * que llega un pedido: no hay que crear nada a mano.
 *
 * También recibe las participaciones del sorteo SVR × Blua Clinic
 * (svr-venezuela.com/blua): pestaña «Sorteo Blua» en la misma hoja, la
 * foto de la factura en Drive y un correo a jmedicalas@ con la foto.
 *
 * ── CÓMO INSTALARLO ─────────────────────────────────────────────────
 *  1. Pega TODO este archivo en Código.gs (borra lo que haya).
 *  2. Guarda con Cmd+S.
 *  3. ⚙️ Configuración del proyecto → Propiedades de la secuencia de
 *     comandos → Agregar propiedad:   CLAVE_ADMIN = (la clave del panel
 *     de agotados). Sin esta propiedad el panel queda cerrado para todos;
 *     los pedidos siguen entrando con normalidad.
 *  4. Implementar → Gestionar implementaciones → lápiz ✏️
 *     → Versión: Nueva versión → Implementar.
 *  5. Abre la URL /exec en el navegador: debe decir que está activo
 *     y darte el enlace de la hoja.
 */

var NOMBRE_HOJA = 'SVR Venezuela — Pedidos tienda online';

/* Clave para entrar a la pantalla de agotados (svr-venezuela.com/tienda/4dm1n/).
   NO se escribe en este archivo: el repositorio de GitHub es público. Vive
   en las Propiedades de la secuencia de comandos (ver paso 3 arriba), y se
   compara aquí, en el servidor: nunca viaja al navegador de nadie. Para
   cambiarla, edita esa propiedad y vuelve a implementar una versión nueva.
   Si no hay clave guardada, nadie entra (una clave vacía no vale). */
function claveValida(clave) {
  var guardada = PropertiesService.getScriptProperties().getProperty('CLAVE_ADMIN');
  return !!guardada && String(clave) === String(guardada);
}

/* Las columnas nuevas se agregan SIEMPRE al final: así las filas que ya
   están en la hoja no se descuadran. Por eso 'Cupón', 'Descuento %' y
   'Total con descuento' siguen aquí aunque los cupones estén apagados:
   quedan vacías, y quitarlas movería de lugar todo lo que viene después. */
var ENCABEZADOS = [
  'Fecha', 'Estado', 'Nombre y apellido', 'Cédula', 'Teléfono',
  'Entrega', 'Transporte', 'Ciudad / Dirección',
  'Productos', 'Unidades', 'Total USD', 'Tasa BCV', 'Total Bs', 'Notas',
  'Cupón', 'Descuento %', 'Total con descuento', 'Ubicación GPS', 'Comprobante', 'Aviso de stock', 'Fuera de horario'
];

var ANCHOS = [140, 90, 190, 110, 120, 210, 100, 260, 420, 80, 90, 100, 110, 200,
              110, 100, 140, 260, 260, 240, 190];

/**
 * Devuelve la hoja de pedidos. La primera vez la crea en tu Drive y
 * guarda su identificador, para usar siempre la misma de ahí en adelante.
 */
function obtenerHoja() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('ID_HOJA');
  var libro = null;

  if (id) {
    try { libro = SpreadsheetApp.openById(id); }
    catch (err) { libro = null; }   // la borraron o la movieron: se crea otra
  }

  if (!libro) {
    libro = SpreadsheetApp.create(NOMBRE_HOJA);
    props.setProperty('ID_HOJA', libro.getId());
  }

  var hoja = libro.getSheets()[0];

  // Deja los encabezados listos la primera vez.
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(ENCABEZADOS);
    formatearEncabezados(hoja);
  } else if (hoja.getLastColumn() < ENCABEZADOS.length) {
    // La hoja es de una versión anterior: se completan las columnas
    // que faltan al final, sin tocar lo que ya está escrito.
    var desde = hoja.getLastColumn() + 1;
    var nuevas = ENCABEZADOS.slice(desde - 1);
    hoja.getRange(1, desde, 1, nuevas.length).setValues([nuevas]);
    formatearEncabezados(hoja);
  }

  return hoja;
}

function formatearEncabezados(hoja) {
  hoja.getRange(1, 1, 1, ENCABEZADOS.length)
      .setFontWeight('bold')
      .setBackground('#1a1a1a')
      .setFontColor('#ffffff')
      .setVerticalAlignment('middle');
  hoja.setRowHeight(1, 34);
  hoja.setFrozenRows(1);
  for (var i = 0; i < ANCHOS.length; i++) hoja.setColumnWidth(i + 1, ANCHOS[i]);
}

/* ── Comprobantes de pago ───────────────────────────────────────────
   La tienda manda la captura ya reducida, en base64. Aquí se guarda
   como archivo en tu Drive y se adjunta al correo de administración.
   El archivo queda PRIVADO: solo lo ve quien tenga acceso a tu Drive. */

var CARPETA_COMPROBANTES = 'SVR — Comprobantes de pago';

function carpetaComprobantes() {
  var it = DriveApp.getFoldersByName(CARPETA_COMPROBANTES);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_COMPROBANTES);
}

/**
 * Guarda el comprobante y devuelve { archivo, url } o null si no vino
 * ninguno. Nunca lanza: un problema con la imagen no debe tumbar el pedido.
 */
function guardarComprobante(datos, nombreCliente) {
  if (!datos) return { vacio: true };
  try {
    // Viene como "data:image/jpeg;base64,XXXX"
    var partes = String(datos).match(/^data:([^;]+);base64,([\s\S]+)$/);
    if (!partes) return { error: 'formato no reconocido' };

    var tipo = partes[1];
    var bytes = Utilities.base64Decode(partes[2]);
    var ext = tipo.indexOf('png') > -1 ? 'png' :
              tipo.indexOf('pdf') > -1 ? 'pdf' : 'jpg';
    var sello = Utilities.formatDate(new Date(), 'America/Caracas', 'yyyy-MM-dd HH-mm');
    var nombre = sello + ' - ' + String(nombreCliente || 'cliente').slice(0, 40) + '.' + ext;

    var blob = Utilities.newBlob(bytes, tipo, nombre);
    var archivo = carpetaComprobantes().createFile(blob);
    return { archivo: archivo, url: archivo.getUrl() };
  } catch (err) {
    // El pedido sigue su curso, pero se dice por qué falló el archivo.
    return { error: String(err).slice(0, 180) };
  }
}

/* ── Inventario ─────────────────────────────────────────────────────
   Una fila por producto en la pestaña «Inventario». La columna Stock
   puede estar VACÍA: eso significa «sin control», y el producto se
   vende sin descontar nada. En cuanto se escribe un número, ese
   producto pasa a llevar cuenta y se descuenta con cada pedido.

   Así se puede activar el control uno por uno, a medida que se cuenta
   el inventario real, sin apagar la tienda mientras tanto. */

var HOJA_INVENTARIO = 'Inventario';
var INV_ENCABEZADOS = ['Producto', 'SKU', 'Slug', 'Stock', 'Actualizado'];

function obtenerHojaInventario() {
  var libro = obtenerHoja().getParent();
  var h = libro.getSheetByName(HOJA_INVENTARIO);
  if (!h) {
    h = libro.insertSheet(HOJA_INVENTARIO);
    h.appendRow(INV_ENCABEZADOS);
    h.getRange(1, 1, 1, INV_ENCABEZADOS.length)
      .setFontWeight('bold').setBackground('#1a1a1a').setFontColor('#ffffff');
    h.setFrozenRows(1);
    [260, 140, 200, 90, 150].forEach(function (a, i) { h.setColumnWidth(i + 1, a); });
    h.hideColumns(3);        // el slug es de uso interno
  }
  return h;
}

/** Devuelve { slug: {stock: number|null, fila: n} }. null = sin control. */
function leerInventario() {
  var h = obtenerHojaInventario();
  var inv = {};
  if (h.getLastRow() < 2) return inv;
  var filas = h.getRange(2, 1, h.getLastRow() - 1, INV_ENCABEZADOS.length).getValues();
  for (var i = 0; i < filas.length; i++) {
    var slug = String(filas[i][2]).trim();
    if (!slug) continue;
    var v = filas[i][3];
    var stock = (v === '' || v === null) ? null : Number(v);
    inv[slug] = { stock: (isNaN(stock) ? null : stock), fila: i + 2 };
  }
  return inv;
}

/** Crea las filas que falten, sin tocar los stocks ya escritos. */
function sembrarInventario(productos) {
  if (!productos || !productos.length) return leerInventario();
  var h = obtenerHojaInventario();
  var inv = leerInventario();
  var nuevas = [];
  for (var i = 0; i < productos.length; i++) {
    var p = productos[i];
    if (!p || !p.slug || inv[p.slug]) continue;
    nuevas.push([p.nombre || '', p.sku || '', p.slug, '', '']);
  }
  if (nuevas.length) {
    h.getRange(h.getLastRow() + 1, 1, nuevas.length, INV_ENCABEZADOS.length).setValues(nuevas);
  }
  return leerInventario();
}

/** Escribe los stocks que manda la pantalla de administración. */
function guardarInventario(cambios) {
  var h = obtenerHojaInventario();
  var inv = leerInventario();
  var sello = new Date().toLocaleString('es-VE');
  for (var slug in cambios) {
    if (!inv[slug]) continue;
    var v = cambios[slug];
    var valor = (v === null || v === '' || v === undefined) ? '' : Math.max(0, Math.floor(Number(v)) || 0);
    h.getRange(inv[slug].fila, 4).setValue(valor);
    h.getRange(inv[slug].fila, 5).setValue(sello);
  }
  return leerInventario();
}

/**
 * Descuenta el stock de un pedido. Usa un candado para que dos pedidos
 * simultáneos no vendan la misma última unidad.
 * Devuelve { faltantes: [...] } con lo que no alcanzó, si hubo algo.
 */
function descontarStock(items) {
  if (!items || !items.length) return { faltantes: [] };

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (err) { return { faltantes: [], error: 'ocupado' }; }

  try {
    var h = obtenerHojaInventario();
    var inv = leerInventario();
    var faltantes = [];
    var sello = new Date().toLocaleString('es-VE');

    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      var reg = inv[it.slug];
      if (!reg || reg.stock === null) continue;      // sin control: no se toca

      var quedan = reg.stock - it.cantidad;
      if (quedan < 0) {
        faltantes.push(it.slug + ' (pidió ' + it.cantidad + ', había ' + reg.stock + ')');
        quedan = 0;
      }
      h.getRange(reg.fila, 4).setValue(quedan);
      h.getRange(reg.fila, 5).setValue(sello);
    }
    return { faltantes: faltantes };
  } catch (err) {
    return { faltantes: [], error: String(err).slice(0, 120) };
  } finally {
    lock.releaseLock();
  }
}

/** Para la tienda: solo lo que necesita saber, sin números de fila. */
function inventarioPublico() {
  var inv = leerInventario(), out = {};
  for (var slug in inv) if (inv[slug].stock !== null) out[slug] = inv[slug].stock;
  return out;
}

/** Para la pantalla de administración: incluye los que están sin control. */
function inventarioPlano(inv) {
  var out = {};
  for (var slug in inv) out[slug] = inv[slug].stock;   // null = sin control
  return out;
}

/* ── Productos agotados ─────────────────────────────────────────────
   Marca manual, independiente del stock. Sirve para retirar algo de
   la venta aunque haya unidades (un lote dañado, por ejemplo). */

function leerAgotados() {
  var raw = PropertiesService.getScriptProperties().getProperty('AGOTADOS');
  try { return raw ? JSON.parse(raw) : []; }
  catch (err) { return []; }
}

function guardarAgotados(lista) {
  if (!lista || !lista.length) {
    PropertiesService.getScriptProperties().deleteProperty('AGOTADOS');
    return [];
  }
  var limpia = [];
  for (var i = 0; i < lista.length && i < 200; i++) {
    var s = String(lista[i]).slice(0, 60);
    if (s && limpia.indexOf(s) === -1) limpia.push(s);
  }
  PropertiesService.getScriptProperties().setProperty('AGOTADOS', JSON.stringify(limpia));
  return limpia;
}

/* ── Precios y ofertas ──────────────────────────────────────────────
   El precio de lista vive en products.json (publicado en el sitio). Aquí
   se guardan los CAMBIOS que administración hace desde /tienda/4dm1n/:
   un precio nuevo y, si hay oferta, el «precio antes» (el que la tienda
   muestra tachado). La tienda los aplica encima del catálogo al cargar,
   así que un cambio se ve al instante, sin publicar nada. Cada guardado
   queda anotado en la pestaña «Historial de precios» de esta hoja.
   Guardado: { slug: { p: precio vigente, a: precio antes (o null), t: fecha } } */
var PRECIO_MIN = 0.5;
var PRECIO_MAX = 2000;
var MAX_PRECIOS = 60;
var HOJA_HIST_PRECIOS = 'Historial de precios';
var HIST_PRECIOS_ENCABEZADOS = ['Fecha', 'Producto', 'Slug', 'Precio de lista', 'Antes (publicado)', 'Ahora',
                                'Precio antes de la oferta', 'Descuento', 'Qué se hizo'];

function leerPreciosCrudos() {
  var raw = PropertiesService.getScriptProperties().getProperty('PRECIOS');
  try {
    var o = raw ? JSON.parse(raw) : {};
    return (o && typeof o === 'object' && !(o instanceof Array)) ? o : {};
  } catch (err) { return {}; }
}

/** Lo que ve la tienda: { slug: { precio, antes (o null), t } } */
function preciosPublicos() {
  var crudos = leerPreciosCrudos(), out = {};
  for (var slug in crudos) {
    var c = crudos[slug];
    out[slug] = { precio: c.p, antes: c.a || null, t: c.t || '' };
  }
  return out;
}

function redondear2(n) { return Math.round(n * 100) / 100; }

/** Un número (o texto numérico) dentro del rango de precios, redondeado a centavos; si no, null. */
function precioValido(v) {
  var esNumero = typeof v === 'number' || (typeof v === 'string' && /^\d+(\.\d{1,6})?$/.test(v));
  if (!esNumero) return null;
  var n = redondear2(Number(v));
  return (isFinite(n) && n >= PRECIO_MIN && n <= PRECIO_MAX) ? n : null;
}

function hojaHistorialPrecios() {
  var libro = obtenerHoja().getParent();
  var hoja = libro.getSheetByName(HOJA_HIST_PRECIOS);
  if (!hoja) {
    hoja = libro.insertSheet(HOJA_HIST_PRECIOS);
    hoja.appendRow(HIST_PRECIOS_ENCABEZADOS);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, HIST_PRECIOS_ENCABEZADOS.length)
        .setFontWeight('bold').setBackground('#d0405b').setFontColor('#ffffff');
    [130, 230, 190, 110, 120, 90, 150, 90, 230].forEach(function (ancho, i) { hoja.setColumnWidth(i + 1, ancho); });
  }
  return hoja;
}

/**
 * Aplica los cambios de precios que manda la pantalla de administración.
 *   cambios: { slug: { precio, antes } | null }   (null = volver al precio de lista)
 *   meta:    { slug: { nombre, lista } }          (solo para anotar el historial)
 * Todo o nada: si un solo dato no es válido no se guarda ninguno.
 */
function guardarPrecios(cambios, meta) {
  if (!cambios || typeof cambios !== 'object' || cambios instanceof Array) return { ok: false, error: 'cambios' };
  var slugs = Object.keys(cambios);
  if (!slugs.length || slugs.length > MAX_PRECIOS) return { ok: false, error: 'cambios' };

  var nuevos = {};
  for (var i = 0; i < slugs.length; i++) {
    var slug = slugs[i], c = cambios[slug];
    if (!/^[a-z0-9][a-z0-9-]{1,58}$/.test(slug)) return { ok: false, error: 'slug', slug: slug };
    if (c === null) { nuevos[slug] = null; continue; }
    if (!c || typeof c !== 'object') return { ok: false, error: 'precio', slug: slug };
    var precio = precioValido(c.precio);
    if (precio === null) return { ok: false, error: 'precio', slug: slug };
    var antes = null;
    if (c.antes !== null && c.antes !== undefined) {
      antes = precioValido(c.antes);
      if (antes === null || antes <= precio) return { ok: false, error: 'antes', slug: slug };
    }
    nuevos[slug] = { p: precio, a: antes };
  }

  var filas = [];
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var actuales = leerPreciosCrudos();
    var ahora = Utilities.formatDate(new Date(), 'America/Caracas', "yyyy-MM-dd'T'HH:mm:ssXXX");
    var fecha = Utilities.formatDate(new Date(), 'America/Caracas', 'yyyy-MM-dd HH:mm');
    for (var s in nuevos) {
      var previo = actuales[s] || null, nuevo = nuevos[s];
      if (nuevo === null && !previo) continue;                                        // ya estaba en el precio de lista
      if (nuevo && previo && previo.p === nuevo.p && (previo.a || null) === nuevo.a) continue;   // sin cambios

      var m = (meta && meta[s] && typeof meta[s] === 'object') ? meta[s] : {};
      var lista = precioValido(m.lista);
      var accion = nuevo === null ? 'Volvió al precio de lista'
                 : nuevo.a ? (previo && previo.a ? 'Cambió la oferta' : 'Puso una oferta')
                 : (previo && previo.a ? 'Quitó la oferta y fijó el precio' : 'Cambió el precio');
      filas.push([fecha, String(m.nombre || s).slice(0, 80), s, lista === null ? '' : lista,
                  previo ? previo.p : (lista === null ? '' : lista),
                  nuevo ? nuevo.p : (lista === null ? 'lista' : lista),
                  nuevo && nuevo.a ? nuevo.a : '',
                  nuevo && nuevo.a ? Math.round((1 - nuevo.p / nuevo.a) * 1000) / 10 + ' %' : '',
                  accion]);
      if (nuevo === null) delete actuales[s];
      else { nuevo.t = ahora; actuales[s] = nuevo; }
    }
    if (Object.keys(actuales).length > MAX_PRECIOS) return { ok: false, error: 'cambios' };
    var props = PropertiesService.getScriptProperties();
    if (Object.keys(actuales).length) props.setProperty('PRECIOS', JSON.stringify(actuales));
    else props.deleteProperty('PRECIOS');
  } finally {
    lock.releaseLock();
  }

  if (filas.length) {
    try {
      var hoja = hojaHistorialPrecios();
      filas.forEach(function (f) { hoja.appendRow(f); });
    } catch (err) { /* el precio ya quedó guardado; si falla la anotación no se deshace */ }
  }
  return { ok: true, precios: preciosPublicos(), cambios: filas.length };
}

/* ── Sorteo SVR × Blua Clinic ──────────────────────────────────────
   La página svr-venezuela.com/blua/ manda cada participación con la
   acción «sorteo-blua». Queda una fila en la pestaña «Sorteo Blua» de
   esta misma hoja, la foto de la factura en Drive y sale un correo a
   SORTEO_CORREOS con la foto adjunta (al responderlo, le contestas al
   participante). Para cerrar la promoción: SORTEO_ABIERTO = false y
   publica una versión nueva; la página deja de aceptar participaciones. */
var SORTEO_CORREOS = 'jmedicalas@jmedicalcorp.com';
var SORTEO_ABIERTO = true;
var SORTEO_HOJA = 'Sorteo Blua';
var SORTEO_CARPETA = 'SVR — Sorteo Blua (facturas)';
var SORTEO_ENCABEZADOS = ['Fecha', 'N.º', 'Nombre y apellido', 'Zona de Caracas', 'Correo', 'Factura', 'Revisión'];
var SORTEO_TIPOS = /^(image\/(jpeg|png|webp|heic|heif)|application\/pdf)$/;

function hojaSorteo() {
  var libro = obtenerHoja().getParent();
  var hoja = libro.getSheetByName(SORTEO_HOJA);
  if (!hoja) {
    hoja = libro.insertSheet(SORTEO_HOJA);
    hoja.appendRow(SORTEO_ENCABEZADOS);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, SORTEO_ENCABEZADOS.length)
        .setFontWeight('bold').setBackground('#0b5a72').setFontColor('#ffffff');
    [130, 60, 200, 170, 230, 320, 140].forEach(function (ancho, i) { hoja.setColumnWidth(i + 1, ancho); });
  }
  return hoja;
}

function carpetaSorteo() {
  var it = DriveApp.getFoldersByName(SORTEO_CARPETA);
  return it.hasNext() ? it.next() : DriveApp.createFolder(SORTEO_CARPETA);
}

/** Valida y guarda una participación. Devuelve { ok, numero } o { ok: false, error }. */
function registrarSorteo(d) {
  if (!SORTEO_ABIERTO) return { ok: false, error: 'cerrado' };
  if (d.sitio) return { ok: true, numero: 0 };          // campo trampa lleno: es un robot, no se guarda nada

  var nombre = String(d.nombre || '').trim().replace(/\s+/g, ' ').slice(0, 80);
  var zona = String(d.zona || '').trim().slice(0, 70);
  var correo = String(d.correo || '').trim().toLowerCase().slice(0, 100);
  if (nombre.split(' ').length < 2) return { ok: false, error: 'nombre' };
  if (!zona) return { ok: false, error: 'zona' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@.]{2,}$/.test(correo)) return { ok: false, error: 'correo' };
  if (d.acepta !== true) return { ok: false, error: 'acepta' };

  // La factura llega como "data:image/jpeg;base64,XXXX"
  var partes = String((d.factura && d.factura.datos) || '').match(/^data:([^;]+);base64,([\s\S]+)$/);
  if (!partes || !SORTEO_TIPOS.test(partes[1])) return { ok: false, error: 'factura' };
  var bytes;
  try { bytes = Utilities.base64Decode(partes[2]); } catch (err) { return { ok: false, error: 'factura' }; }
  if (!bytes.length || bytes.length > 10 * 1024 * 1024) return { ok: false, error: 'factura' };

  // El número se asigna con candado: dos participaciones a la vez no se pisan.
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var hoja = hojaSorteo();
    var numero = Math.max(0, hoja.getLastRow() - 1) + 1;
    var tipo = partes[1];
    var ext = tipo === 'application/pdf' ? 'pdf' : tipo === 'image/png' ? 'png' :
              tipo === 'image/webp' ? 'webp' : /hei[cf]/.test(tipo) ? 'heic' : 'jpg';
    var archivo = carpetaSorteo().createFile(
      Utilities.newBlob(bytes, tipo, 'N' + ('000' + numero).slice(-4) + ' - ' + nombre.slice(0, 40) + '.' + ext));
    var fecha = Utilities.formatDate(new Date(), 'America/Caracas', 'yyyy-MM-dd HH:mm');
    hoja.appendRow([fecha, numero, nombre, zona, correo, archivo.getUrl(), '']);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  avisarSorteo({ numero: numero, nombre: nombre, zona: zona, correo: correo, fecha: fecha, archivo: archivo,
                 urlHoja: hoja.getParent().getUrl() + '#gid=' + hoja.getSheetId() });
  return { ok: true, numero: numero };
}

function avisarSorteo(p) {
  if (!SORTEO_CORREOS) return;
  try {
    var n = ('000' + p.numero).slice(-4);
    MailApp.sendEmail({
      to: SORTEO_CORREOS,
      replyTo: p.correo,
      name: 'SVR × Blua Clinic',
      subject: 'Sorteo SVR × Blua — participación N.º ' + n + ' — ' + p.nombre,
      body:
        'Nueva participación en el sorteo SVR × Blua Clinic\n\n' +
        'N.º:     ' + n + '\n' +
        'Fecha:   ' + p.fecha + '\n' +
        'Nombre:  ' + p.nombre + '\n' +
        'Zona:    ' + p.zona + '\n' +
        'Correo:  ' + p.correo + '\n\n' +
        'La foto de la factura va adjunta. También queda en Drive:\n' + p.archivo.getUrl() + '\n\n' +
        'Todas las participaciones (pestaña «Sorteo Blua»):\n' + p.urlHoja + '\n\n' +
        'Para escribirle al participante, responde este correo.',
      attachments: [p.archivo.getBlob()]
    });
  } catch (err) {
    // Si el correo falla, la participación ya quedó guardada en la hoja.
  }
}

/**
 * Recibe el pedido que manda la tienda, o el cambio de agotados
 * que manda la pantalla de administración.
 */
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);

    // Cambio de disponibilidad desde la pantalla de administración.
    if (d.accion === 'agotados') {
      if (!claveValida(d.clave)) {
        return responder({ ok: false, error: 'clave' });
      }
      return responder({ ok: true, agotados: guardarAgotados(d.agotados) });
    }

    // La pantalla de administración crea las filas que falten y lee el inventario.
    if (d.accion === 'inventario-leer') {
      if (!claveValida(d.clave)) return responder({ ok: false, error: 'clave' });
      return responder({ ok: true, inventario: inventarioPlano(sembrarInventario(d.productos)) });
    }

    // La pantalla de administración escribe los stocks.
    if (d.accion === 'inventario-guardar') {
      if (!claveValida(d.clave)) return responder({ ok: false, error: 'clave' });
      return responder({ ok: true, inventario: inventarioPlano(guardarInventario(d.cambios || {})) });
    }

    // Precios y ofertas desde la pantalla de administración.
    if (d.accion === 'precios-guardar') {
      if (!claveValida(d.clave)) return responder({ ok: false, error: 'clave' });
      return responder(guardarPrecios(d.cambios, d.meta));
    }

    // Participación en el sorteo SVR × Blua Clinic (página /blua/).
    if (d.accion === 'sorteo-blua') return responder(registrarSorteo(d));

    // Un pedido no trae «accion». Si trae una que no existe (una página
    // vieja pidiendo algo que ya se quitó), se rechaza: no crea una fila.
    if (d.accion) return responder({ ok: false, error: 'accion' });

    var hoja = obtenerHoja();

    // Se descuenta el stock ANTES de escribir la fila, para que el aviso
    // de faltantes quede registrado junto al pedido.
    var desc = descontarStock(d.items || []);
    var avisoStock = desc.faltantes.length
      ? 'STOCK INSUFICIENTE: ' + desc.faltantes.join(' · ')
      : (desc.error ? 'No se pudo descontar: ' + desc.error : '');

    var comp = guardarComprobante(d.comprobante, d.nombre);

    hoja.appendRow([
      d.fecha || new Date().toLocaleString('es-VE'),
      'Nuevo',
      d.nombre || '',
      d.cedula || '',
      d.telefono || '',
      d.entrega || '',
      d.transporte || '',
      d.destino || '',
      d.productos || '',
      d.unidades || '',
      d.totalUsd || '',
      d.tasa || '',
      d.totalBs || '',
      '',                         // Notas
      '', '', '',                 // Cupón · Descuento % · Total con descuento (reservadas, ver ENCABEZADOS)
      d.gps || '',
      (comp && comp.url) || '',
      avisoStock,
      d.fueraDeHorario || ''
    ]);

    // Resalta la fila nueva para que salte a la vista.
    var n = hoja.getLastRow();
    hoja.getRange(n, 1, 1, ENCABEZADOS.length).setBackground('#fff4f6');
    hoja.getRange(n, 2).setFontWeight('bold').setFontColor('#d0405b');

    avisarPorCorreo(d, hoja.getParent().getUrl(), (comp && comp.url) ? comp : null);

    return responder({ ok: true, fila: n,
      comprobante: (comp && comp.url) || '',
      comprobanteError: (comp && comp.error) || '' });

  } catch (err) {
    return responder({ ok: false, error: String(err) });
  }
}

/**
 * Al abrir la URL en el navegador confirma que está activo y da el
 * enlace de la hoja, para no tener que buscarla en Drive.
 */
function doGet(e) {
  var accion = (e && e.parameter && e.parameter.accion) || '';

  // La tienda pregunta qué productos están agotados y cuánto stock hay.
  if (accion === 'agotados') {
    return responder({ ok: true, agotados: leerAgotados(), inventario: inventarioPublico(), precios: preciosPublicos() });
  }

  // La página del sorteo pregunta si ya puede mandar participaciones
  // (una versión vieja de este archivo las guardaría como pedidos).
  if (accion === 'sorteo-estado') {
    return responder({ ok: true, sorteo: 'blua', abierto: SORTEO_ABIERTO });
  }

  // La pantalla de administración comprueba la clave al entrar.
  if (accion === 'entrar') {
    var bien = claveValida(e.parameter.clave);
    return responder(bien ? { ok: true, agotados: leerAgotados(), precios: preciosPublicos() }
                          : { ok: false, error: 'clave' });
  }

  try {
    var hoja = obtenerHoja();
    var pedidos = Math.max(0, hoja.getLastRow() - 1);
    return ContentService.createTextOutput(
      'Receptor de pedidos SVR activo.\n\n' +
      'Pedidos recibidos: ' + pedidos + '\n' +
      'Agotados: ' + (leerAgotados().length || 'ninguno') + '\n' +
      'Hoja: ' + hoja.getParent().getUrl()
    ).setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput('Error: ' + err)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * Avisa a administración por correo. Deja CORREOS vacío para desactivarlo.
 */
var CORREOS = 'santa.requiz@jmedicalcorp.com';

function avisarPorCorreo(d, urlHoja, comp) {
  if (!CORREOS) return;
  try {
    var cuerpo =
      'Pedido nuevo en la tienda\n\n' +
      'Cliente:    ' + d.nombre + '\n' +
      'Cédula:     ' + d.cedula + '\n' +
      'Teléfono:   ' + d.telefono + '\n\n' +
      'Entrega:    ' + d.entrega + '\n' +
      (d.transporte ? 'Transporte: ' + d.transporte + '\n' : '') +
      'Destino:    ' + d.destino + '\n' +
      (d.gps ? 'Ubicación:  ' + d.gps + '\n' : '') + '\n' +
      'Productos:\n  ' + String(d.productos).split(' | ').join('\n  ') + '\n\n' +
      'Total:      $' + d.totalUsd +
      (d.totalBs ? '  (Bs ' + d.totalBs + ' a tasa ' + d.tasa + ')' : '') + '\n' +
      (comp ? '\nComprobante de pago adjunto.\n' + comp.url + '\n' :
              '\nSin comprobante: el pago se coordina por WhatsApp.\n') + '\n' +
      urlHoja;

    var correo = {
      to: CORREOS,
      subject: 'Pedido nuevo — ' + d.nombre + ' — $' + d.totalUsd +
               (comp ? ' — con comprobante' : ''),
      body: cuerpo
    };
    if (comp) correo.attachments = [comp.archivo.getBlob()];

    MailApp.sendEmail(correo);
  } catch (err) {
    // Si el correo falla, el pedido ya quedó guardado en la hoja.
  }
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Opcional: córrelo a mano desde el editor (botón Ejecutar) para crear
 * la hoja de una vez y ver el enlace en el registro de ejecución.
 */
function crearHojaAhora() {
  var hoja = obtenerHoja();
  Logger.log('Hoja lista: ' + hoja.getParent().getUrl());
}

/**
 * ⚠️ CÓRREME UNA VEZ desde el editor (elige «autorizarDrive» arriba y
 * dale a Ejecutar) la primera vez que instales la parte de comprobantes.
 *
 * Google no pide el permiso de Drive al reimplementar, solo al ejecutar.
 * Sin ese permiso los comprobantes no se guardan y el pedido llega sin
 * la imagen. Esto crea la carpeta y un archivo de prueba que puedes
 * borrar; con eso el permiso queda dado para siempre.
 */
function autorizarDrive() {
  var carpeta = carpetaComprobantes();
  var prueba = carpeta.createFile(
    Utilities.newBlob('Permiso de Drive concedido. Puedes borrar este archivo.',
                      'text/plain', 'prueba-permiso.txt'));
  Logger.log('✅ Listo. Carpeta: ' + carpeta.getUrl());
  Logger.log('Archivo de prueba (puedes borrarlo): ' + prueba.getUrl());
  return carpeta.getUrl();
}
