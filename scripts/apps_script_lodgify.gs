/**
 * Lodgify -> hoja "Ingresos Diarios" (Hoja de Google). Pegar en Extensiones > Apps Script.
 * Antes: Configuración del proyecto > Propiedades del script > agregar LODGIFY_API_KEY.
 * Luego ejecutar una vez crearActivador() y aceptar los permisos.
 */
var HOJA = 'Ingresos Diarios';
var CORREO = 'admicanavida@gmail.com';
var DIAS_ATRAS = 2;          // revisa reservas creadas desde hace N días (no duplica)
var PRIMERA_FILA = 6;

var UNIDAD = {417302: 'Caracara (2 hab)', 417829: 'Lapa (3 hab)', 418298: 'Loro (4 hab)', 418311: 'Love Nest',
  418312: 'Quetzal (4 hab)', 418313: 'Torogoz (4 hab)', 537300: 'Toucan (3 hab)', 537321: 'Oropendola (2 hab)',
  537325: 'Cusinga (1 hab)', 555489: 'Gecko (Condo 3)', 556474: 'Lizardo (Condo 5)', 556475: 'Salamandra (Condo 4)',
  556478: 'Rana (Condo 6)', 598738: 'Iguana (Condo 2)', 733216: 'Pavon', 740510: 'Tortuga'};
var CANAL = {AirbnbIntegration: 'Airbnb', BookingCom: 'Booking.com', Expedia: 'Expedia', Manual: 'Directo / WhatsApp',
  OH: 'Web (Lodgify)'};

function crearActivador() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('actualizarReservas').timeBased().everyDays(1).atHour(20).create();
}

function lodgify_(path) {
  var key = PropertiesService.getScriptProperties().getProperty('LODGIFY_API_KEY');
  var r = UrlFetchApp.fetch('https://api.lodgify.com/v2' + path,
    {headers: {'X-ApiKey': key, accept: 'application/json'}, muteHttpExceptions: true});
  if (r.getResponseCode() !== 200) throw new Error('Lodgify ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 200));
  return JSON.parse(r.getContentText());
}

function dia_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

function actualizarReservas() {
  var desde = new Date(Date.now() - DIAS_ATRAS * 86400000);
  var desdeStr = dia_(desde);
  var props = {};
  lodgify_('/properties?size=100').items.forEach(function (p) { props[p.id] = p.name; });

  var reservas = [], page = 1, items;
  do {
    items = lodgify_('/reservations/bookings?size=50&page=' + page + '&includeCount=false&updatedSince=' +
      dia_(new Date(desde.getTime() - 86400000))).items;
    items.forEach(function (b) {
      if (b.created_at.slice(0, 10) < desdeStr || b.is_deleted || b.status === 'Declined' || b.status === 'Open') return;
      var sub = b.subtotals || {}, suma = 0;
      Object.keys(sub).forEach(function (k) { suma += sub[k] || 0; });
      var noches = Math.round((new Date(b.departure) - new Date(b.arrival)) / 86400000);
      reservas.push({
        llegada: b.arrival, unidad: UNIDAD[b.property_id] || props[b.property_id] || String(b.property_id),
        nombre: b.guest.name, canal: CANAL[b.source] || b.source, noches: noches,
        monto: Math.round((b.total_amount || suma) * 100) / 100
      });
    });
    page++;
  } while (items.length === 50);
  reservas.sort(function (a, b) { return a.llegada < b.llegada ? -1 : a.llegada > b.llegada ? 1 : 0; });

  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA);
  var n = Math.max(hoja.getMaxRows() - PRIMERA_FILA + 1, 1);
  var datos = hoja.getRange(PRIMERA_FILA, 1, n, 8).getValues();
  var ultima = PRIMERA_FILA - 1, vistos = {};
  datos.forEach(function (f, i) {
    if (f[0] !== '' || f[7] !== '') ultima = PRIMERA_FILA + i;
    if (f[0] instanceof Date) {
      var k = dia_(f[0]) + '|' + String(f[3]).trim().toLowerCase();
      vistos[k] = (vistos[k] || 0) + 1;
    }
  });

  var nuevas = reservas.filter(function (r) {
    var k = r.llegada + '|' + r.nombre.trim().toLowerCase();
    if (vistos[k] > 0) { vistos[k]--; return false; }
    return true;
  });
  if (!nuevas.length) return;

  var filas = nuevas.map(function (r) {
    var p = r.llegada.split('-');
    return [new Date(+p[0], +p[1] - 1, +p[2]), r.unidad, 'Hospedaje', r.nombre, r.canal, r.noches, '', r.monto, 'Tarjeta'];
  });
  hoja.getRange(ultima + 1, 1, filas.length, 9).setValues(filas);
  hoja.getRange(ultima + 1, 1, filas.length, 1).setNumberFormat('dd/MM/yyyy');

  var total = nuevas.reduce(function (s, r) { return s + r.monto; }, 0);
  var detalle = nuevas.map(function (r) {
    return r.llegada + ' · ' + r.unidad + ' · ' + r.nombre + ' · ' + r.canal + ' · ' + r.noches + ' noches · $' + r.monto.toFixed(2);
  }).join('\n');
  MailApp.sendEmail(CORREO, 'Ingresos Diarios actualizado: ' + nuevas.length + ' reservas nuevas',
    'Se agregaron ' + nuevas.length + ' reservas (total $' + total.toFixed(2) + ') a "' +
    SpreadsheetApp.getActive().getName() + '".\n\n' + detalle);
}
