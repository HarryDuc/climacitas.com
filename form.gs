// ============================================================
// Express MX Trámites - Google Apps Script webhook
// Form 4 truong: tema, nombre, correo, mensaje
// Ghi vao file Sheets, doc header dong 1 de biet thu tu cot
// ============================================================

// File chua du lieu landing page
var SHEET_IDS = [
  '1mwbsgcxsG7S0LD0l6DO2W7WftUX2JP4VLquGkq0FcPw' // climacitas.com
];

// Ten tab muon ghi (se doi chieu khong phan biet hoa thuong)
var TAB_NAME = 'data-contact';

// map header trong sheet -> ten field cua form
var FIELD_MAP = {
  timestamp: ['timestamp', 'fecha', 'time', 'hora', 'created'],
  tema:      ['tema', 'topic', 'asunto', 'motivo', 'tipo'],
  nombre:    ['nombre', 'name', 'nombre completo', 'fullname'],
  correo:    ['correo', 'email', 'e-mail', 'mail', 'correo electronico'],
  mensaje:   ['mensaje', 'message', 'msg', 'descripcion', 'detalle', 'consulta']
};

var DEFAULT_HEADERS = ['Tema', 'Nombre', 'Correo', 'Mensaje', 'Timestamp'];

function doGet(e) { return handleResponse(e); }
function doPost(e) { return handleResponse(e); }

function handleResponse(e) {
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    var out = [];

    for (var i = 0; i < SHEET_IDS.length; i++) {
      var res = writeTo(SHEET_IDS[i], p);
      out.push(res);
    }

    var ok = false;
    for (var j = 0; j < out.length; j++) if (out[j].ok) ok = true;

    if (!ok) {
      return json({ 'result': 'error', 'error': 'Khong ghi duoc file nao', 'detail': out });
    }
    return json({ 'result': 'success', 'detail': out });

  } catch (err) {
    return json({ 'result': 'error', 'error': String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Ghi 1 dong vao 1 file
function writeTo(spreadsheetId, p) {
  var info = { id: spreadsheetId, ok: false };

  var doc;
  try {
    doc = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    info.error = 'Khong mo duoc file: ' + e;
    Logger.log('ERR open ' + spreadsheetId + ' : ' + e);
    return info;
  }

  info.file = doc.getName();
  var allTabs = doc.getSheets();
  info.tabs = allTabs.map(function (s) { return s.getName(); });

  // Chon tab: uu tien ten khop, khong co thi dung tab dau tien
  var sheet = null;
  var want = String(TAB_NAME).trim().toLowerCase();
  for (var i = 0; i < allTabs.length; i++) {
    if (String(allTabs[i].getName()).trim().toLowerCase() === want) { sheet = allTabs[i]; break; }
  }
  if (!sheet) sheet = allTabs[0];

  info.tab = sheet.getName();
  info.gid = sheet.getSheetId();
  info.url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/edit#gid=' + sheet.getSheetId();

  // Neu sheet trong -> tao header
  if (sheet.getLastRow() === 0) writeHeader(sheet);

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h || '').trim(); });
  info.header = headers;

  // Map tung cot -> key cua form (keyOfCol[chiSoCot] = 'tema' | 'nombre' ...)
  var keyOfCol = [];
  var usedKeys = {};
  for (var c = 0; c < headers.length; c++) {
    var key = matchKey(headers[c]);
    if (key && !usedKeys[key]) {
      keyOfCol[c] = key;
      usedKeys[key] = true;
    } else {
      keyOfCol[c] = null;
    }
  }
  info.mapping = keyOfCol;

  // Dung dong du lieu
  var row = [];
  for (var j = 0; j < headers.length; j++) {
    var k = keyOfCol[j];
    if (k === 'timestamp') row.push(new Date());
    else if (k) row.push(p[k] || '');
    else row.push('');
  }
  info.row = row;

  var nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);

  info.writtenRow = nextRow;
  info.ok = true;

  Logger.log('WRITE -> file="' + info.file + '" tab="' + info.tab + '" row=' + nextRow + ' data=' + JSON.stringify(row));
  return info;
}

function normalize(s) {
  return String(s || '').toLowerCase().normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, ' ').trim();
}

function matchKey(header) {
  var h = normalize(header);
  if (!h) return null;
  for (var key in FIELD_MAP) {
    if (!FIELD_MAP.hasOwnProperty(key)) continue;
    var aliases = FIELD_MAP[key];
    for (var i = 0; i < aliases.length; i++) {
      var a = normalize(aliases[i]);
      if (h === a || h.indexOf(a) !== -1) return key;
    }
  }
  return null;
}

function writeHeader(sheet) {
  var range = sheet.getRange(1, 1, 1, DEFAULT_HEADERS.length);
  range.setValues([DEFAULT_HEADERS]);
  range.setFontWeight('bold');
  range.setFontSize(11);
  range.setBackground('#f1f3f4');
  range.setFontColor('#202124');
  range.setBorder(true, true, true, true, false, false);
  var widths = [130, 200, 220, 320, 140];
  for (var i = 0; i < widths.length; i++) sheet.setColumnWidth(i + 1, widths[i]);
  sheet.setFrozenRows(1);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============ HAM CHAY 1 LAN ============

// Xem het: file, tab, header, dong cuoi + link
function inspect() {
  var report = [];
  for (var i = 0; i < SHEET_IDS.length; i++) {
    var doc;
    try { doc = SpreadsheetApp.openById(SHEET_IDS[i]); } catch (e) {
      report.push({ id: SHEET_IDS[i], error: String(e) });
      continue;
    }
    var tabs = doc.getSheets();
    var info = [];
    for (var j = 0; j < tabs.length; j++) {
      var s = tabs[j];
      var last = s.getLastRow();
      var preview = [];
      if (last > 1) {
        preview = s.getRange(2, 1, Math.min(3, last - 1), Math.min(s.getLastColumn(), 5)).getValues();
      }
      info.push({
        tab: s.getName(),
        gid: s.getSheetId(),
        lastRow: last,
        header: s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0],
        preview: preview,
        url: 'https://docs.google.com/spreadsheets/d/' + SHEET_IDS[i] + '/edit#gid=' + s.getSheetId()
      });
    }
    report.push({ id: SHEET_IDS[i], file: doc.getName(), tabs: info });
  }
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

// Test ghi 1 dong gia - ket qua se hien o "Execution log"
function testSubmit() {
  var res = handleResponse({ parameter: {
    tema: 'Test', nombre: 'Nguyen Van A',
    correo: 'test@example.com', mensaje: 'Dong thu nghiem.'
  }});
  var body = res.getContent();
  Logger.log('RESULT: ' + body);
  return body;
}
