var SPREADSHEET_ID = '10ZXj-EvXKgBZMdkeMWhVy_wpTlWr0mDxmzBKCxZPXEs';
var SHEET_NAME = 'Sheet1';
var HEADERS = ['Timestamp', 'Nama Tamu', 'Ucapan', 'Konfirmasi Kehadiran', 'Jumlah Tamu'];

function getSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#6D8998')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getParams_(e) {
  if (!e) return {};

  if (e.parameter && Object.keys(e.parameter).length > 0) {
    return e.parameter;
  }

  var params = {};

  if (e.postData && e.postData.contents) {
    var contents = e.postData.contents;

    try {
      var parsed = JSON.parse(contents);
      if (parsed && typeof parsed === 'object') return parsed;
    } catch (ignore) {}

    contents.split('&').forEach(function (pair) {
      var parts = pair.split('=');
      if (!parts[0]) return;
      params[decodeURIComponent(parts[0].replace(/\+/g, ' '))] =
        decodeURIComponent((parts[1] || '').replace(/\+/g, ' '));
    });
  }

  return params;
}

function clean_(value) {
  return String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
}

function getWishes_() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var wishes = [];

  for (var i = values.length - 1; i >= 0; i--) {
    var row = values[i];
    var nama = clean_(row[1]);
    var ucapan = String(row[2] == null ? '' : row[2]).trim();

    if (!nama && !ucapan) continue;

    var timestamp = row[0];
    wishes.push({
      timestamp: timestamp instanceof Date ? timestamp.toISOString() : String(timestamp || ''),
      nama: nama,
      ucapan: ucapan,
      kehadiran: clean_(row[3]),
      jumlah_tamu: Number(row[4]) || 0
    });
  }

  return wishes;
}

function doGet(e) {
  e = e || {};

  var action = e.parameter && e.parameter.action ? e.parameter.action : 'list';

  if (action === 'list') {
    var result = {
      status: 'success',
      data: getWishes_()
    };

    var callback = e.parameter && e.parameter.callback
      ? String(e.parameter.callback).replace(/[^\w.$]/g, '')
      : '';

    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + JSON.stringify(result) + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return json_(result);
  }

  return json_({
    status: 'error',
    message: 'Action tidak dikenal.'
  });
}

function doPost(e) {
  try {
    var params = getParams_(e);
    var nama = clean_(params.nama || params.name);
    var ucapan = String(params.ucapan || params.message || params.comment || '').trim();
    var kehadiran = clean_(params.kehadiran || params.attendance);
    var jumlahTamu = parseInt(
      String(params.jumlah_tamu || params.guest || '0').replace(/\D/g, ''),
      10
    );

    if (!nama) {
      return json_({ status: 'error', message: 'Nama tamu wajib diisi.' });
    }

    if (ucapan.length < 2) {
      return json_({ status: 'error', message: 'Ucapan minimal 2 karakter.' });
    }

    if (nama.length > 100) nama = nama.substring(0, 100);
    if (ucapan.length > 1000) ucapan = ucapan.substring(0, 1000);
    if (!kehadiran) kehadiran = 'Belum Konfirmasi';

    if (isNaN(jumlahTamu) || jumlahTamu < 0) jumlahTamu = 0;

    getSheet_().appendRow([new Date(), nama, ucapan, kehadiran, jumlahTamu]);

    return json_({
      status: 'success',
      message: 'Ucapan berhasil disimpan.'
    });
  } catch (error) {
    return json_({
      status: 'error',
      message: error && error.message ? error.message : 'Terjadi kesalahan.'
    });
  }
}
