const SHEET_ID = '1hJnmYf7dV7J1DD4EkV1M18LonrVwRwMKz1lVT5MUnKI';
const SHEET_NAME = 'tournaments';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('バドミントン大会管理');
}

function getSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('tournaments');
  if (!sheet) {
    sheet = ss.insertSheet('tournaments');
    sheet.appendRow(['id', 'name', 'summary', 'updatedAt', 'payload']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:E1')
      .setBackground('#263238')
      .setFontColor('#00e5a0')
      .setFontWeight('bold');
    sheet.setColumnWidth(1, 280);  // id
    sheet.setColumnWidth(2, 200);  // name
    sheet.setColumnWidth(3, 300);  // summary
    sheet.setColumnWidth(4, 150);  // updatedAt
    sheet.setColumnWidth(5, 600);  // payload
  }
  return sheet;
}

function loadTournamentList() {
  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    list.push({
      id: r[0],
      name: r[1] || '名前なし',
      summary: r[2] || '',
      updatedAt: r[3] ? Utilities.formatDate(new Date(r[3]), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : ''
    });
  }
  return list.reverse();
}

function saveTournament(id, payloadJson) {
  var sheet = getSheet();
  var payload = JSON.parse(payloadJson);
  var now = new Date();
  var name = payload.cfg.tournamentName || '名前なし';
  var summary = payload.cfg.divisions.map(function(d) {
    return d.name + ':' + d.leagueCount + 'リーグ×' + d.teamsPerLeague + 'チーム';
  }).join(' / ');

  if (!id) {
    id = Utilities.getUuid();
    sheet.appendRow([id, name, summary, now, payloadJson]);
  } else {
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        sheet.getRange(i + 1, 2).setValue(name);
        sheet.getRange(i + 1, 3).setValue(summary);
        sheet.getRange(i + 1, 4).setValue(now);
        sheet.getRange(i + 1, 5).setValue(payloadJson);
        break;
      }
    }
  }
  return id;
}

function loadTournament(id) {
  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      try { return JSON.parse(rows[i][4]); } catch(e) { return null; }
    }
  }
  return null;
}