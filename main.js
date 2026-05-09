const SHEET_ID = '1hJnmYf7dV7J1DD4EkV1M18LonrVwRwMKz1lVT5MUnKI';
const SHEET_NAME = 'tournaments';

function doGet(e) {
  var params = (e && e.parameter) || {};
  var mode = params.mode || '';
  if (mode === 'public') {
    var t = HtmlService.createTemplateFromFile('public');
    t.gasPage = mode;
    t.gasId = params.id || '';
    t.gasAction = params.action || '';
    return t.evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('バドミントン大会エントリー');
  }
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
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = getSheet();
    var payload = JSON.parse(payloadJson);
    var lastLoadTime = payload.lastLoadTime || '';
    delete payload.lastLoadTime;
    var now = new Date();

    if (id) {
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0] === id) {
          try {
            var serverPayload = JSON.parse(rows[i][4]) || {};
            mergeNewServerEntries(payload, serverPayload, lastLoadTime);
          } catch(err) {}
          break;
        }
      }
    }

    var name = (payload.cfg && payload.cfg.tournamentName) || '名前なし';
    var summary = ((payload.cfg && payload.cfg.divisions) || []).map(function(d) {
      return d.name + ':' + d.leagueCount + 'リーグ×' + d.teamsPerLeague + 'チーム';
    }).join(' / ');
    var fullJson = JSON.stringify(payload);

    if (!id) {
      id = Utilities.getUuid();
      sheet.appendRow([id, name, summary, now, fullJson]);
    } else {
      var rows2 = sheet.getDataRange().getValues();
      for (var i = 1; i < rows2.length; i++) {
        if (rows2[i][0] === id) {
          sheet.getRange(i + 1, 2).setValue(name);
          sheet.getRange(i + 1, 3).setValue(summary);
          sheet.getRange(i + 1, 4).setValue(now);
          sheet.getRange(i + 1, 5).setValue(fullJson);
          break;
        }
      }
    }
    return id;
  } finally {
    lock.releaseLock();
  }
}

function mergeNewServerEntries(clientPayload, serverPayload, lastLoadTime) {
  var serverEntries = (serverPayload && serverPayload.entries) || [];
  if (serverEntries.length === 0) return;
  var threshold = lastLoadTime ? new Date(lastLoadTime).getTime() : 0;
  var clientEntries = clientPayload.entries || [];
  var clientKeys = {};
  clientEntries.forEach(function(e) {
    if (e && e.submittedAt) clientKeys[e.submittedAt + '|' + (e.teamName || '')] = true;
  });
  var newEntries = serverEntries.filter(function(e) {
    if (!e || !e.submittedAt) return false;
    if (threshold) {
      var t = new Date(e.submittedAt).getTime();
      if (isNaN(t) || t <= threshold) return false;
    }
    var k = e.submittedAt + '|' + (e.teamName || '');
    return !clientKeys[k];
  });
  if (newEntries.length === 0) return;
  newEntries.forEach(normalizeEntry);
  clientPayload.entries = clientEntries.concat(newEntries);
  var divs = (clientPayload.cfg && clientPayload.cfg.divisions) || [];
  var alpha = 'ABCDEFGHIJKLMNOPQRST';
  newEntries.forEach(function(e) {
    if (!e.placedSlot) return;
    var s = e.placedSlot;
    if (!clientPayload.data || !clientPayload.data[s.di] || !clientPayload.data[s.di][s.li]) return;
    var lg = clientPayload.data[s.di][s.li];
    if (!lg.teams || s.ti >= lg.teams.length) return;
    var divName = (divs[s.di] && divs[s.di].name) || '';
    var defaultName = divName + alpha.charAt(s.li) + (s.ti + 1);
    if (lg.teams[s.ti] === defaultName) {
      lg.teams[s.ti] = e.teamName;
    }
  });
}

function loadTournament(id) {
  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      try {
        var p = JSON.parse(rows[i][4]);
        if (p && p.entries) p.entries.forEach(normalizeEntry);
        return p;
      } catch(e) { return null; }
    }
  }
  return null;
}

function getPublicTournamentInfo(id) {
  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      try {
        var payload = JSON.parse(rows[i][4]) || {};
        var cfg = payload.cfg || {};
        return {
          id: id,
          name: cfg.tournamentName || rows[i][1] || '名前なし',
          membersPerTeam: cfg.membersPerTeam || 4,
          entryCount: (payload.entries || []).length
        };
      } catch(err) { return null; }
    }
  }
  return null;
}

function submitEntry(id, entryJson) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = getSheet();
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        var payload;
        try { payload = JSON.parse(rows[i][4]) || {}; } catch(err) { payload = {}; }
        if (!payload.entries) payload.entries = [];
        payload.entries.forEach(normalizeEntry);
        var entry = JSON.parse(entryJson);
        var email = (entry.email || '').trim();
        if (!email) return {ok: false, error: 'メールアドレスは必須です'};
        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return {ok: false, error: 'メールアドレスの形式が正しくありません'};
        var emailKey = email.toLowerCase();
        var dup = payload.entries.some(function(e) {
          return e && e.email && String(e.email).toLowerCase().trim() === emailKey;
        });
        if (dup) return {ok: false, error: 'このメールアドレスは既にエントリー済みです'};
        if (!entry.teamName || !String(entry.teamName).trim()) return {ok: false, error: 'チーム名は必須です'};
        entry.email = email;
        entry.teamName = String(entry.teamName).trim();
        entry.submittedAt = new Date().toISOString();
        entry.status = 'pending';
        payload.entries.push(entry);
        var now = new Date();
        var name = (payload.cfg && payload.cfg.tournamentName) || rows[i][1] || '名前なし';
        var summary = (payload.cfg && payload.cfg.divisions) ? payload.cfg.divisions.map(function(d) {
          return d.name + ':' + d.leagueCount + 'リーグ×' + d.teamsPerLeague + 'チーム';
        }).join(' / ') : (rows[i][2] || '');
        sheet.getRange(i + 1, 2).setValue(name);
        sheet.getRange(i + 1, 3).setValue(summary);
        sheet.getRange(i + 1, 4).setValue(now);
        sheet.getRange(i + 1, 5).setValue(JSON.stringify(payload));
        appendEntryToEntriesSheet(id, name, entry);
        return {ok: true, status: 'pending'};
      }
    }
    return {ok: false, error: '大会が見つかりません'};
  } finally {
    lock.releaseLock();
  }
}

function checkEntryStatus(id, email) {
  var emailKey = (email || '').toLowerCase().trim();
  if (!emailKey) return {ok: false, error: 'メールアドレスを入力してください'};
  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      try {
        var payload = JSON.parse(rows[i][4]) || {};
        var entries = payload.entries || [];
        var found = null;
        for (var k = 0; k < entries.length; k++) {
          var e = entries[k];
          if (e && e.email && String(e.email).toLowerCase().trim() === emailKey) { found = e; break; }
        }
        var tournamentName = (payload.cfg && payload.cfg.tournamentName) || rows[i][1] || '名前なし';
        if (!found) return {ok: true, status: 'not_found', tournamentName: tournamentName};
        normalizeEntry(found);
        return {
          ok: true,
          status: found.status || 'pending',
          teamName: found.teamName || '',
          placedLabel: found.placedLabel || '',
          submittedAt: found.submittedAt || '',
          tournamentName: tournamentName,
          memberCount: (found.members || []).filter(function(m){return m && m.name;}).length
        };
      } catch(err) { return {ok: false, error: 'データ読み込みエラー'}; }
    }
  }
  return {ok: false, error: '大会が見つかりません'};
}

function setEntryStatus(tournamentId, submittedAt, newStatus) {
  if (newStatus !== 'approved' && newStatus !== 'rejected' && newStatus !== 'pending') {
    return {ok: false, error: '不正なステータスです'};
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = getSheet();
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === tournamentId) {
        var payload;
        try { payload = JSON.parse(rows[i][4]) || {}; } catch(err) { return {ok: false, error: 'データエラー'}; }
        if (!payload.entries) payload.entries = [];
        payload.entries.forEach(normalizeEntry);
        var entry = null;
        for (var k = 0; k < payload.entries.length; k++) {
          if (payload.entries[k] && payload.entries[k].submittedAt === submittedAt) { entry = payload.entries[k]; break; }
        }
        if (!entry) return {ok: false, error: 'エントリーが見つかりません'};
        var prevPlacedSlot = entry.placedSlot ? {di: entry.placedSlot.di, li: entry.placedSlot.li, ti: entry.placedSlot.ti} : null;
        if (newStatus === 'approved') {
          if (!entry.placedSlot) {
            var placement = placeTeamInLeague(payload, entry.teamName);
            if (placement) {
              entry.placedSlot = {di: placement.di, li: placement.li, ti: placement.ti};
              entry.placedLabel = placement.label;
            }
          }
        } else {
          if (entry.placedSlot) {
            unplaceTeamInLeague(payload, entry.placedSlot, entry.teamName);
            delete entry.placedSlot;
            delete entry.placedLabel;
          }
        }
        entry.status = newStatus;
        var now = new Date();
        sheet.getRange(i + 1, 4).setValue(now);
        sheet.getRange(i + 1, 5).setValue(JSON.stringify(payload));
        updateEntryStateInSheet(tournamentId, entry);
        return {ok: true, entry: entry, prevPlacedSlot: prevPlacedSlot};
      }
    }
    return {ok: false, error: '大会が見つかりません'};
  } finally {
    lock.releaseLock();
  }
}

function unplaceTeamInLeague(payload, slot, teamName) {
  if (!payload.data || !payload.data[slot.di] || !payload.data[slot.di][slot.li]) return;
  var lg = payload.data[slot.di][slot.li];
  if (!lg.teams || slot.ti >= lg.teams.length) return;
  var divs = (payload.cfg && payload.cfg.divisions) || [];
  var alpha = 'ABCDEFGHIJKLMNOPQRST';
  var divName = (divs[slot.di] && divs[slot.di].name) || ((slot.di + 1) + '部');
  if (lg.teams[slot.ti] === teamName) {
    lg.teams[slot.ti] = divName + alpha.charAt(slot.li) + (slot.ti + 1);
  }
}

function normalizeEntry(e) {
  if (!e) return e;
  if (!e.status) {
    if (e.placedSlot) e.status = 'approved';
    else if (e.submittedAt) e.status = 'pending';
    else e.status = 'approved';
  }
  return e;
}

function statusLabel(status) {
  if (status === 'approved') return '承認済み';
  if (status === 'rejected') return '却下';
  return '承認待ち';
}

function placeTeamInLeague(payload, teamName) {
  if (!payload.data || !payload.data.length) return null;
  var divs = (payload.cfg && payload.cfg.divisions) || [];
  var alpha = 'ABCDEFGHIJKLMNOPQRST';
  for (var di = 0; di < payload.data.length; di++) {
    var divLeagues = payload.data[di];
    if (!divLeagues || !divLeagues.length) continue;
    for (var li = 0; li < divLeagues.length; li++) {
      var lg = divLeagues[li];
      if (!lg || !lg.teams) continue;
      var divName = (divs[di] && divs[di].name) || ((di + 1) + '部');
      for (var ti = 0; ti < lg.teams.length; ti++) {
        var defaultName = divName + alpha.charAt(li) + (ti + 1);
        if (lg.teams[ti] === defaultName || !lg.teams[ti]) {
          lg.teams[ti] = teamName;
          return {di: di, li: li, ti: ti, label: divName + ' ' + alpha.charAt(li) + 'リーグ ' + (ti + 1) + '番'};
        }
      }
    }
  }
  return null;
}

function getEntriesSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('entries');
  var newHeaders = ['受付日時', '大会ID', '大会名', 'チーム名', 'メール', 'メンバー', '状態', '配置先'];
  if (sheet) {
    var lastCol = sheet.getLastColumn();
    var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    var matches = lastCol >= 8 && headers[4] === 'メール' && headers[6] === '状態';
    if (!matches) {
      sheet.setName('entries_old_' + new Date().getTime());
      sheet = null;
    }
  }
  if (!sheet) {
    sheet = ss.insertSheet('entries');
    sheet.appendRow(newHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, newHeaders.length)
      .setBackground('#263238')
      .setFontColor('#00e5a0')
      .setFontWeight('bold');
    [150, 280, 200, 180, 220, 320, 100, 180].forEach(function(w, i) {
      sheet.setColumnWidth(i + 1, w);
    });
  }
  return sheet;
}

function appendEntryToEntriesSheet(tournamentId, tournamentName, entry) {
  var sheet = getEntriesSheet();
  var membersStr = ((entry && entry.members) || [])
    .filter(function(m) { return m && m.name; })
    .map(function(m) {
      var ageStr = (m.age !== '' && m.age != null && m.age !== 0) ? ' (' + m.age + '歳)' : '';
      return m.name + ageStr;
    })
    .join('\n');
  sheet.appendRow([
    new Date(),
    tournamentId,
    tournamentName,
    entry.teamName || '',
    entry.email || '',
    membersStr,
    statusLabel(entry.status),
    entry.placedLabel || ''
  ]);
}

function updateEntryStateInSheet(tournamentId, entry) {
  if (!entry || !entry.email) return;
  var sheet = getEntriesSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var emailKey = String(entry.email).toLowerCase().trim();
  for (var i = data.length - 1; i >= 0; i--) {
    var r = data[i];
    if (r[1] === tournamentId && String(r[4] || '').toLowerCase().trim() === emailKey) {
      sheet.getRange(i + 2, 7).setValue(statusLabel(entry.status));
      sheet.getRange(i + 2, 8).setValue(entry.placedLabel || '');
      return;
    }
  }
}