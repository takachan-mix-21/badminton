const SHEET_ID = '1hJnmYf7dV7J1DD4EkV1M18LonrVwRwMKz1lVT5MUnKI';
const SHEET_NAME = 'tournaments';

var FAVICON_URL = 'https://raw.githubusercontent.com/takachan-mix-21/badminton/main/6342.ico';

function doGet(e) {
  var params = (e && e.parameter) || {};
  var mode = params.mode || '';
  if (mode === 'admin') {
    return HtmlService.createHtmlOutputFromFile('index')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setFaviconUrl(FAVICON_URL)
      .setTitle('バドミントン大会管理');
  }
  var publicModes = {'':'', public:'', view:'view', status:'status', entry:''};
  var actionMap = (mode in publicModes) ? publicModes[mode] : '';
  var t = HtmlService.createTemplateFromFile('public');
  t.gasPage = 'public';
  t.gasId = params.id || '';
  t.gasAction = params.action || actionMap;
  t.gasUrl = ScriptApp.getService().getUrl();
  var title = (mode === 'view') ? 'バドミントン大会 リーグ表' :
              (mode === 'status') ? 'バドミントン大会 状況確認' : 'バドミントン大会エントリー';
  return t.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl(FAVICON_URL)
    .setTitle(title);
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

function exportTournamentToSheet(id) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = getSheet();
    var rows = sheet.getDataRange().getValues();
    var payload = null;
    var name = '';
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        try { payload = JSON.parse(rows[i][4]) || {}; } catch(err) { return {ok:false, error:'データ読み込みエラー'}; }
        name = (payload.cfg && payload.cfg.tournamentName) || rows[i][1] || '名前なし';
        break;
      }
    }
    if (!payload) return {ok:false, error:'大会が見つかりません'};
    var sheetName = sanitizeSheetName('📋 ' + name);
    var existing = ss.getSheetByName(sheetName);
    if (existing) ss.deleteSheet(existing);
    var s = ss.insertSheet(sheetName);
    writeTournamentReport(s, payload);
    return {ok:true, sheetName:sheetName, url: ss.getUrl() + '#gid=' + s.getSheetId()};
  } finally {
    lock.releaseLock();
  }
}

function sanitizeSheetName(s) {
  var n = String(s || '名前なし').replace(/[:\\\/\?\*\[\]]/g, '');
  if (n.length > 90) n = n.substring(0, 90);
  return n || '名前なし';
}

function writeTournamentReport(sheet, payload) {
  var cfg = payload.cfg || {};
  var data = payload.data || [];
  var entries = (payload.entries || []).map(normalizeEntry);
  var tournamentStates = payload.tournamentStates || [];
  var alpha = 'ABCDEFGHIJKLMNOPQRST';
  var row = 1;

  sheet.getRange(row, 1).setValue(cfg.tournamentName || '名前なし');
  sheet.getRange(row, 1).setFontSize(18).setFontWeight('bold');
  row += 2;

  sheet.getRange(row, 1, 2, 2).setValues([
    ['作成日時', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')],
    ['1チームのメンバー数', cfg.membersPerTeam || 4]
  ]);
  sheet.getRange(row, 1, 2, 1).setFontWeight('bold');
  row += 3;

  if (entries.length > 0) {
    sheet.getRange(row, 1).setValue('▼ エントリー一覧 (' + entries.length + '件)');
    sheet.getRange(row, 1).setFontSize(13).setFontWeight('bold').setBackground('#e8f5e9');
    sheet.getRange(row, 1, 1, 9).merge();
    row++;
    var entryHeaders = ['チーム名', '希望部門', '状態', '配置先', '代表者氏名', '電話番号', 'メール', '住所', 'メンバー'];
    sheet.getRange(row, 1, 1, entryHeaders.length).setValues([entryHeaders]).setBackground('#f0f0f0').setFontWeight('bold');
    row++;
    var entryRows = entries.map(function(e) {
      var members = (e.members || []).filter(function(m){return m && m.name;}).map(function(m) {
        var meta = [];
        if (m.gender) meta.push(m.gender);
        if (m.age) meta.push(m.age+'歳');
        return m.name + (meta.length ? '('+meta.join('・')+')' : '');
      }).join(', ');
      return [
        e.teamName || '',
        e.division || '',
        statusLabel(e.status),
        e.placedLabel || '',
        e.repName || '',
        e.phone || '',
        e.email || '',
        e.address || '',
        members
      ];
    });
    if (entryRows.length > 0) {
      sheet.getRange(row, 1, entryRows.length, 9).setValues(entryRows);
      for (var rr = 0; rr < entries.length; rr++) {
        var st = entries[rr].status || 'approved';
        var bg = st === 'approved' ? '#e8f5e9' : st === 'rejected' ? '#ffebee' : '#fffde7';
        sheet.getRange(row + rr, 3).setBackground(bg);
      }
      row += entryRows.length;
    }
    row += 2;
  }

  (cfg.divisions || []).forEach(function(div, di) {
    var divLeagues = data[di] || [];
    if (!divLeagues.length) return;
    sheet.getRange(row, 1).setValue('▼ ' + div.name);
    sheet.getRange(row, 1).setFontSize(13).setFontWeight('bold').setBackground('#e3f2fd');
    sheet.getRange(row, 1, 1, 9).merge();
    row += 2;

    divLeagues.forEach(function(lg, li) {
      sheet.getRange(row, 1).setValue('■ ' + div.name + ' ' + alpha.charAt(li) + 'リーグ');
      sheet.getRange(row, 1).setFontSize(12).setFontWeight('bold');
      row++;
      sheet.getRange(row, 1).setValue('順位表').setFontWeight('bold').setFontColor('#666');
      row++;
      var rankHeaders = ['順位', 'チーム', '総勝', 'MD', 'WD', 'XD', '得点', '失点', '差'];
      sheet.getRange(row, 1, 1, rankHeaders.length).setValues([rankHeaders]).setBackground('#f0f0f0').setFontWeight('bold');
      row++;
      var standings = computeStandingsServer(lg);
      var advCount = div.advanceCount || 0;
      var standingRows = standings.map(function(s, idx) {
        return [idx+1, s.name, s.totalW, s.mdW, s.wdW, s.xdW, s.pf, s.pa, (s.diff>=0?'+':'')+s.diff];
      });
      if (standingRows.length > 0) {
        sheet.getRange(row, 1, standingRows.length, 9).setValues(standingRows);
        for (var idx = 0; idx < standings.length; idx++) {
          if (idx < advCount) sheet.getRange(row + idx, 1, 1, 9).setBackground('#c8e6c9');
        }
        row += standingRows.length;
      }
      row++;

      var teams = lg.teams || [];
      if (teams.length > 0) {
        sheet.getRange(row, 1).setValue('対戦結果 (総勝数 A-B)').setFontWeight('bold').setFontColor('#666');
        row++;
        var headerRow = [''].concat(teams);
        sheet.getRange(row, 1, 1, headerRow.length).setValues([headerRow]).setBackground('#f0f0f0').setFontWeight('bold');
        row++;
        var matrixRows = teams.map(function(name, ri) {
          var rowData = [name];
          teams.forEach(function(_, ci) {
            if (ri === ci) { rowData.push('—'); return; }
            var match = getMatchAtServer(lg, ri, ci);
            if (!match) { rowData.push(''); return; }
            var wA = 0, wB = 0;
            ['md','wd','xd'].forEach(function(c) {
              var w = catWinSrv(match[c]);
              if (w === 'A') wA++; else if (w === 'B') wB++;
            });
            rowData.push(wA + '-' + wB);
          });
          return rowData;
        });
        sheet.getRange(row, 1, matrixRows.length, headerRow.length).setValues(matrixRows);
        sheet.getRange(row, 1, matrixRows.length, 1).setFontWeight('bold').setBackground('#f0f0f0');
        row += matrixRows.length;
      }
      row += 2;
    });

    var ts = tournamentStates[di];
    if (ts && ts.seeds && ts.seeds.length) {
      sheet.getRange(row, 1).setValue('■ ' + div.name + ' トーナメント');
      sheet.getRange(row, 1).setFontSize(12).setFontWeight('bold');
      row++;
      sheet.getRange(row, 1).setValue('シード順').setFontWeight('bold').setFontColor('#666');
      row++;
      sheet.getRange(row, 1, 1, 4).setValues([['順位','チーム','元のリーグ','成績']]).setBackground('#f0f0f0').setFontWeight('bold');
      row++;
      var seedRows = ts.seeds.filter(function(s){return s;}).map(function(s, i) {
        return [
          i+1,
          s.name || '',
          (s.leagueLabel || '') + 'リーグ ' + ((s.rankInLeague||0)+1) + '位',
          (s.totalW||0) + '勝 ' + ((s.diff||0)>=0?'+':'') + (s.diff||0)
        ];
      });
      if (seedRows.length > 0) {
        sheet.getRange(row, 1, seedRows.length, 4).setValues(seedRows);
        row += seedRows.length;
      }
      row++;

      var rounds = computeRoundsServer(ts.seeds, ts.matchResults || {});
      var totalRounds = rounds.length - 1;
      if (totalRounds > 0) {
        sheet.getRange(row, 1).setValue('試合結果').setFontWeight('bold').setFontColor('#666');
        row++;
        sheet.getRange(row, 1, 1, 4).setValues([['ラウンド','チームA','チームB','勝者']]).setBackground('#f0f0f0').setFontWeight('bold');
        row++;
        var roundNames = ['1回戦','2回戦','準々決勝','準決勝','決勝'];
        for (var ri = 0; ri < totalRounds; ri++) {
          var matchCount = rounds[ri].length / 2;
          var rName = roundNames[Math.max(0, roundNames.length - totalRounds + ri)] || ('第'+(ri+1)+'回戦');
          for (var mi = 0; mi < matchCount; mi++) {
            var teamA = rounds[ri][mi*2], teamB = rounds[ri][mi*2+1];
            var aName = teamA ? teamA.name : '(BYE)';
            var bName = teamB ? teamB.name : '(BYE)';
            var res = (ts.matchResults || {})[ri+'_'+mi];
            var winnerName = res && res.winnerName ? res.winnerName : (teamA && !teamB ? aName : !teamA && teamB ? bName : '');
            sheet.getRange(row, 1, 1, 4).setValues([[rName, aName, bName, winnerName]]);
            row++;
          }
        }
        row++;
        var champion = rounds[totalRounds] ? rounds[totalRounds][0] : null;
        if (champion) {
          sheet.getRange(row, 1).setValue('🏆 優勝');
          sheet.getRange(row, 1).setFontWeight('bold').setFontColor('#e65100');
          sheet.getRange(row, 2).setValue(champion.name).setFontWeight('bold').setFontColor('#e65100');
          row++;
        }
      }
      row += 2;
    }
  });

  sheet.setColumnWidth(1, 200);
  for (var c = 2; c <= 9; c++) sheet.setColumnWidth(c, 130);
  sheet.setFrozenRows(1);
}

function getMatchAtServer(lg, ri, ci) {
  if (!lg.results) return null;
  var key = ri < ci ? (ri + '_' + ci) : (ci + '_' + ri);
  if (!lg.results[key]) return null;
  var r = lg.results[key];
  if (ri < ci) return r;
  var flip = function(g){return {a:g.b, b:g.a};};
  return {md:(r.md||[]).map(flip), wd:(r.wd||[]).map(flip), xd:(r.xd||[]).map(flip)};
}

function catWinSrv(games) {
  var wA=0, wB=0;
  (games || []).forEach(function(g){if(g.a>g.b)wA++;else if(g.b>g.a)wB++;});
  return wA > wB ? 'A' : (wB > wA ? 'B' : null);
}

function computeStandingsServer(lg) {
  var teams = lg.teams || [];
  var standings = teams.map(function(name, ti) {
    var totalW=0, mdW=0, wdW=0, xdW=0, pf=0, pa=0;
    teams.forEach(function(_, oi) {
      if (oi === ti) return;
      var r = getMatchAtServer(lg, ti, oi);
      if (!r) return;
      ['md','wd','xd'].forEach(function(c) { if (catWinSrv(r[c]) === 'A') totalW++; });
      if (catWinSrv(r.md)==='A') mdW++;
      if (catWinSrv(r.wd)==='A') wdW++;
      if (catWinSrv(r.xd)==='A') xdW++;
      ['md','wd','xd'].forEach(function(c) {
        (r[c] || []).forEach(function(g) { pf += (g.a||0); pa += (g.b||0); });
      });
    });
    return {name:name, totalW:totalW, mdW:mdW, wdW:wdW, xdW:xdW, pf:pf, pa:pa, diff:pf-pa};
  });
  standings.sort(function(a,b){return b.totalW-a.totalW||b.diff-a.diff||b.pf-a.pf;});
  return standings;
}

function computeRoundsServer(seeds, matchResults) {
  if (!seeds || !seeds.length) return [];
  var size = 1; while (size < seeds.length) size *= 2;
  var slots = new Array(size).fill(null);
  seeds.forEach(function(s, i) { slots[i] = s; });
  var rounds = [slots.slice()];
  var current = slots.slice();
  var ri = 0;
  while (current.length > 1) {
    var next = [];
    for (var mi = 0; mi < current.length; mi += 2) {
      var a = current[mi], b = current[mi+1];
      if (a && !b) { next.push(a); continue; }
      if (!a && b) { next.push(b); continue; }
      var res = matchResults[ri+'_'+(mi/2)];
      var winner = null;
      if (res && res.winnerName) {
        for (var k = 0; k < seeds.length; k++) {
          if (seeds[k] && seeds[k].name === res.winnerName) { winner = seeds[k]; break; }
        }
      }
      next.push(winner);
    }
    rounds.push(next);
    current = next;
    ri++;
  }
  return rounds;
}

function deleteTournament(id) {
  if (!id) return {ok: false, error: 'IDが指定されていません'};
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = getSheet();
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        sheet.deleteRow(i + 1);
        return {ok: true};
      }
    }
    return {ok: false, error: '大会が見つかりません'};
  } finally {
    lock.releaseLock();
  }
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

function getPublicTournamentView(id) {
  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      try {
        var p = JSON.parse(rows[i][4]) || {};
        var rawEntries = (p.entries || []).map(normalizeEntry);
        var publicEntries = rawEntries
          .filter(function(e) { return (e.status||'pending') === 'approved' && e.teamName; })
          .map(function(e) {
            return {
              teamName: e.teamName,
              division: e.division || '',
              members: ((e.members) || []).filter(function(m){return m && m.name;}).map(function(m) {
                return {name: m.name, age: m.age, gender: m.gender || ''};
              })
            };
          });
        return {
          id: id,
          name: (p.cfg && p.cfg.tournamentName) || rows[i][1] || '名前なし',
          divisions: (p.cfg && p.cfg.divisions) || [],
          gamePoints: (p.cfg && p.cfg.gamePoints) || [11, 21, 21],
          data: p.data || [],
          tournamentStates: (p.tournamentStates || []).map(function(ts) {
            if (!ts) return null;
            return {seeds: ts.seeds || [], matchResults: ts.matchResults || {}};
          }),
          entries: publicEntries
        };
      } catch(err) { return null; }
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
          entryCount: (payload.entries || []).length,
          divisions: (cfg.divisions || []).map(function(d) { return {name: d.name}; })
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
        var divs = (payload.cfg && payload.cfg.divisions) || [];
        var division = (entry.division || '').trim();
        if (divs.length > 0) {
          var validNames = divs.map(function(d) { return d.name; });
          if (!division || validNames.indexOf(division) === -1) {
            return {ok: false, error: '希望部門を選択してください'};
          }
        }
        var repName = (entry.repName || '').trim();
        var address = (entry.address || '').trim();
        var phone = (entry.phone || '').trim();
        if (!repName) return {ok: false, error: '代表者氏名は必須です'};
        if (!address) return {ok: false, error: '住所は必須です'};
        if (!phone) return {ok: false, error: '電話番号は必須です'};
        if (!/^[\d\-+()\s]+$/.test(phone)) return {ok: false, error: '電話番号の形式が正しくありません'};
        entry.repName = repName;
        entry.address = address;
        entry.phone = phone;
        entry.division = division;
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
          division: found.division || '',
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
            var placement = placeTeamInLeague(payload, entry.teamName, entry.division);
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

function placeTeamInLeague(payload, teamName, preferredDivision) {
  if (!payload.data || !payload.data.length) return null;
  var divs = (payload.cfg && payload.cfg.divisions) || [];
  var alpha = 'ABCDEFGHIJKLMNOPQRST';
  var divIndices = [];
  if (preferredDivision) {
    for (var di = 0; di < divs.length; di++) {
      if (divs[di] && divs[di].name === preferredDivision) { divIndices.push(di); break; }
    }
  }
  if (divIndices.length === 0) {
    for (var di = 0; di < payload.data.length; di++) divIndices.push(di);
  }
  for (var k = 0; k < divIndices.length; k++) {
    var di = divIndices[k];
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
  var newHeaders = ['受付日時', '大会ID', '大会名', 'チーム名', 'メール', 'メンバー', '状態', '配置先', '希望部門', '代表者氏名', '住所', '電話番号'];
  if (sheet) {
    var lastCol = sheet.getLastColumn();
    var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    var hasFull = lastCol >= 12 && headers[8] === '希望部門' && headers[9] === '代表者氏名';
    var hasV2 = lastCol >= 9 && headers[8] === '希望部門';
    var hasV1 = lastCol >= 8 && headers[4] === 'メール' && headers[6] === '状態';
    if (hasFull) {
      // OK
    } else if (hasV2) {
      sheet.getRange(1, 10).setValue('代表者氏名').setBackground('#263238').setFontColor('#00e5a0').setFontWeight('bold');
      sheet.getRange(1, 11).setValue('住所').setBackground('#263238').setFontColor('#00e5a0').setFontWeight('bold');
      sheet.getRange(1, 12).setValue('電話番号').setBackground('#263238').setFontColor('#00e5a0').setFontWeight('bold');
      sheet.setColumnWidth(10, 160);
      sheet.setColumnWidth(11, 280);
      sheet.setColumnWidth(12, 140);
    } else if (hasV1) {
      sheet.getRange(1, 9).setValue('希望部門').setBackground('#263238').setFontColor('#00e5a0').setFontWeight('bold');
      sheet.getRange(1, 10).setValue('代表者氏名').setBackground('#263238').setFontColor('#00e5a0').setFontWeight('bold');
      sheet.getRange(1, 11).setValue('住所').setBackground('#263238').setFontColor('#00e5a0').setFontWeight('bold');
      sheet.getRange(1, 12).setValue('電話番号').setBackground('#263238').setFontColor('#00e5a0').setFontWeight('bold');
      sheet.setColumnWidth(9, 120);
      sheet.setColumnWidth(10, 160);
      sheet.setColumnWidth(11, 280);
      sheet.setColumnWidth(12, 140);
    } else {
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
    [150, 280, 200, 180, 220, 320, 100, 180, 120, 160, 280, 140].forEach(function(w, i) {
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
      var meta = [];
      if (m.gender) meta.push(m.gender);
      if (m.age !== '' && m.age != null && m.age !== 0) meta.push(m.age + '歳');
      return m.name + (meta.length ? ' (' + meta.join('・') + ')' : '');
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
    entry.placedLabel || '',
    entry.division || '',
    entry.repName || '',
    entry.address || '',
    entry.phone || ''
  ]);
}

function updateEntryStateInSheet(tournamentId, entry) {
  if (!entry || !entry.email) return;
  var sheet = getEntriesSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var emailKey = String(entry.email).toLowerCase().trim();
  for (var i = data.length - 1; i >= 0; i--) {
    var r = data[i];
    if (r[1] === tournamentId && String(r[4] || '').toLowerCase().trim() === emailKey) {
      sheet.getRange(i + 2, 7).setValue(statusLabel(entry.status));
      sheet.getRange(i + 2, 8).setValue(entry.placedLabel || '');
      if (lastCol >= 9) sheet.getRange(i + 2, 9).setValue(entry.division || '');
      if (lastCol >= 10) sheet.getRange(i + 2, 10).setValue(entry.repName || '');
      if (lastCol >= 11) sheet.getRange(i + 2, 11).setValue(entry.address || '');
      if (lastCol >= 12) sheet.getRange(i + 2, 12).setValue(entry.phone || '');
      return;
    }
  }
}