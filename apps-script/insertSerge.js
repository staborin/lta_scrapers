/**
 * populateSerge()
 *
 * Reads SERGE_TOURNAMENTS + Open_rankings, computes all values,
 * and writes them into SERGE.
 *
 * Each of the 5 blocks has two independent sub-sections:
 *   Cols A-G  → non-Grade 5 MS tournaments (draw sorted by WTN, lower = better)
 *   Cols I-R  → Grade 5 MS tournaments     (entries in source order for I/J,
 *                                            sorted by entry date for M/N/Q/R)
 *
 * Rankings use Open_rankings: col B = player name, col E = WTN Singles
 * WTN is a decimal — lower number = better (same sort direction as numeric rank)
 *
 * Run this whenever tournament data changes.
 */

var BLOCKS_SERGE = [
  { titleRow: 1,   firstData: 5,   lastData: 64  },
  { titleRow: 73,  firstData: 77,  lastData: 136 },
  { titleRow: 145, firstData: 149, lastData: 208 },
  { titleRow: 217, firstData: 221, lastData: 280 },
  { titleRow: 289, firstData: 293, lastData: 352 },
];

function populateSerge() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var stSheet    = ss.getSheetByName('SERGE_TOURNAMENTS');
  var rkSheet    = ss.getSheetByName('Open_rankings');
  var sergeSheet = ss.getSheetByName('SERGE');

  if (!stSheet)    { Logger.log("Sheet 'SERGE_TOURNAMENTS' not found"); return; }
  if (!rkSheet)    { Logger.log("Sheet 'Open_rankings' not found"); return; }
  if (!sergeSheet) { Logger.log("Sheet 'SERGE' not found"); return; }

  var wtnMap      = buildWtnMap_Serge(rkSheet);
  var nrSheet     = ss.getSheetByName('Non_ranked_WTN');
  if (nrSheet) {
    var nrMap = buildNonRankedWtnMap_Serge(nrSheet);
    for (var key in nrMap) {
      if (!wtnMap[key]) wtnMap[key] = nrMap[key];
    }
  }
  var tournaments = readTournaments_Serge(stSheet);

  var grade5    = tournaments.filter(function(t) { return t.grade === 'Grade 5'; });
  var nonGrade5 = tournaments.filter(function(t) { return t.grade !== 'Grade 5'; });

  Logger.log('Grade 5: ' + grade5.length + ', Non-Grade 5: ' + nonGrade5.length);

  clearSerge(sergeSheet);

  for (var i = 0; i < BLOCKS_SERGE.length; i++) {
    populateRightBlock_Serge(sergeSheet, BLOCKS_SERGE[i], grade5[i]    || null, wtnMap);
    populateLeftBlock_Serge (sergeSheet, BLOCKS_SERGE[i], nonGrade5[i] || null, wtnMap);
  }

  Logger.log(
    'Serge - Grade 5: ' + grade5.length + ', Non-Grade 5: ' + nonGrade5.length + ' tournament(s) written.'
  );
}

function readTournaments_Serge(stSheet) {
  var lastCol = stSheet.getLastColumn();
  var lastRow = stSheet.getLastRow();
  if (lastCol < 2 || lastRow < 6) return [];

  var data = stSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var tournaments = [];

  for (var lc = 0; lc + 1 < data[0].length; lc += 3) {
    var vc = lc + 1;

    var eventType = String(data[2][vc] || '').trim();
    if (!eventType.toUpperCase().startsWith('MS')) continue;

    var rawName   = String(data[0][lc] || '').trim();
    var tournName = rawName.replace(/^Tournament:\s*/i, '');
    var dateStr   = String(data[1][vc] || '').trim();
    var grade     = String(data[4][vc] || '').trim();
    var drawSize  = parseInt(data[5][vc], 10) || 0;

    var entries = [];
    for (var r = 9; r < data.length; r++) {
      var name = String(data[r][lc] || '').trim();
      if (!name || name === 'Entry Name') continue;
      entries.push({ name: name, dateStr: String(data[r][vc] || '').trim() });
    }

    tournaments.push({ name: tournName, dateStr: dateStr, grade: grade, drawSize: drawSize, entries: entries });
  }

  return tournaments;
}

function buildWtnMap_Serge(rkSheet) {
  var lastRow = rkSheet.getLastRow();
  var data = rkSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][1] || '').trim();
    var wtn  = data[i][4];
    if (name && wtn !== '' && wtn !== null) {
      map[name.toLowerCase()] = String(wtn);
    }
  }
  return map;
}

function buildNonRankedWtnMap_Serge(nrSheet) {
  var lastRow = nrSheet.getLastRow();
  if (lastRow < 2) return {};
  var data = nrSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][1] || '').trim();
    var wtn  = data[i][4];
    if (name && wtn !== '' && wtn !== null) {
      map[name.toLowerCase()] = String(wtn);
    }
  }
  return map;
}

function lookupWtn_Serge(name, wtnMap) {
  if (!name) return 'NO WTN';
  return wtnMap[String(name).trim().toLowerCase()] || 'NO WTN';
}

function wtnAsNumber_Serge(name, wtnMap) {
  var w = parseFloat(lookupWtn_Serge(name, wtnMap));
  return isNaN(w) ? 99999 : w;
}

function parseEntryDate_Serge(str) {
  if (!str) return 0;
  var m = str.match(/(\d{2})\/(\d{2})\/(\d{4})\s*(\d{1,2}):(\d{2})/);
  if (!m) return 0;
  return new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5]).getTime();
}

function clearSerge(sheet) {
  var ranges = [
    'A1:R3',     'A5:R64',
    'A73:R75',   'A77:R136',
    'A145:R147', 'A149:R208',
    'A217:R219', 'A221:R280',
    'A289:R291', 'A293:R352',
  ];
  for (var i = 0; i < ranges.length; i++) {
    sheet.getRange(ranges[i]).clearContent();
  }
}

function writeAsText_Serge(sheet, startRow, col, values) {
  var range = sheet.getRange(startRow, col, values.length, 1);
  range.setNumberFormat('@');
  range.setValues(values);
}

function populateRightBlock_Serge(sheet, block, tourn, wtnMap) {
  if (!tourn) return;

  var titleRow  = block.titleRow;
  var firstData = block.firstData;
  var maxRows   = block.lastData - block.firstData + 1;
  var drawSize  = Math.min(tourn.drawSize, maxRows);
  var entries   = tourn.entries;

  sheet.getRange(titleRow,     10).setValue(tourn.drawSize);
  sheet.getRange(titleRow,     11).setValue(tourn.grade);
  sheet.getRange(titleRow + 1, 9).setValue(tourn.name);
  sheet.getRange(titleRow + 2, 9).setValue(tourn.dateStr);

  var colI = [], colJ = [], colL = [];
  for (var i = 0; i < maxRows; i++) {
    var e = entries[i] || null;
    colI.push([e ? e.name    : '']);
    colJ.push([e ? e.dateStr : '']);
    colL.push([e ? i + 1    : '']);
  }
  sheet.getRange(firstData, 9,  maxRows, 1).setValues(colI);
  writeAsText_Serge(sheet, firstData, 10, colJ);
  sheet.getRange(firstData, 12, maxRows, 1).setValues(colL);

  var sorted = entries.slice().sort(function(a, b) {
    return parseEntryDate_Serge(a.dateStr) - parseEntryDate_Serge(b.dateStr);
  });

  var colM = [], colN = [], colO = [];
  for (var i = 0; i < maxRows; i++) {
    var e    = sorted[i] || null;
    var name = e ? e.name    : '';
    var date = e ? e.dateStr : '';
    colM.push([name]);
    colN.push([date]);
    colO.push([name ? lookupWtn_Serge(name, wtnMap) : '']);
  }
  sheet.getRange(firstData, 13, maxRows, 1).setValues(colM);
  writeAsText_Serge(sheet, firstData, 14, colN);
  sheet.getRange(firstData, 15, maxRows, 1).setValues(colO);

  var topEntries = sorted.slice(0, drawSize);
  topEntries.sort(function(a, b) {
    return wtnAsNumber_Serge(a.name, wtnMap) - wtnAsNumber_Serge(b.name, wtnMap);
  });

  var colQ = [], colR = [];
  for (var i = 0; i < topEntries.length; i++) {
    var name = topEntries[i].name;
    colQ.push([name]);
    colR.push([lookupWtn_Serge(name, wtnMap)]);
  }
  var actualQ = colQ.length;
  if (actualQ > 0) {
    sheet.getRange(firstData, 17, actualQ, 1).setValues(colQ);
    sheet.getRange(firstData, 18, actualQ, 1).setValues(colR);
  }
}

function populateLeftBlock_Serge(sheet, block, tourn, wtnMap) {
  if (!tourn) return;

  var titleRow  = block.titleRow;
  var firstData = block.firstData;
  var maxRows   = block.lastData - block.firstData + 1;
  var drawSize  = Math.min(tourn.drawSize, maxRows);
  var entries   = tourn.entries;

  sheet.getRange(titleRow,     2).setValue(tourn.drawSize);
  sheet.getRange(titleRow + 1, 1).setValue(tourn.name);
  sheet.getRange(titleRow + 2, 1).setValue(tourn.dateStr);

  var colA = [], colB = [];
  for (var i = 0; i < maxRows; i++) {
    var name = entries[i] ? entries[i].name : '';
    colA.push([name]);
    colB.push([name ? lookupWtn_Serge(name, wtnMap) : '']);
  }
  sheet.getRange(firstData, 1, maxRows, 1).setValues(colA);
  sheet.getRange(firstData, 2, maxRows, 1).setValues(colB);

  var colE = [];
  for (var i = 0; i < maxRows; i++) {
    colE.push([i < drawSize ? i + 1 : '']);
  }
  sheet.getRange(firstData, 5, maxRows, 1).setValues(colE);

  var ranked = [], unranked = [];
  for (var i = 0; i < entries.length; i++) {
    var name = entries[i].name;
    var w    = wtnAsNumber_Serge(name, wtnMap);
    if (w < 99999) {
      ranked.push({ name: name, wtn: w });
    } else {
      unranked.push({ name: name });
    }
  }
  ranked.sort(function(a, b) { return a.wtn - b.wtn; });
  var draw = ranked.concat(unranked).slice(0, drawSize);

  var colF = [], colG = [];
  for (var i = 0; i < drawSize; i++) {
    var name = draw[i] ? draw[i].name : '';
    colF.push([name]);
    colG.push([name ? lookupWtn_Serge(name, wtnMap) : '']);
  }
  var actualF = colF.length;
  if (actualF > 0) {
    sheet.getRange(firstData, 6, actualF, 1).setValues(colF);
    sheet.getRange(firstData, 7, actualF, 1).setValues(colG);
  }
}