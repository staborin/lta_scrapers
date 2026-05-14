/**
 * populateDylanU10()
 *
 * Reads DYLAN_TOURNAMENTS + U10_rankings, computes all values,
 * and writes them into DYLAN_U10.
 *
 * DYLAN_U10 has both A-G (non-Grade 5) and I-R (Grade 5) sections.
 * Points lookup: U10_rankings col A = name, col F = points (higher = better).
 * Sort F/G descending. Q/R: top drawSize from M sorted by points descending.
 */

var BLOCKS_DYLANU10 = [
  { titleRow: 1,   firstData: 5,   lastData: 64  },
  { titleRow: 73,  firstData: 77,  lastData: 136 },
  { titleRow: 145, firstData: 149, lastData: 208 },
  { titleRow: 217, firstData: 221, lastData: 280 },
  { titleRow: 289, firstData: 293, lastData: 352 },
];

function populateDylanU10() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var dtSheet  = ss.getSheetByName('DYLAN_TOURNAMENTS');
  var rkSheet  = ss.getSheetByName('U10_rankings');
  var u10Sheet = ss.getSheetByName('DYLAN_U10');

  if (!dtSheet)  { Logger.log("Sheet 'DYLAN_TOURNAMENTS' not found"); return; }
  if (!rkSheet)  { Logger.log("Sheet 'U10_rankings' not found"); return; }
  if (!u10Sheet) { Logger.log("Sheet 'DYLAN_U10' not found"); return; }

  var ptsMap      = buildPointsMapU10_DylanU10(rkSheet);
  var tournaments = readTournamentsDylanU10(dtSheet);

  var grade5    = tournaments.filter(function(t) { return t.grade === 'Grade 5'; });
  var nonGrade5 = tournaments.filter(function(t) { return t.grade !== 'Grade 5'; });

  Logger.log('Grade 5: ' + grade5.length + ', Non-Grade 5: ' + nonGrade5.length);

  clearDylanU10(u10Sheet);

  for (var i = 0; i < BLOCKS_DYLANU10.length; i++) {
    populateRightBlockU10(u10Sheet, BLOCKS_DYLANU10[i], grade5[i]    || null, ptsMap);
    populateLeftBlockU10 (u10Sheet, BLOCKS_DYLANU10[i], nonGrade5[i] || null, ptsMap);
  }

  Logger.log(
    'Dylan 10U - Grade 5: ' + grade5.length + ', Non-Grade 5: ' + nonGrade5.length + ' tournament(s) written.'
  );
}

function readTournamentsDylanU10(dtSheet) {
  var lastCol = dtSheet.getLastColumn();
  var lastRow = dtSheet.getLastRow();
  if (lastCol < 2 || lastRow < 6) return [];

  var data = dtSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var tournaments = [];

  for (var lc = 0; lc + 1 < data[0].length; lc += 3) {
    var vc = lc + 1;
    if (String(data[2][vc] || '').trim() !== '10U BS') continue;

    var rawName  = String(data[0][lc] || '').trim();
    var tournName = rawName.replace(/^Tournament:\s*/i, '');
    var dateStr  = String(data[1][vc] || '').trim();
    var grade    = String(data[4][vc] || '').trim();
    var drawSize = parseInt(data[5][vc], 10) || 0;

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

function buildPointsMapU10_DylanU10(rkSheet) {
  var lastRow = rkSheet.getLastRow();
  var data = rkSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][0] || '').trim();
    var pts  = data[i][5];
    if (name) map[name.toLowerCase()] = (pts !== null && pts !== '') ? String(pts) : '0';
  }
  return map;
}

function lookupPointsU10_DylanU10(name, ptsMap) {
  if (!name) return '0';
  return ptsMap[String(name).trim().toLowerCase()] || '0';
}

function pointsAsNumberU10_DylanU10(name, ptsMap) {
  return parseFloat(lookupPointsU10_DylanU10(name, ptsMap)) || 0;
}

function parseEntryDateDylan_DylanU10(str) {
  if (!str) return 0;
  var m = str.match(/(\d{2})\/(\d{2})\/(\d{4})\s*(\d{1,2}):(\d{2})/);
  if (!m) return 0;
  return new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5]).getTime();
}

function writeAsTextU10_DylanU10(sheet, startRow, col, values) {
  var range = sheet.getRange(startRow, col, values.length, 1);
  range.setNumberFormat('@');
  range.setValues(values);
}

function clearDylanU10(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    sheet.getRange(1, 1, lastRow, 18).clearContent();
  }
}

// ─── RIGHT block: Grade 5 — cols I-R ─────────────────────────────────────────
// Same as LUKA but Q/R sorted by points descending
function populateRightBlockU10(sheet, block, tourn, ptsMap) {
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
  writeAsTextU10_DylanU10(sheet, firstData, 10, colJ);
  sheet.getRange(firstData, 12, maxRows, 1).setValues(colL);

  var sorted = entries.slice().sort(function(a, b) {
    return parseEntryDateDylan_DylanU10(a.dateStr) - parseEntryDateDylan_DylanU10(b.dateStr);
  });

  var colM = [], colN = [], colO = [];
  for (var i = 0; i < maxRows; i++) {
    var e    = sorted[i] || null;
    var name = e ? e.name    : '';
    var date = e ? e.dateStr : '';
    colM.push([name]);
    colN.push([date]);
    colO.push([name ? lookupPointsU10_DylanU10(name, ptsMap) : '']);
  }
  sheet.getRange(firstData, 13, maxRows, 1).setValues(colM);
  writeAsTextU10_DylanU10(sheet, firstData, 14, colN);
  sheet.getRange(firstData, 15, maxRows, 1).setValues(colO);

  // Q/R: top drawSize from M, sorted by points DESCENDING
  var topEntries = sorted.slice(0, drawSize);
  topEntries.sort(function(a, b) {
    return pointsAsNumberU10_DylanU10(b.name, ptsMap) - pointsAsNumberU10_DylanU10(a.name, ptsMap);
  });

  var colQ = [], colR = [];
  for (var i = 0; i < topEntries.length; i++) {
    var name = topEntries[i].name;
    colQ.push([name]);
    colR.push([lookupPointsU10_DylanU10(name, ptsMap)]);
  }
  var actualQ = colQ.length;
  if (actualQ > 0) {
    sheet.getRange(firstData, 17, actualQ, 1).setValues(colQ);
    sheet.getRange(firstData, 18, actualQ, 1).setValues(colR);
  }
}

// ─── LEFT block: non-Grade 5 — cols A-G ──────────────────────────────────────
function populateLeftBlockU10(sheet, block, tourn, ptsMap) {
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
    colB.push([name ? lookupPointsU10_DylanU10(name, ptsMap) : '']);
  }
  sheet.getRange(firstData, 1, maxRows, 1).setValues(colA);
  sheet.getRange(firstData, 2, maxRows, 1).setValues(colB);

  var colE = [];
  for (var i = 0; i < maxRows; i++) {
    colE.push([i < drawSize ? i + 1 : '']);
  }
  sheet.getRange(firstData, 5, maxRows, 1).setValues(colE);

  // F/G: sorted by points DESCENDING
  var withPts = [], zeroPts = [];
  for (var i = 0; i < entries.length; i++) {
    var name = entries[i].name;
    var pts  = pointsAsNumberU10_DylanU10(name, ptsMap);
    if (pts > 0) {
      withPts.push({ name: name, pts: pts });
    } else {
      zeroPts.push({ name: name });
    }
  }
  withPts.sort(function(a, b) { return b.pts - a.pts; }); // descending
  var draw = withPts.concat(zeroPts).slice(0, drawSize);

  var colF = [], colG = [];
  for (var i = 0; i < drawSize; i++) {
    var name = draw[i] ? draw[i].name : '';
    colF.push([name]);
    colG.push([name ? lookupPointsU10_DylanU10(name, ptsMap) : '']);
  }
  var actualF = colF.length;
  if (actualF > 0) {
    sheet.getRange(firstData, 6, actualF, 1).setValues(colF);
    sheet.getRange(firstData, 7, actualF, 1).setValues(colG);
  }
}