/**
 * populateDylanU9()
 *
 * Reads DYLAN_TOURNAMENTS + U9_rankings, computes all values,
 * and writes them into DYLAN_U9.
 *
 * DYLAN_U9 has only cols A-G (no I-R Grade 5 section).
 * All 9U BS tournaments populate cols A-G regardless of grade.
 * Points lookup: U9_rankings col A = name, col F = points (higher = better).
 * Sort F/G descending by points.
 */

var BLOCKS_DYLANU9 = [
  { titleRow: 1,   firstData: 5,   lastData: 64  },
  { titleRow: 73,  firstData: 77,  lastData: 136 },
  { titleRow: 145, firstData: 149, lastData: 208 },
  { titleRow: 217, firstData: 221, lastData: 280 },
  { titleRow: 289, firstData: 293, lastData: 352 },
];

function populateDylanU9() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var dtSheet  = ss.getSheetByName('DYLAN_TOURNAMENTS');
  var rkSheet  = ss.getSheetByName('U9_rankings');
  var u9Sheet  = ss.getSheetByName('DYLAN_U9');

  if (!dtSheet) { Logger.log("Sheet 'DYLAN_TOURNAMENTS' not found"); return; }
  if (!rkSheet) { Logger.log("Sheet 'U9_rankings' not found"); return; }
  if (!u9Sheet) { Logger.log("Sheet 'DYLAN_U9' not found"); return; }

  var ptsMap      = buildPointsMap_DylanU9(rkSheet);
  var tournaments = readTournamentsDylan(dtSheet, '9U BS');

  Logger.log('9U BS tournaments: ' + tournaments.length);

  clearDylanU9(u9Sheet);

  for (var i = 0; i < BLOCKS_DYLANU9.length; i++) {
    populateLeftBlockDylan(u9Sheet, BLOCKS_DYLANU9[i], tournaments[i] || null, ptsMap);
  }

  Logger.log('Dylan. ' + tournaments.length + ' 9U BS tournament(s) written.');
}

// ─── Read tournaments ─────────────────────────────────────────────────────────
// 3-col stride: label col, value col, empty separator
function readTournamentsDylan(dtSheet, eventType) {
  var lastCol = dtSheet.getLastColumn();
  var lastRow = dtSheet.getLastRow();
  if (lastCol < 2 || lastRow < 6) return [];

  var data = dtSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var tournaments = [];

  for (var lc = 0; lc + 1 < data[0].length; lc += 3) {
    var vc = lc + 1;
    if (String(data[2][vc] || '').trim() !== eventType) continue;

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

// ─── Points lookup: col A = name, col F = points ─────────────────────────────
function buildPointsMap_DylanU9(rkSheet) {
  var lastRow = rkSheet.getLastRow();
  var data = rkSheet.getRange(2, 1, lastRow - 1, 6).getValues(); // cols A-F, skip header
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][0] || '').trim(); // col A
    var pts  = data[i][5];                       // col F
    if (name) map[name.toLowerCase()] = (pts !== null && pts !== '') ? String(pts) : '0';
  }
  return map;
}

function lookupPoints_DylanU9(name, ptsMap) {
  if (!name) return '0';
  return ptsMap[String(name).trim().toLowerCase()] || '0';
}

function pointsAsNumber_DylanU9(name, ptsMap) {
  return parseFloat(lookupPoints_DylanU9(name, ptsMap)) || 0;
}

// ─── Clear data areas ─────────────────────────────────────────────────────────
function clearDylanU9(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    sheet.getRange(1, 1, lastRow, 7).clearContent();
  }
}

// ─── Populate block: cols A-G only ───────────────────────────────────────────
// Col A/B:  entries in source order + points
// Col E:    ranking order numbers 1..drawSize
// Col F/G:  top drawSize sorted by points DESCENDING (higher = better)
// Title:    B1 = draw size, A2 = name, A3 = date
function populateLeftBlockDylan(sheet, block, tourn, ptsMap) {
  if (!tourn) return;

  var titleRow  = block.titleRow;
  var firstData = block.firstData;
  var maxRows   = block.lastData - block.firstData + 1;
  var drawSize  = Math.min(tourn.drawSize, maxRows);
  var entries   = tourn.entries;

  // Title rows
  sheet.getRange(titleRow,     2).setValue(tourn.drawSize); // B1: draw size
  sheet.getRange(titleRow + 1, 1).setValue(tourn.name);     // A2: name
  sheet.getRange(titleRow + 2, 1).setValue(tourn.dateStr);  // A3: date

  // Cols A/B: source order + points
  var colA = [], colB = [];
  for (var i = 0; i < maxRows; i++) {
    var name = entries[i] ? entries[i].name : '';
    colA.push([name]);
    colB.push([name ? lookupPoints_DylanU9(name, ptsMap) : '']);
  }
  sheet.getRange(firstData, 1, maxRows, 1).setValues(colA);
  sheet.getRange(firstData, 2, maxRows, 1).setValues(colB);

  // Col E: 1..drawSize
  var colE = [];
  for (var i = 0; i < maxRows; i++) {
    colE.push([i < drawSize ? i + 1 : '']);
  }
  sheet.getRange(firstData, 5, maxRows, 1).setValues(colE);

  // Cols F/G: sorted by points DESCENDING, capped at drawSize
  var withPts = [], zeroPts = [];
  for (var i = 0; i < entries.length; i++) {
    var name = entries[i].name;
    var pts  = pointsAsNumber_DylanU9(name, ptsMap);
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
    colG.push([name ? lookupPoints_DylanU9(name, ptsMap) : '']);
  }
  var actualF = colF.length;
  if (actualF > 0) {
    sheet.getRange(firstData, 6, actualF, 1).setValues(colF);
    sheet.getRange(firstData, 7, actualF, 1).setValues(colG);
  }
}