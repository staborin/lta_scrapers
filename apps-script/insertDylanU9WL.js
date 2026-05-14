/**
 * populateDylanU9_WL
 * Reads DYLAN_WATCHLIST + U9_rankings → writes DYLAN_U9_WL
 */

var BLOCKS_U9_WL = [
  { titleRow: 1,   firstData: 6,   lastData: 65  },
  { titleRow: 74,  firstData: 79,  lastData: 138 },
  { titleRow: 147, firstData: 152, lastData: 211 },
  { titleRow: 220, firstData: 225, lastData: 284 },
  { titleRow: 293, firstData: 298, lastData: 357 },
  { titleRow: 366, firstData: 371, lastData: 430 },
  { titleRow: 439, firstData: 444, lastData: 503 },
  { titleRow: 512, firstData: 517, lastData: 576 },
  { titleRow: 585, firstData: 590, lastData: 649 },
  { titleRow: 658, firstData: 663, lastData: 722 },
];

function populateDylanU9_WL() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wlSheet  = ss.getSheetByName('DYLAN_WATCHLIST');
  var rkSheet  = ss.getSheetByName('U9_rankings');
  var outSheet = ss.getSheetByName('DYLAN_U9_WL');
  if (!wlSheet)  { Logger.log("Sheet 'DYLAN_WATCHLIST' not found");  return; }
  if (!rkSheet)  { Logger.log("Sheet 'U9_rankings' not found");       return; }
  if (!outSheet) { Logger.log("Sheet 'DYLAN_U9_WL' not found");       return; }
  var ptsMap      = buildPtsMap_U9_WL(rkSheet);
  var tournaments = readTournaments_U9_WL(wlSheet);
  Logger.log('U9 Watchlist tournaments: ' + tournaments.length);
  clearSheet_U9_WL(outSheet);
  for (var i = 0; i < BLOCKS_U9_WL.length; i++) {
    populateBlock_U9_WL(outSheet, BLOCKS_U9_WL[i], tournaments[i] || null, ptsMap);
  }
  Logger.log('Dylan U9 Watchlist: ' + tournaments.length + ' tournament(s) written.');
}

function readTournaments_U9_WL(wlSheet) {
  var lastCol = wlSheet.getLastColumn();
  var lastRow = wlSheet.getLastRow();
  if (lastCol < 2 || lastRow < 6) return [];
  var data = wlSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var tournaments = [];
  for (var lc = 0; lc + 1 < data[0].length; lc += 3) {
    var vc = lc + 1;
    var rawName = String(data[0][lc] || '').trim();
    if (!rawName || !rawName.match(/^Tournament:/i)) continue;
    var event = String(data[2][vc] || '').trim();
    if (event !== '9U BS') continue;
    var tournName   = rawName.replace(/^Tournament:\s*/i, '');
    var dateStr     = String(data[1][vc] || '').trim();
    var closingDate = String(data[3][vc] || '').trim();
    var grade       = String(data[4][vc] || '').trim();
    var drawSize    = parseInt(data[5][vc], 10) || 0;
    var entries = [];
    for (var r = 8; r < data.length; r++) {
      var name = String(data[r][lc] || '').trim();
      if (!name || name === 'Entry Name') continue;
      entries.push({ name: name });
    }
    tournaments.push({ name: tournName, dateStr: dateStr, event: event,
      closingDate: closingDate, grade: grade, drawSize: drawSize, entries: entries });
  }
  return tournaments;
}

function buildPtsMap_U9_WL(rkSheet) {
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

function lookupPts_U9_WL(name, ptsMap) {
  if (!name) return '0';
  return ptsMap[String(name).trim().toLowerCase()] || '0';
}

function ptsAsNumber_U9_WL(name, ptsMap) {
  return parseFloat(lookupPts_U9_WL(name, ptsMap)) || 0;
}

function clearSheet_U9_WL(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 0) sheet.getRange(1, 1, lastRow, 7).clearContent();
}

function populateBlock_U9_WL(sheet, block, tourn, ptsMap) {
  if (!tourn) return;
  var titleRow  = block.titleRow;
  var firstData = block.firstData;
  var maxRows   = block.lastData - block.firstData + 1;
  var drawSize  = Math.min(tourn.drawSize, maxRows);
  var entries   = tourn.entries;
  sheet.getRange(titleRow,     2).setValue(tourn.drawSize);
  sheet.getRange(titleRow + 1, 1).setValue(tourn.name);
  sheet.getRange(titleRow + 2, 1).setValue(tourn.dateStr);
  sheet.getRange(titleRow + 3, 1).setValue('Closing Date:');
  sheet.getRange(titleRow + 3, 2).setValue(tourn.closingDate.split(' ')[0]);
  var colA = [], colB = [];
  for (var i = 0; i < maxRows; i++) {
    var name = entries[i] ? entries[i].name : '';
    colA.push([name]);
    colB.push([name ? lookupPts_U9_WL(name, ptsMap) : '']);
  }
  sheet.getRange(firstData, 1, maxRows, 1).setValues(colA);
  sheet.getRange(firstData, 2, maxRows, 1).setValues(colB);
  var colE = [];
  for (var i = 0; i < maxRows; i++) {
    colE.push([i < drawSize ? i + 1 : '']);
  }
  sheet.getRange(firstData, 5, maxRows, 1).setValues(colE);
  var withPts = [], zeroPts = [];
  for (var i = 0; i < entries.length; i++) {
    var name = entries[i].name;
    var pts  = ptsAsNumber_U9_WL(name, ptsMap);
    if (pts > 0) withPts.push({ name: name, pts: pts });
    else         zeroPts.push({ name: name });
  }
  withPts.sort(function(a, b) { return b.pts - a.pts; });
  var draw = withPts.concat(zeroPts).slice(0, drawSize);
  var colF = [], colG = [];
  for (var i = 0; i < draw.length; i++) {
    var name = draw[i] ? draw[i].name : '';
    colF.push([name]);
    colG.push([name ? lookupPts_U9_WL(name, ptsMap) : '']);
  }
  var actualF = colF.length;
  if (actualF > 0) {
    sheet.getRange(firstData, 6, actualF, 1).setValues(colF);
    sheet.getRange(firstData, 7, actualF, 1).setValues(colG);
  }
}