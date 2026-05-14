/**
 * populateSerge_WL
 * Reads SERGE_WATCHLIST + Open_rankings → writes SERGE_WL
 */

var BLOCKS_SERGE_WL = [
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

function populateSerge_WL() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wlSheet  = ss.getSheetByName('SERGE_WATCHLIST');
  var rkSheet  = ss.getSheetByName('Open_rankings');
  var outSheet = ss.getSheetByName('SERGE_WL');
  if (!wlSheet)  { Logger.log("Sheet 'SERGE_WATCHLIST' not found");  return; }
  if (!rkSheet)  { Logger.log("Sheet 'Open_rankings' not found");     return; }
  if (!outSheet) { Logger.log("Sheet 'SERGE_WL' not found");          return; }
  var wtnMap      = buildWtnMap_Serge_WL(rkSheet);
  var nrSheet     = ss.getSheetByName('Non_ranked_WTN');
  if (nrSheet) {
    var nrMap = buildNonRankedWtnMap_Serge_WL(nrSheet);
    for (var key in nrMap) {
      if (!wtnMap[key]) wtnMap[key] = nrMap[key];
    }
  }
  var tournaments = readTournaments_Serge_WL(wlSheet);
  Logger.log('Serge Watchlist tournaments: ' + tournaments.length);
  clearSheet_Serge_WL(outSheet);
  for (var i = 0; i < BLOCKS_SERGE_WL.length; i++) {
    populateBlock_Serge_WL(outSheet, BLOCKS_SERGE_WL[i], tournaments[i] || null, wtnMap);
  }
  Logger.log('Serge Watchlist: ' + tournaments.length + ' tournament(s) written.');
}

function readTournaments_Serge_WL(wlSheet) {
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
    if (!event.toUpperCase().startsWith('MS')) continue;
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

function buildWtnMap_Serge_WL(rkSheet) {
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

function buildNonRankedWtnMap_Serge_WL(nrSheet) {
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

function lookupWtn_Serge_WL(name, wtnMap) {
  if (!name) return 'NO WTN';
  return wtnMap[String(name).trim().toLowerCase()] || 'NO WTN';
}

function wtnAsNumber_Serge_WL(name, wtnMap) {
  var w = parseFloat(lookupWtn_Serge_WL(name, wtnMap));
  return isNaN(w) ? 99999 : w;
}

function clearSheet_Serge_WL(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 0) sheet.getRange(1, 1, lastRow, 7).clearContent();
}

function populateBlock_Serge_WL(sheet, block, tourn, wtnMap) {
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
    colB.push([name ? lookupWtn_Serge_WL(name, wtnMap) : '']);
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
    var w    = wtnAsNumber_Serge_WL(name, wtnMap);
    if (w < 99999) ranked.push({ name: name, wtn: w });
    else           unranked.push({ name: name });
  }
  ranked.sort(function(a, b) { return a.wtn - b.wtn; });
  var draw = ranked.concat(unranked).slice(0, drawSize);
  var colF = [], colG = [];
  for (var i = 0; i < draw.length; i++) {
    var name = draw[i] ? draw[i].name : '';
    colF.push([name]);
    colG.push([name ? lookupWtn_Serge_WL(name, wtnMap) : '']);
  }
  var actualF = colF.length;
  if (actualF > 0) {
    sheet.getRange(firstData, 6, actualF, 1).setValues(colF);
    sheet.getRange(firstData, 7, actualF, 1).setValues(colG);
  }
}