/**
 * populateLukaU14_WL()
 *
 * Reads LUKA_WATCHLIST + U14_rankings, computes hypothetical ranking position,
 * and writes into LUKA_U14_WL.
 *
 * All watchlist tournaments are non-Grade 5 — left block only (cols A-G).
 * 10 blocks to accommodate up to 10 watchlist tournaments.
 *
 * Col A/B:  entries in source order + their rankings
 * Col E:    ranking order numbers 1..drawSize
 * Col F/G:  top drawSize players sorted by ranking (best first)
 *           Luka Taborin will appear at his hypothetical position.
 *
 * Row layout per block:
 *   titleRow+0: col B = draw size
 *   titleRow+1: col A = tournament name
 *   titleRow+2: col A = tournament date
 *   titleRow+3: col A = closing date
 */

var BLOCKS_U14_WL = [
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

// ─── Main ─────────────────────────────────────────────────────────────────────
function populateLukaU14_WL() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var wlSheet  = ss.getSheetByName('LUKA_WATCHLIST');
  var rkSheet  = ss.getSheetByName('U14_rankings');
  var outSheet = ss.getSheetByName('LUKA_U14_WL');

  if (!wlSheet)  { Logger.log("Sheet 'LUKA_WATCHLIST' not found");  return; }
  if (!rkSheet)  { Logger.log("Sheet 'U14_rankings' not found");     return; }
  if (!outSheet) { Logger.log("Sheet 'LUKA_U14_WL' not found");      return; }

  var rankMap     = buildRankMap_U14_WL(rkSheet);
  var tournaments = readTournaments_U14_WL(wlSheet);

  Logger.log('Watchlist tournaments: ' + tournaments.length);

  clearU14_WL(outSheet);

  for (var i = 0; i < BLOCKS_U14_WL.length; i++) {
    populateBlock_U14_WL(outSheet, BLOCKS_U14_WL[i], tournaments[i] || null, rankMap);
  }

  Logger.log('Luka U14 Watchlist: ' + tournaments.length + ' tournament(s) written.');
}

// ─── Read tournaments from LUKA_WATCHLIST ─────────────────────────────────────
// Same 3-col stride as LUKA_TOURNAMENTS but row 4 is "Closing Date:" not "Status:"
function readTournaments_U14_WL(wlSheet) {
  var lastCol = wlSheet.getLastColumn();
  var lastRow = wlSheet.getLastRow();
  if (lastCol < 2 || lastRow < 6) return [];

  var data = wlSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var tournaments = [];

  for (var lc = 0; lc + 1 < data[0].length; lc += 3) {
    var vc = lc + 1;

    var rawName = String(data[0][lc] || '').trim();
    if (!rawName || !rawName.match(/^Tournament:/i)) continue;

    var tournName    = rawName.replace(/^Tournament:\s*/i, '');
    var dateStr      = String(data[1][vc] || '').trim();
    var event        = String(data[2][vc] || '').trim();
    var closingDate  = String(data[3][vc] || '').trim();
    var grade        = String(data[4][vc] || '').trim();
    var drawSize     = parseInt(data[5][vc], 10) || 0;

    var entries = [];
    for (var r = 8; r < data.length; r++) {
      var name = String(data[r][lc] || '').trim();
      if (!name || name === 'Entry Name') continue;
      entries.push({ name: name });
    }

    tournaments.push({
      name:        tournName,
      dateStr:     dateStr,
      event:       event,
      closingDate: closingDate,
      grade:       grade,
      drawSize:    drawSize,
      entries:     entries,
    });
  }

  return tournaments;
}

// ─── Rankings ─────────────────────────────────────────────────────────────────
function buildRankMap_U14_WL(rkSheet) {
  var data = rkSheet.getRange(2, 1, rkSheet.getLastRow() - 1, 2).getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][1] || '').trim();
    if (name) map[name.toLowerCase()] = String(data[i][0]);
  }
  return map;
}

function lookupRank_U14_WL(name, rankMap) {
  if (!name) return 'NO RANK';
  return rankMap[String(name).trim().toLowerCase()] || 'NO RANK';
}

function rankAsNumber_U14_WL(name, rankMap) {
  var r = parseInt(lookupRank_U14_WL(name, rankMap), 10);
  return isNaN(r) ? 99999 : r;
}

// ─── Clear ────────────────────────────────────────────────────────────────────
function clearU14_WL(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    sheet.getRange(1, 1, lastRow, 7).clearContent();
  }
}

// ─── Populate block: cols A-G ─────────────────────────────────────────────────
function populateBlock_U14_WL(sheet, block, tourn, rankMap) {
  if (!tourn) return;

  var titleRow  = block.titleRow;
  var firstData = block.firstData;
  var maxRows   = block.lastData - block.firstData + 1;
  var drawSize  = Math.min(tourn.drawSize, maxRows);
  var entries   = tourn.entries;

  // Title rows
  sheet.getRange(titleRow,     2).setValue(tourn.drawSize);          // B1: draw size
  sheet.getRange(titleRow + 1, 1).setValue(tourn.name);              // A2: name
  sheet.getRange(titleRow + 2, 1).setValue(tourn.dateStr);           // A3: date
  sheet.getRange(titleRow + 3, 1).setValue('Closing Date:');         // A4: label
  sheet.getRange(titleRow + 3, 2).setValue(tourn.closingDate.split(' ')[0]); // B4: date only (strips time)

  // Cols A/B: source order + rankings
  var colA = [], colB = [];
  for (var i = 0; i < maxRows; i++) {
    var name = entries[i] ? entries[i].name : '';
    colA.push([name]);
    colB.push([name ? lookupRank_U14_WL(name, rankMap) : '']);
  }
  sheet.getRange(firstData, 1, maxRows, 1).setValues(colA);
  sheet.getRange(firstData, 2, maxRows, 1).setValues(colB);

  // Col E: 1..drawSize
  var colE = [];
  for (var i = 0; i < maxRows; i++) {
    colE.push([i < drawSize ? i + 1 : '']);
  }
  sheet.getRange(firstData, 5, maxRows, 1).setValues(colE);

  // Cols F/G: sorted by ranking, capped at drawSize
  var ranked = [], unranked = [];
  for (var i = 0; i < entries.length; i++) {
    var name = entries[i].name;
    var r    = rankAsNumber_U14_WL(name, rankMap);
    if (r < 99999) {
      ranked.push({ name: name, rank: r });
    } else {
      unranked.push({ name: name });
    }
  }
  ranked.sort(function(a, b) { return a.rank - b.rank; });
  var draw = ranked.concat(unranked).slice(0, drawSize);

  var colF = [], colG = [];
  for (var i = 0; i < draw.length; i++) {
    var name = draw[i] ? draw[i].name : '';
    colF.push([name]);
    colG.push([name ? lookupRank_U14_WL(name, rankMap) : '']);
  }
  var actualF = colF.length;
  if (actualF > 0) {
    sheet.getRange(firstData, 6, actualF, 1).setValues(colF);
    sheet.getRange(firstData, 7, actualF, 1).setValues(colG);
  }
}