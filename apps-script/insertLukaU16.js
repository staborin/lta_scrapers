/**
 * populateLukaU16()
 *
 * Reads LUKA_TOURNAMENTS + U16_rankings, computes all values,
 * and writes them into LUKA_U16.
 *
 * Each of the 5 blocks has two independent sub-sections:
 *   Cols A-G  → non-Grade 5 tournaments (draw sorted by ranking)
 *   Cols I-R  → Grade 5 tournaments     (entries in source order for I/J,
 *                                         sorted by entry date for M/N/Q/R)
 *
 * Run this whenever tournament data changes.
 */

var BLOCKS_U16 = [
  { titleRow: 1,   firstData: 5,   lastData: 64  },
  { titleRow: 73,  firstData: 77,  lastData: 136 },
  { titleRow: 145, firstData: 149, lastData: 208 },
  { titleRow: 217, firstData: 221, lastData: 280 },
  { titleRow: 289, firstData: 293, lastData: 352 },
];

// ─── Main ─────────────────────────────────────────────────────────────────────
function populateLukaU16() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var ltSheet  = ss.getSheetByName('LUKA_TOURNAMENTS');
  var rkSheet  = ss.getSheetByName('U16_rankings');
  var u14Sheet = ss.getSheetByName('LUKA_U16');

  if (!ltSheet)  { Logger.log("Sheet 'LUKA_TOURNAMENTS' not found"); return; }
  if (!rkSheet)  { Logger.log("Sheet 'U16_rankings' not found"); return; }
  if (!u14Sheet) { Logger.log("Sheet 'LUKA_U16' not found"); return; }

  var rankMap     = buildRankMap_U16(rkSheet);
  var tournaments = readTournaments_U16(ltSheet);

  var grade5    = tournaments.filter(function(t) { return t.grade === 'Grade 5'; });
  var nonGrade5 = tournaments.filter(function(t) { return t.grade !== 'Grade 5'; });

  Logger.log('Grade 5: ' + grade5.length + ', Non-Grade 5: ' + nonGrade5.length);

  clearU14(u14Sheet);

  for (var i = 0; i < BLOCKS_U16.length; i++) {
    populateRightBlock_U16(u14Sheet, BLOCKS_U16[i], grade5[i]    || null, rankMap);
    populateLeftBlock_U16 (u14Sheet, BLOCKS_U16[i], nonGrade5[i] || null, rankMap);
  }

  Logger.log(
    'Luka 16U - Grade 5: ' + grade5.length + ', Non-Grade 5: ' + nonGrade5.length + ' tournament(s) written.'
  );
}

// ─── Read tournaments ─────────────────────────────────────────────────────────
// Column layout: 3-col stride — label col, value col, empty separator
//   Row 1: "Tournament: Name"  (label col)
//   Row 2: "Date:"   | date string  (value col)
//   Row 3: "Event:"  | type         (value col)
//   Row 5: "Grade:"  | grade        (value col)
//   Row 6: "Draw Size:" | number    (value col)
//   Row 9+: player name (label col) | entry datetime string (value col)
//
// IMPORTANT: Google Sheets may convert date strings to Date objects when reading.
// We use formatDate() to convert them back to "DD/MM/YYYY" strings.
function readTournaments_U16(ltSheet) {
  var lastCol = ltSheet.getLastColumn();
  var lastRow = ltSheet.getLastRow();
  if (lastCol < 2 || lastRow < 6) return [];

  // Read as display values to avoid date auto-conversion
  var data = ltSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var tournaments = [];

  for (var lc = 0; lc + 1 < data[0].length; lc += 3) {
    var vc = lc + 1;

    var eventType = String(data[2][vc] || '').trim();
    if (eventType !== '16U BS') continue;

    var rawName   = String(data[0][lc] || '').trim();
    var tournName = rawName.replace(/^Tournament:\s*/i, '');
    var dateStr   = String(data[1][vc] || '').trim();
    var grade     = String(data[4][vc] || '').trim();
    var drawSize  = parseInt(data[5][vc], 10) || 0;

    // Entries: preserve original source order (do NOT sort here)
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

// ─── Rankings ─────────────────────────────────────────────────────────────────
function buildRankMap_U16(rkSheet) {
  var data = rkSheet.getRange(2, 1, rkSheet.getLastRow() - 1, 2).getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][1] || '').trim();
    if (name) map[name.toLowerCase()] = String(data[i][0]);
  }
  return map;
}

function lookupRank_U16(name, rankMap) {
  if (!name) return 'NO RANK';
  return rankMap[String(name).trim().toLowerCase()] || 'NO RANK';
}

function rankAsNumber_U16(name, rankMap) {
  var r = parseInt(lookupRank_U16(name, rankMap), 10);
  return isNaN(r) ? 99999 : r;
}

// ─── Parse "Wed 21/01/2026 22:13" → timestamp for sorting ────────────────────
function parseEntryDate_U16(str) {
  if (!str) return 0;
  var m = str.match(/(\d{2})\/(\d{2})\/(\d{4})\s*(\d{1,2}):(\d{2})/);
  if (!m) return 0;
  return new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5]).getTime();
}

// ─── Clear all data areas ─────────────────────────────────────────────────────
function clearU14(sheet) {
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

// ─── Set a column as plain text then write values ─────────────────────────────
function writeAsText_U16(sheet, startRow, col, values) {
  var range = sheet.getRange(startRow, col, values.length, 1);
  range.setNumberFormat('@');       // force plain text — prevents date conversion
  range.setValues(values);
}

// ─── RIGHT block: Grade 5 — cols I-R ─────────────────────────────────────────
// Col I/J:  straight copy of source entries in original order
// Col L:    entry order number (1, 2, 3...)
// Col M/N:  entries sorted by entry date
// Col O:    ranking for each M entry
// Col Q:    top drawSize players in entry-date order (names)
// Col R:    their rankings
function populateRightBlock_U16(sheet, block, tourn, rankMap) {
  if (!tourn) return;

  var titleRow  = block.titleRow;
  var firstData = block.firstData;
  var maxRows   = block.lastData - block.firstData + 1;
  var drawSize  = Math.min(tourn.drawSize, maxRows);
  var entries   = tourn.entries; // original source order

  // Title rows (right block: cols I-K)
  // I1="Draw Size" is a static label already in sheet
  sheet.getRange(titleRow,     10).setValue(tourn.drawSize); // J1: draw size number
  sheet.getRange(titleRow,     11).setValue(tourn.grade);    // K1: grade
  sheet.getRange(titleRow + 1, 9).setValue(tourn.name);      // I2: tournament name
  sheet.getRange(titleRow + 2, 9).setValue(tourn.dateStr);   // I3: date

  // Cols I/J: straight copy of source order, dates as plain text
  var colI = [], colJ = [], colL = [];
  for (var i = 0; i < maxRows; i++) {
    var e = entries[i] || null;
    colI.push([e ? e.name    : '']);
    colJ.push([e ? e.dateStr : '']);
    colL.push([e ? i + 1    : '']);
  }
  sheet.getRange(firstData, 9,  maxRows, 1).setValues(colI); // I: name
  writeAsText_U16(sheet, firstData, 10, colJ);                   // J: entry date (plain text)
  sheet.getRange(firstData, 12, maxRows, 1).setValues(colL); // L: entry order

  // Cols M/N: sorted by entry date ascending
  var sorted = entries.slice().sort(function(a, b) {
    return parseEntryDate_U16(a.dateStr) - parseEntryDate_U16(b.dateStr);
  });

  var colM = [], colN = [], colO = [];
  for (var i = 0; i < maxRows; i++) {
    var e    = sorted[i] || null;
    var name = e ? e.name    : '';
    var date = e ? e.dateStr : '';
    colM.push([name]);
    colN.push([date]);
    colO.push([name ? lookupRank_U16(name, rankMap) : '']);
  }
  sheet.getRange(firstData, 13, maxRows, 1).setValues(colM); // M
  writeAsText_U16(sheet, firstData, 14, colN);                   // N: date (plain text)
  sheet.getRange(firstData, 15, maxRows, 1).setValues(colO); // O

  // Cols Q/R: take first drawSize from entry-date sorted list (col M),
  // then sort THOSE by ranking ascending
  var topEntries = sorted.slice(0, drawSize);
  topEntries.sort(function(a, b) {
    return rankAsNumber_U16(a.name, rankMap) - rankAsNumber_U16(b.name, rankMap);
  });

  var colQ = [], colR = [];
  for (var i = 0; i < topEntries.length; i++) {
    var name = topEntries[i].name;
    colQ.push([name]);
    colR.push([lookupRank_U16(name, rankMap)]);
  }
  var actualQ = colQ.length;
  if (actualQ > 0) {
    sheet.getRange(firstData, 17, actualQ, 1).setValues(colQ);
    sheet.getRange(firstData, 18, actualQ, 1).setValues(colR);
  }
}

// ─── LEFT block: non-Grade 5 — cols A-G ──────────────────────────────────────
// Col A/B:  entries in source order + their rankings
// Col E:    ranking order numbers 1..drawSize
// Col F/G:  top drawSize players sorted by ranking (best first)
function populateLeftBlock_U16(sheet, block, tourn, rankMap) {
  if (!tourn) return;

  var titleRow  = block.titleRow;
  var firstData = block.firstData;
  var maxRows   = block.lastData - block.firstData + 1;
  var drawSize  = Math.min(tourn.drawSize, maxRows);
  var entries   = tourn.entries;

  // Title rows (left block: cols A-B)
  // A1="Draw Size" is a static label already in sheet
  sheet.getRange(titleRow,     2).setValue(tourn.drawSize); // B1: draw size number
  sheet.getRange(titleRow + 1, 1).setValue(tourn.name);     // A2: tournament name
  sheet.getRange(titleRow + 2, 1).setValue(tourn.dateStr);  // A3: date

  // Cols A/B: source order
  var colA = [], colB = [];
  for (var i = 0; i < maxRows; i++) {
    var name = entries[i] ? entries[i].name : '';
    colA.push([name]);
    colB.push([name ? lookupRank_U16(name, rankMap) : '']);
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
    var r    = rankAsNumber_U16(name, rankMap);
    if (r < 99999) {
      ranked.push({ name: name, rank: r });
    } else {
      unranked.push({ name: name });
    }
  }
  ranked.sort(function(a, b) { return a.rank - b.rank; });
  var draw = ranked.concat(unranked).slice(0, drawSize);

  var colF = [], colG = [];
  for (var i = 0; i < drawSize; i++) {
    var name = draw[i] ? draw[i].name : '';
    colF.push([name]);
    colG.push([name ? lookupRank_U16(name, rankMap) : '']);
  }
  var actualF = colF.length;
  if (actualF > 0) {
    sheet.getRange(firstData, 6, actualF, 1).setValues(colF);
    sheet.getRange(firstData, 7, actualF, 1).setValues(colG);
  }
}