/**
 * populateWLDashboard()
 *
 * Writes watchlist tournament data into WATCHLIST_DASHBOARD.
 * Shows hypothetical ranking position for each player.
 *
 * Layout mirrors main DASHBOARD:
 *   LUKA:  labels col B, values col C
 *   DYLAN: labels col I, values col J
 *   SERGE: labels col N, values col O
 *
 * 10 slots per player, starting row 5, spacing 10 rows.
 * Each slot: Tournament, Date, Event, Closing Date, Grade, Draw Size, Ranking, URL
 *
 * Filters out tournaments whose closing date has passed.
 */

var WL_SLOT_ROWS = [5, 15, 25, 35, 45, 55, 65, 75, 85, 95];

var LUKA_WL_BLOCKS_F  = ['F6:F65',   'F79:F138',  'F152:F211', 'F225:F284', 'F298:F357',
                          'F371:F430','F444:F503', 'F517:F576', 'F590:F649', 'F663:F722'];
var DYLAN_WL_BLOCKS_F = ['F6:F65',   'F79:F138',  'F152:F211', 'F225:F284', 'F298:F357',
                          'F371:F430','F444:F503', 'F517:F576', 'F590:F649', 'F663:F722'];
var SERGE_WL_BLOCKS_F = ['F6:F65',   'F79:F138',  'F152:F211', 'F225:F284', 'F298:F357',
                          'F371:F430','F444:F503', 'F517:F576', 'F590:F649', 'F663:F722'];

// Row in each output sheet block where the tournament name is written (titleRow + 1)
var WL_BLOCK_NAME_ROWS = [2, 75, 148, 221, 294, 367, 440, 513, 586, 659];

// ─── Main ─────────────────────────────────────────────────────────────────────
function populateWLDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var dbSheet  = ss.getSheetByName('WATCHLIST_DASHBOARD');
  var luWL     = ss.getSheetByName('LUKA_WATCHLIST');
  var dyWL     = ss.getSheetByName('DYLAN_WATCHLIST');
  var seWL     = ss.getSheetByName('SERGE_WATCHLIST');
  var luU14WL  = ss.getSheetByName('LUKA_U14_WL');
  var luU16WL  = ss.getSheetByName('LUKA_U16_WL');
  var dyU9WL   = ss.getSheetByName('DYLAN_U9_WL');
  var dyU10WL  = ss.getSheetByName('DYLAN_U10_WL');
  var seWLOut  = ss.getSheetByName('SERGE_WL');

  if (!dbSheet) { Logger.log("Sheet 'WATCHLIST_DASHBOARD' not found"); return; }

  var today = new Date(); today.setHours(0,0,0,0);

  // Clear value columns only
  dbSheet.getRange('B5:D105').clearContent();
  dbSheet.getRange('H5:K105').clearContent();
  dbSheet.getRange('N5:P105').clearContent();

  var lukaTourns  = getUpcomingWL(luWL,  today);
  var dylanTourns = getUpcomingWL(dyWL,  today);
  var sergeTourns = getUpcomingWL(seWL,  today);

  var lukaUrls  = getWatchlistUrls(ss, 'luka');
  var dylanUrls = getWatchlistUrls(ss, 'dylan');
  var sergeUrls = getWatchlistUrls(ss, 'serge');

  writeSlotsWL(dbSheet, lukaTourns,  3,  luU14WL,  luU16WL, 'Luka Taborin',  'luka',  lukaUrls);
  writeSlotsWL(dbSheet, dylanTourns, 10, dyU9WL,   dyU10WL, 'Dylan Taborin', 'dylan', dylanUrls);
  writeSlotsWL(dbSheet, sergeTourns, 15, seWLOut,  null,    'Serge Taborin', 'serge', sergeUrls);

  Logger.log('Watchlist Dashboard updated.');
}

// ─── Read URLs from WATCHLIST sheet by player ─────────────────────────────────
function getWatchlistUrls(ss, playerKey) {
  var sheet = ss.getSheetByName('WATCHLIST');
  if (!sheet) return [];

  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  if (lastRow < 6) return [];

  var data = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  // Col indices (0-based): luka=2, dylan=6, serge=10
  var urlCol = playerKey === 'luka' ? 2 : playerKey === 'dylan' ? 6 : 10;

  var urls = [];
  for (var r = 5; r < data.length; r++) {  // data from row 6 (index 5)
    var url = String(data[r][urlCol] || '').trim();
    if (url) urls.push(url);
  }
  return urls;
}

// ─── Read watchlist tournaments, filter out past closing dates ────────────────
function getUpcomingWL(wlSheet, today) {
  if (!wlSheet) return [];
  var lastCol = wlSheet.getLastColumn();
  var lastRow = wlSheet.getLastRow();
  if (lastCol < 2 || lastRow < 6) return [];

  var data = wlSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var result = [];

  for (var lc = 0; lc + 1 < data[0].length; lc += 3) {
    var vc      = lc + 1;
    var rawName = String(data[0][lc] || '').trim();
    if (!rawName || !rawName.match(/^Tournament:/i)) continue;

    var name        = rawName.replace(/^Tournament:\s*/i, '');
    var dateStr     = String(data[1][vc] || '').trim();
    var startTime   = String(data[1][lc + 2] || '').trim();
    var event       = String(data[2][vc] || '').trim();
    var closingDate = String(data[3][vc] || '').trim();
    var grade       = String(data[4][vc] || '').trim();
    var drawSize    = String(data[5][vc] || '').trim();

    // Filter out if closing date has passed
    var closing = parseClosingWL(closingDate);
    if (closing && closing < today) continue;

    result.push({
      name: name, dateStr: dateStr, startTime: startTime, event: event,
      closingDate: closingDate, grade: grade, drawSize: drawSize,
      closing: closing
    });
  }

  return result;
}

function parseClosingWL(str) {
  if (!str) return null;
  var m = str.match(/(\d{2})\/(\d{2})\/(\d{4})/);
  if (!m) return null;
  return new Date(+m[3], +m[2] - 1, +m[1]);
}

// ─── Build tournament name → block index map from output sheet ───────────────
function buildBlockNameMap(sheet) {
  if (!sheet) return {};
  var map = {};
  for (var i = 0; i < WL_BLOCK_NAME_ROWS.length; i++) {
    var name = String(sheet.getRange(WL_BLOCK_NAME_ROWS[i], 1).getValue() || '').trim();
    if (name) map[name.toLowerCase()] = i;
  }
  return map;
}

// ─── Find player position in a draw column ────────────────────────────────────
function findPosWL(sheet, rangeStr, playerName) {
  if (!sheet || !rangeStr) return 'N/A';
  try {
    var vals = sheet.getRange(rangeStr).getValues();
    for (var r = 0; r < vals.length; r++) {
      if (String(vals[r][0]).trim().toLowerCase() === playerName.toLowerCase()) return r + 1;
    }
    return 'N/A';
  } catch(e) { return 'N/A'; }
}

// ─── Write slots for one player ───────────────────────────────────────────────
function writeSlotsWL(dbSheet, tourns, valueCol, outSheet1, outSheet2, playerName, playerKey, urls) {
  // Build name → block index maps from the output sheets
  var map1 = buildBlockNameMap(outSheet1);
  var map2 = buildBlockNameMap(outSheet2);

  for (var i = 0; i < Math.min(tourns.length, 10); i++) {
    var t       = tourns[i];
    var slotRow = WL_SLOT_ROWS[i];

    // Determine correct output sheet and find block by tournament name
    var pos = 'N/A';
    var targetSheet = null;
    var targetMap   = null;
    var blocksF     = null;

    if (playerKey === 'luka') {
      targetSheet = t.event === '16U BS' ? outSheet2 : outSheet1;
      targetMap   = t.event === '16U BS' ? map2 : map1;
      blocksF     = LUKA_WL_BLOCKS_F;
    } else if (playerKey === 'dylan') {
      targetSheet = t.event === '9U BS' ? outSheet1 : outSheet2;
      targetMap   = t.event === '9U BS' ? map1 : map2;
      blocksF     = DYLAN_WL_BLOCKS_F;
    } else {
      targetSheet = outSheet1;
      targetMap   = map1;
      blocksF     = SERGE_WL_BLOCKS_F;
    }

    var blockIdx = targetMap[t.name.toLowerCase()];
    if (blockIdx !== undefined && blockIdx < blocksF.length) {
      pos = findPosWL(targetSheet, blocksF[blockIdx], playerName);
    }

    // Write labels and values
    var labelCol  = valueCol - 1;
    var rankLabel = playerKey === 'luka'  ? 'Luka ranking:'  :
                    playerKey === 'dylan' ? 'Dylan ranking:' : 'Serge ranking:';

    var labels = ['Tournament:', 'Date:', 'Event:', 'Closing Date:', 'Grade:', 'Draw Size:', rankLabel, 'URL:'];
    var values = [t.name, t.dateStr, t.event, t.closingDate, t.grade, t.drawSize, pos, urls[i] || ''];

    for (var r = 0; r < 8; r++) {
      dbSheet.getRange(slotRow + r, labelCol).setValue(labels[r]);
      dbSheet.getRange(slotRow + r, valueCol).setValue(values[r]);
    }
    // Write start time to the column after the date value
    if (t.startTime) {
      dbSheet.getRange(slotRow + 1, valueCol + 1).setValue(t.startTime);
    }
  }
}