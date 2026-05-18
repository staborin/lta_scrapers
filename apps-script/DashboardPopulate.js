/**
 * populateDashboard()
 *
 * Writes only VALUES into the dashboard (labels are static in the sheet).
 *
 * GS column layout (confirmed from screenshot):
 *   LUKA:  values in col C (3)
 *   DYLAN: values in col J (10)
 *   SERGE: values in col O (15)
 *
 * Slot start rows: 5, 15, 25, 35, 45
 * Each slot is 8 rows: Tournament, Date, Event, Status, Grade, Draw Size, Ranking, URL
 */

var DASH_SLOT_ROWS_DB = [5, 15, 25, 35, 45];

var LUKA_BLOCKS_F_DB  = ['F5:F64',   'F77:F136',  'F149:F208', 'F221:F280', 'F293:F352'];
var LUKA_BLOCKS_Q_DB  = ['Q5:Q64',   'Q77:Q136',  'Q149:Q208', 'Q221:Q280', 'Q293:Q352'];
var DU9_BLOCKS_F_DB   = ['F5:F64',   'F77:F136',  'F149:F208', 'F221:F280', 'F293:F352'];
var DU10_BLOCKS_F_DB  = ['F5:F64',   'F77:F136',  'F149:F208', 'F221:F280', 'F293:F352'];
var DU10_BLOCKS_Q_DB  = ['Q5:Q64',   'Q77:Q136',  'Q149:Q208', 'Q221:Q280', 'Q293:Q352'];
var SERGE_BLOCKS_F_DB = ['F5:F64',   'F77:F136',  'F149:F208', 'F221:F280', 'F293:F352'];
var SERGE_BLOCKS_Q_DB = ['Q5:Q64',   'Q77:Q136',  'Q149:Q208', 'Q221:Q280', 'Q293:Q352'];

// Rows where tournament names are written in each output sheet (titleRow + 1)
var LUKA_NAME_ROWS_DB   = [2, 74, 146, 218, 290];   // LUKA_U14, LUKA_U16
var DU9_NAME_ROWS_DB    = [2, 74, 146, 218, 290];   // DYLAN_U9
var DU10_NAME_ROWS_DB   = [2, 74, 146, 218, 290];   // DYLAN_U10
var SERGE_NAME_ROWS_DB  = [2, 74, 146, 218, 290];   // SERGE

function populateDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var dbSheet = ss.getSheetByName('DASHBOARD');
  var ltSheet = ss.getSheetByName('LUKA_TOURNAMENTS');
  var dtSheet = ss.getSheetByName('DYLAN_TOURNAMENTS');
  var stSheet = ss.getSheetByName('SERGE_TOURNAMENTS');
  var lu14    = ss.getSheetByName('LUKA_U14');
  var lu16    = ss.getSheetByName('LUKA_U16');
  var du9     = ss.getSheetByName('DYLAN_U9');
  var du10    = ss.getSheetByName('DYLAN_U10');
  var serge   = ss.getSheetByName('SERGE');

  if (!dbSheet) { Logger.log("Sheet 'DASHBOARD' not found"); return; }

  var today = new Date(); today.setHours(0,0,0,0);

  dbSheet.getRange('B5:D53').clearContent();
  dbSheet.getRange('H5:J53').clearContent();
  dbSheet.getRange('N5:P53').clearContent();

  var lukaTourns  = getUpcomingDB(ltSheet,  today);
  var dylanTourns = getUpcomingDB(dtSheet,  today);
  var sergeTourns = getUpcomingDB(stSheet,  today);

  writeSlotsDB(dbSheet, lukaTourns,  3,  lu14, lu16,  null,  null,  'Luka Taborin',  'luka');
  writeSlotsDB(dbSheet, dylanTourns, 9,  null, null,  du9,   du10,  'Dylan Taborin', 'dylan');
  writeSlotsDB(dbSheet, sergeTourns, 15, null, null,  null,  serge, 'Serge Taborin', 'serge');

  Logger.log('Dashboard updated.');
}

function getUpcomingDB(tSheet, today) {
  var lastCol = tSheet.getLastColumn();
  var lastRow = tSheet.getLastRow();
  if (lastCol < 2 || lastRow < 6) return [];

  var data = tSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var result = [];

  for (var lc = 0; lc + 1 < data[0].length; lc += 3) {
    var vc      = lc + 1;
    var rawName = String(data[0][lc] || '').trim();
    if (!rawName) continue;

    var name     = rawName.replace(/^Tournament:\s*/i, '');
    var dateStr  = String(data[1][vc] || '').trim();
    var startTime = String(data[1][lc + 2] || '').trim();
    var event    = String(data[2][vc] || '').trim();
    var status      = String(data[3][vc] || '').trim();
    var closingDate = (lc + 2 < data[3].length) ? String(data[3][lc + 2] || '').trim() : '';
    var grade    = String(data[4][vc] || '').trim();
    var drawSize = String(data[5][vc] || '').trim();
    var url      = String(data[6][vc] || '').trim();
    var endDate  = parseEndDateDB(dateStr);

    if (endDate < today) continue;

    result.push({ name: name, dateStr: dateStr, startTime: startTime, event: event,
      status: status, closingDate: closingDate, grade: grade, drawSize: drawSize, url: url, endDate: endDate });
  }

  result.sort(function(a, b) { return a.endDate - b.endDate; });
  return result;
}

function parseEndDateDB(str) {
  if (!str) return new Date(0);
  var dates = str.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/g);
  if (!dates) return new Date(0);
  var last = dates[dates.length - 1].match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return new Date(+last[3], +last[2] - 1, +last[1]);
}

function findPosDB(sheet, rangeStr, playerName) {
  if (!sheet || !rangeStr) return 'N/A';
  try {
    var vals = sheet.getRange(rangeStr).getValues();
    for (var r = 0; r < vals.length; r++) {
      if (String(vals[r][0]).trim().toLowerCase() === playerName.toLowerCase()) return r + 1;
    }
    return 'N/A';
  } catch(e) { return 'N/A'; }
}

// ─── Build tournament name|date → block index map from an output sheet ───────
function buildBlockNameMapDB(sheet, nameRows, col) {
  if (!sheet) return {};
  var map = {};
  for (var i = 0; i < nameRows.length; i++) {
    var name = String(sheet.getRange(nameRows[i], col).getValue() || '').trim();
    var date = String(sheet.getRange(nameRows[i] + 1, col).getDisplayValue() || '').trim();
    if (name) map[(name + '|' + date).toLowerCase()] = i;
  }
  return map;
}

function writeSlotsDB(dbSheet, tourns, valueCol, sheet14or9, sheet16or10, sheetU9, sheetU10orSerge, playerName, playerKey) {
  // Build name → block index maps from the output sheets (left=col 1, right=col 9)
  var maps = {};

  if (playerKey === 'luka') {
    maps.u14Left  = buildBlockNameMapDB(sheet14or9,  LUKA_NAME_ROWS_DB, 1);
    maps.u14Right = buildBlockNameMapDB(sheet14or9,  LUKA_NAME_ROWS_DB, 9);
    maps.u16Left  = buildBlockNameMapDB(sheet16or10, LUKA_NAME_ROWS_DB, 1);
    maps.u16Right = buildBlockNameMapDB(sheet16or10, LUKA_NAME_ROWS_DB, 9);
  } else if (playerKey === 'dylan') {
    maps.u9Left   = buildBlockNameMapDB(sheetU9,           DU9_NAME_ROWS_DB,  1);
    maps.u10Left  = buildBlockNameMapDB(sheetU10orSerge,   DU10_NAME_ROWS_DB, 1);
    maps.u10Right = buildBlockNameMapDB(sheetU10orSerge,   DU10_NAME_ROWS_DB, 9);
  } else {
    maps.sergeLeft  = buildBlockNameMapDB(sheetU10orSerge, SERGE_NAME_ROWS_DB, 1);
    maps.sergeRight = buildBlockNameMapDB(sheetU10orSerge, SERGE_NAME_ROWS_DB, 9);
  }

  for (var i = 0; i < Math.min(tourns.length, 5); i++) {
    var t       = tourns[i];
    var slotRow = DASH_SLOT_ROWS_DB[i];
    var tName   = (t.name + '|' + t.dateStr).toLowerCase();

    var pos = 'N/A';
    if (playerKey === 'luka') {
      var sheet    = t.event === '14U BS' ? sheet14or9 : sheet16or10;
      var nameMap  = t.event === '14U BS'
        ? (t.grade === 'Grade 5' ? maps.u14Right : maps.u14Left)
        : (t.grade === 'Grade 5' ? maps.u16Right : maps.u16Left);
      var blocksF  = t.grade === 'Grade 5' ? LUKA_BLOCKS_Q_DB : LUKA_BLOCKS_F_DB;
      var blockIdx = nameMap[tName];
      if (blockIdx !== undefined && blockIdx < blocksF.length) {
        pos = findPosDB(sheet, blocksF[blockIdx], playerName);
      }

    } else if (playerKey === 'dylan') {
      if (t.event === '9U BS') {
        var blockIdx = maps.u9Left[tName];
        if (blockIdx !== undefined && blockIdx < DU9_BLOCKS_F_DB.length) {
          pos = findPosDB(sheetU9, DU9_BLOCKS_F_DB[blockIdx], playerName);
        }
      } else {
        var nameMap  = t.grade === 'Grade 5' ? maps.u10Right : maps.u10Left;
        var blocksF  = t.grade === 'Grade 5' ? DU10_BLOCKS_Q_DB : DU10_BLOCKS_F_DB;
        var blockIdx = nameMap[tName];
        if (blockIdx !== undefined && blockIdx < blocksF.length) {
          pos = findPosDB(sheetU10orSerge, blocksF[blockIdx], playerName);
        }
      }

    } else {
      var nameMap  = t.grade === 'Grade 5' ? maps.sergeRight : maps.sergeLeft;
      var blocksF  = t.grade === 'Grade 5' ? SERGE_BLOCKS_Q_DB : SERGE_BLOCKS_F_DB;
      var blockIdx = nameMap[tName];
      if (blockIdx !== undefined && blockIdx < blocksF.length) {
        pos = findPosDB(sheetU10orSerge, blocksF[blockIdx], playerName);
      }
      if (pos === 'N/A' && t.grade !== 'Grade 5' && blockIdx !== undefined) {
        pos = findPosDB(sheetU10orSerge, blocksF[blockIdx], 'Serge TABORIN');
      }
    }

    var labelCol = valueCol - 1;
    var labels = ['Tournament:', 'Date:', 'Event:', 'Status:', 'Grade:', 'Draw Size:',
                  playerKey === 'luka' ? 'Luka ranking:' : playerKey === 'dylan' ? 'Dylan ranking:' : 'Serge ranking:',
                  'URL:'];
    var values = [t.name, t.dateStr, t.event, t.status, t.grade, t.drawSize, pos, t.url || ''];
    for (var r = 0; r < 8; r++) {
      dbSheet.getRange(slotRow + r, labelCol).setValue(labels[r]);
      dbSheet.getRange(slotRow + r, valueCol).setValue(values[r]);
    }
    if (t.startTime) {
      dbSheet.getRange(slotRow + 1, valueCol + 1).setValue(t.startTime);
    }
    // Write closing date to the column after the status value
    if (t.closingDate) {
      dbSheet.getRange(slotRow + 3, valueCol + 1).setValue(t.closingDate);
    }
  }
}