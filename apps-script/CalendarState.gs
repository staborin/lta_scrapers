/**
 * CalendarState.gs
 *
 * Manages the CALENDAR_STATE sheet which tracks tournaments
 * added to a user's calendar via the web app.
 *
 * Sheet structure:
 *   Column A: Identifier (tournament URL)
 *   Column B: Added At (ISO timestamp)
 *
 * Called from Code.js doPost when body.action === 'markCalendarAdded'.
 */

function markCalendarAdded(identifier) {
  if (!identifier) {
    return { status: 'error', message: "Missing 'identifier'" };
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CALENDAR_STATE');
  if (!sheet) {
    return { status: 'error', message: "Sheet 'CALENDAR_STATE' not found" };
  }
  sheet.appendRow([identifier, new Date().toISOString()]);
  return { status: 'ok', message: 'Calendar state updated' };
}
