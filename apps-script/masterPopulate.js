/**
 * masterPopulate()
 *
 * Runs all populate scripts in the correct order:
 * 1. Player sheets first (they read from tournament sheets)
 * 2. Dashboard last (it reads from player sheets)
 * 3. Writes last-run timestamp to DASHBOARD!B1
 *
 * Set this as a time-based trigger, or call it from a webhook.
 */
function masterPopulate() {
  // When called from a webhook (doPost), getUi() is unavailable.
  // Patch it so downstream alert() calls log silently instead of crashing.
  try {
    SpreadsheetApp.getUi();
  } catch(e) {
    SpreadsheetApp.getUi = function() {
      return { alert: function(msg) { Logger.log(msg); } };
    };
  }

  populateLukaU14();
  populateLukaU16();
  populateSerge();
  populateDylanU9();
  populateDylanU10();

  populateLukaU14_WL();
  populateLukaU16_WL();
  populateDylanU9_WL();
  populateDylanU10_WL();
  populateSerge_WL();

  populateDashboard();
  populateWLDashboard();

  // Write last-run timestamp to DASHBOARD!B1 (overwrites previous value)
  var now = new Date();
  var dd  = String(now.getDate()).padStart(2, '0');
  var mm  = String(now.getMonth() + 1).padStart(2, '0');
  var yy  = String(now.getFullYear()).slice(-2);
  var hh  = String(now.getHours()).padStart(2, '0');
  var min = String(now.getMinutes()).padStart(2, '0');
  var ts  = dd + '/' + mm + '/' + yy + ' ' + hh + ':' + min;

  SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('DASHBOARD')
    .getRange('B1')
    .setValue(ts);
}

/**
 * To set up a time-based trigger:
 * 1. In Apps Script, click "Triggers" (alarm clock icon, left sidebar)
 * 2. Click "+ Add Trigger" (bottom right)
 * 3. Choose function: masterPopulate
 * 4. Event source: Time-driven
 * 5. Type: Hour timer or Day timer depending on how often data changes
 * 6. Save
 */