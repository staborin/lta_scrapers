/**
 * LTA Tournament Scraper — Google Apps Script Web App
 *
 * Acts as a data layer only — reads and writes sheet data.
 * All orchestration (refreshAllEntries) is handled by the Cloud Function.
 *
 * doPost handles:
 *   - Writing a grid of data to a sheet tab
 *   - Triggering masterPopulate
 *
 * doGet handles:
 *   - Reading a sheet tab and returning its data as a 2D array
 */

const SECRET = ""; // Set to a string to require a shared secret, e.g. "abc123"
                   // Leave blank to disable secret checking.

function doPost(e) {
  try {
    // --- Auth check ---
    if (SECRET) {
      const incoming = (e.parameter && e.parameter.secret) || "";
      if (incoming !== SECRET) {
        return jsonResponse({ status: "error", message: "Unauthorised" }, 403);
      }
    }

    // --- Parse body ---
    const body = JSON.parse(e.postData.contents);

    // --- masterPopulate trigger ---
    if (body.action === 'masterPopulate') {
      masterPopulate();
      return jsonResponse({ status: 'ok', message: 'masterPopulate triggered' });
    }

    const sheetName = body.sheet;
    const clearFirst = body.clearFirst !== false; // default true
    const startRow = body.startRow || 1;          // default row 1
    const startCol = body.startCol || 1;          // default col 1
    const rows = body.rows;

    if (!sheetName) return jsonResponse({ status: "error", message: "Missing 'sheet'" });
    if (!Array.isArray(rows) || rows.length === 0) {
      return jsonResponse({ status: "error", message: "Missing or empty 'rows'" });
    }

    // --- Get or create sheet ---
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    // --- Clear if requested ---
    if (clearFirst) {
      sheet.clearContents();
    }

    // --- Write the grid ---
    const numRows = rows.length;
    const numCols = rows[0].length;

    const normalised = rows.map(row => {
      const padded = row.slice();
      while (padded.length < numCols) padded.push("");
      return padded;
    });

    sheet.getRange(startRow, startCol, numRows, numCols).setValues(normalised);

    // --- Bold row 9 (entry headers) — only applies when writing from row 1 ---
    if (startRow === 1 && numRows >= 9) {
      sheet.getRange(9, 1, 1, numCols).setFontWeight("bold");
    }

    // --- Bold row 1 (tournament name headers) ---
    if (startRow === 1) {
      sheet.getRange(1, 1, 1, numCols).setFontWeight("bold");
    }

    return jsonResponse({
      status: "ok",
      rowsWritten: numRows,
      colsWritten: numCols,
      sheet: sheetName,
      startRow: startRow,
      startCol: startCol,
    });

  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function doGet(e) {
  try {
    const action = e.parameter && e.parameter.action;

    // ── Read a sheet and return its data as a 2D array ──
    if (action === "read") {
      if (SECRET) {
        const incoming = (e.parameter && e.parameter.secret) || "";
        if (incoming !== SECRET) {
          return jsonResponse({ status: "error", message: "Unauthorised" });
        }
      }
      const sheetName = e.parameter.sheet;
      if (!sheetName) return jsonResponse({ status: "error", message: "Missing 'sheet'" });

      const ss    = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return jsonResponse({ status: "error", message: `Sheet '${sheetName}' not found` });

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow === 0 || lastCol === 0) {
        return jsonResponse({ status: "ok", rows: [] });
      }

      const rows = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      return jsonResponse({ status: "ok", rows: rows });
    }

    // ── Default: health check ──
    return jsonResponse({ status: "ok", message: "LTA Sheets webapp is running." });

  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function jsonResponse(obj, code) {
  const output = ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}


function scheduledRefresh() {
  UrlFetchApp.fetch(
    "https://europe-west2-lta-tournament-tracker.cloudfunctions.net/refresh-entries",
    {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ action: "refreshAllEntries" }),
      muteHttpExceptions: true
    }
  );
}
