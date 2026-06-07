/**
 * Google Apps Script — Data collection endpoint for BART & BRET demos
 * =====================================================================
 * SETUP (takes ~5 minutes):
 *
 *  1. Go to https://script.google.com and click "New project"
 *  2. Paste this entire file into the editor (replace the default code)
 *  3. Save (Ctrl+S) and give the project a name, e.g. "BART BRET Data"
 *  4. Click "Deploy" → "New deployment"
 *     - Type: Web app
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  5. Click "Deploy", accept permissions, and copy the Web App URL
 *  6. In tools/bart/index.html AND tools/bret/index.html, find:
 *       var DATA_ENDPOINT = '';
 *     and replace the empty string with your URL:
 *       var DATA_ENDPOINT = 'https://script.google.com/macros/s/YOUR_ID/exec';
 *  7. Commit and push — data collection is live.
 *
 * The script automatically creates a Google Sheet called
 * "BART & BRET — Research Data" on the first request.
 * Columns: Timestamp | Game | Session | Condition | Rounds | Total Earned | Score | Raw JSON
 */

// ── Auto-creates a Google Sheet on first run, reuses it afterwards ────────────
function getOrCreateSheet() {
  var props = PropertiesService.getScriptProperties();
  var id    = props.getProperty('SHEET_ID');

  if (id) {
    try {
      return SpreadsheetApp.openById(id).getActiveSheet();
    } catch (e) {
      // Sheet was deleted — fall through to create a new one
    }
  }

  var ss = SpreadsheetApp.create('BART & BRET — Research Data');
  props.setProperty('SHEET_ID', ss.getId());
  return ss.getActiveSheet();
}

// ── Main POST handler — called by the games ───────────────────────────────────
function doPost(e) {
  try {
    var sheet = getOrCreateSheet();
    var data  = JSON.parse(e.postData.contents);

    // Ensure header row exists
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp', 'Game', 'Session ID', 'Condition',
        'Rounds (summary)', 'Total Earned (£)', 'Adj. Score', 'Raw JSON'
      ]);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    }

    // Build a human-readable rounds summary
    var roundsSummary = '';
    if (data.rounds && Array.isArray(data.rounds)) {
      roundsSummary = data.rounds.map(function(r) {
        if (data.game === 'BART') {
          return 'R' + r.round + ': ' + r.pumps + ' pumps, ' + (r.burst ? 'burst' : '£' + r.earned.toFixed(2));
        } else if (data.game === 'BRET') {
          return 'R' + r.round + ': ' + r.boxes + ' boxes, ' + (r.bombs_hit > 0 ? 'bomb' : '£' + r.earned.toFixed(2));
        } else {
          return JSON.stringify(r);
        }
      }).join(' | ');
    }

    sheet.appendRow([
      data.timestamp      || new Date().toISOString(),
      data.game           || '?',
      data.session_id     || '?',
      data.condition      || 'risk',
      roundsSummary,
      data.total_earned   != null ? Number(data.total_earned).toFixed(2) : '',
      data.adj_bart_score || '',
      JSON.stringify(data)
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Health check — open the URL in a browser to confirm it's live ─────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'alive', message: 'BART/BRET data endpoint is running' }))
    .setMimeType(ContentService.MimeType.JSON);
}
