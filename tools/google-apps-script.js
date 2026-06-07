/**
 * Google Apps Script — Data collection endpoint for BART & BRET demos
 * =====================================================================
 * Creates one tab per game type (BART, BRET, DICTATOR, ...) automatically.
 * All data lands in one file: "BART & BRET — Research Data" in Google Drive.
 */

// ── Get or create the main spreadsheet ───────────────────────────────────────
function getSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var id    = props.getProperty('SHEET_ID');

  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {}
  }

  var ss = SpreadsheetApp.create('BART & BRET — Research Data');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

// ── Get or create the tab for this game ──────────────────────────────────────
function getGameSheet(ss, gameName) {
  var tab = ss.getSheetByName(gameName);

  if (!tab) {
    tab = ss.insertSheet(gameName);

    // Write header row for this game
    var headers = {
      'BART':     ['Timestamp', 'Session ID', 'Round 1 Pumps', 'Round 1 Burst', 'Round 1 Earned',
                   'Round 2 Pumps', 'Round 2 Burst', 'Round 2 Earned',
                   'Round 3 Pumps', 'Round 3 Burst', 'Round 3 Earned',
                   'Total Earned (£)', 'Adj. BART Score', 'Max Pumps'],
      'BRET':     ['Timestamp', 'Session ID', 'Condition',
                   'Round 1 Boxes', 'Round 1 Bomb Hit', 'Round 1 Earned',
                   'Round 2 Boxes', 'Round 2 Bomb Hit', 'Round 2 Earned',
                   'Round 3 Boxes', 'Round 3 Bomb Hit', 'Round 3 Earned',
                   'Total Earned (£)', 'Payoff Round'],
      'DICTATOR': ['Timestamp', 'Session ID', 'Tokens Given', 'Tokens Kept', 'Share Given (%)']
    };

    var h = headers[gameName] || ['Timestamp', 'Session ID', 'Raw JSON'];
    tab.appendRow(h);
    tab.getRange(1, 1, 1, h.length).setFontWeight('bold');
  }

  return tab;
}

// ── Main POST handler ─────────────────────────────────────────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var game = (data.game || 'UNKNOWN').toUpperCase();
    var ss   = getSpreadsheet();
    var tab  = getGameSheet(ss, game);
    var ts   = data.timestamp || new Date().toISOString();
    var sid  = data.session_id || '?';
    var row  = [];

    if (game === 'BART') {
      var rounds = data.rounds || [];
      row = [ts, sid];
      for (var i = 0; i < 3; i++) {
        var r = rounds[i];
        row.push(r ? r.pumps   : '');
        row.push(r ? (r.burst ? 'yes' : 'no') : '');
        row.push(r ? r.earned  : '');
      }
      row.push(data.total_earned != null ? Number(data.total_earned).toFixed(2) : '');
      row.push(data.adj_bart_score || '');
      row.push(data.max_pumps || 32);

    } else if (game === 'BRET') {
      var rounds = data.rounds || [];
      row = [ts, sid, data.condition || 'risk'];
      for (var i = 0; i < 3; i++) {
        var r = rounds[i];
        row.push(r ? r.boxes     : '');
        row.push(r ? (r.bombs_hit > 0 ? 'yes' : 'no') : '');
        row.push(r ? r.earned    : '');
      }
      row.push(data.total_earned != null ? Number(data.total_earned).toFixed(2) : '');
      row.push(data.payoff_round || '');

    } else if (game === 'DICTATOR') {
      row = [
        ts, sid,
        data.tokens_given != null ? data.tokens_given : '',
        data.tokens_kept  != null ? data.tokens_kept  : '',
        data.share_given  != null ? (data.share_given * 100).toFixed(1) + '%' : ''
      ];

    } else {
      row = [ts, sid, JSON.stringify(data)];
    }

    tab.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Health check ──────────────────────────────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'alive', message: 'Data endpoint running — separate tabs per game' }))
    .setMimeType(ContentService.MimeType.JSON);
}
