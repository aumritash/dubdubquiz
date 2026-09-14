/**
 * dubdub quiz — Google Sheets backend
 *
 * Setup:
 *  1. Create a Google Sheet. Name it whatever you like.
 *  2. Extensions → Apps Script. Delete the placeholder code.
 *  3. Paste this whole file in and save.
 *  4. Deploy → New deployment → type "Web app".
 *       Execute as:        Me
 *       Who has access:    Anyone
 *     (It must be "Anyone", not "Anyone with a Google account",
 *      or players will hit a login wall and nothing will save.)
 *  5. Authorise when prompted. Google will warn the app is unverified;
 *     click Advanced → Go to <project> (unsafe). It's your own script.
 *  6. Copy the Web app URL. It ends in /exec.
 *  7. Paste it into SHEET_ENDPOINT in index.html.
 *
 * Re-deploying after an edit: Deploy → Manage deployments → pencil icon →
 * Version: New version → Deploy. Editing the code alone changes nothing live.
 */

var TABS = {
  leads:           ['created_at', 'id', 'name', 'contact', 'city', 'is_email', 'referred_by'],
  attempts:        ['created_at', 'lead_id', 'category', 'score', 'total', 'tier', 'time_seconds'],
  shares:          ['created_at', 'lead_id', 'channel'],
  referral_visits: ['created_at', 'ref', 'visit_id']
};

var LEADERBOARD_SIZE = 10;

function doPost(e) {
  // Writes are serialised. Without this, two players finishing at the same
  // moment can land on the same row and one of them is lost.
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return reply({ ok: false, error: 'busy' });
  }

  try {
    var payload = JSON.parse(e.postData.contents);
    var cols = TABS[payload.table];
    if (!cols) return reply({ ok: false, error: 'unknown table' });

    var sheet = tab(payload.table, cols);
    var row = payload.row || {};

    sheet.appendRow(cols.map(function (c) {
      if (c === 'created_at') return new Date();
      return row[c] === undefined ? '' : row[c];
    }));

    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ?leaderboard=<category>&lead=<id> returns the ranked board for that
// category. Any other GET just confirms the deployment is live.
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.leaderboard) return reply(leaderboard(String(p.leaderboard), String(p.lead || '')));
  return reply({ ok: true, service: 'dubdub-quiz', tabs: Object.keys(TABS) });
}

/**
 * Best attempt per player in one category, ranked by score then by time.
 * Only a display name, score and time ever leave the sheet: this endpoint
 * is public, so contact details must never be part of the response.
 */
function leaderboard(category, leadId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var names = {};
  var leads = ss.getSheetByName('leads');
  if (leads && leads.getLastRow() > 1) {
    var lrows = leads.getRange(2, 1, leads.getLastRow() - 1, TABS.leads.length).getValues();
    var iId = TABS.leads.indexOf('id'), iName = TABS.leads.indexOf('name');
    lrows.forEach(function (r) { names[String(r[iId])] = displayName(r[iName]); });
  }

  var best = {};
  var attempts = ss.getSheetByName('attempts');
  if (attempts && attempts.getLastRow() > 1) {
    var arows = attempts.getRange(2, 1, attempts.getLastRow() - 1, TABS.attempts.length).getValues();
    var iLead = TABS.attempts.indexOf('lead_id'),   iCat  = TABS.attempts.indexOf('category'),
        iScore = TABS.attempts.indexOf('score'),    iTime = TABS.attempts.indexOf('time_seconds');

    arows.forEach(function (r) {
      if (String(r[iCat]) !== category) return;
      var id = String(r[iLead]);
      var score = Number(r[iScore]) || 0;
      var time = Number(r[iTime]);
      if (!(time > 0)) time = Infinity;   // rows from before the timer existed sort last on ties

      var cur = best[id];
      if (!cur || score > cur.score || (score === cur.score && time < cur.time)) {
        best[id] = { id: id, score: score, time: time };
      }
    });
  }

  var rows = Object.keys(best).map(function (k) { return best[k]; });
  rows.sort(function (a, b) { return (b.score - a.score) || (a.time - b.time); });

  var top = [], you = null;
  rows.forEach(function (r, i) {
    var entry = {
      rank:  i + 1,
      name:  names[r.id] || 'Player',
      score: r.score,
      time:  isFinite(r.time) ? r.time : null
    };
    if (i < LEADERBOARD_SIZE) top.push(entry);
    if (r.id === leadId) you = entry;
  });

  return { ok: true, category: category, total: rows.length, top: top, you: you };
}

// "Aumritash Maitra" → "Aumritash M.", "Priya" → "Priya". Enough to tell
// two players apart without publishing anyone's full name.
function displayName(full) {
  var parts = String(full || '').trim().split(/\s+/);
  if (!parts[0]) return 'Player';
  if (parts.length === 1) return parts[0];
  return parts[0] + ' ' + parts[1].charAt(0).toUpperCase() + '.';
}

function tab(name, cols) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(cols);
    sheet.getRange(1, 1, 1, cols.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function reply(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
