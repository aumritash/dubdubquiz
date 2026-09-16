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

// ?leaderboard=<category>&lead=<id> returns the ranked board for that category.
// ?landing=1&ref=<id> returns the counts for the landing page and, if a
// referral code is given, that player's best run. Any other GET just
// confirms the deployment is live.
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.leaderboard) return reply(leaderboard(String(p.leaderboard), String(p.lead || '')));
  if (p.landing) return reply(landing(String(p.ref || '')));
  return reply({ ok: true, service: 'dubdub-quiz', tabs: Object.keys(TABS) });
}

/**
 * Every visitor hits this, so it is cached for a minute: the counts under
 * one key, each referrer's best run under its own. Like the leaderboard it
 * only ever returns a display name, category, score and time.
 */
function landing(ref) {
  var cache = CacheService.getScriptCache();

  var counts = cache.get('counts');
  if (counts) {
    counts = JSON.parse(counts);
  } else {
    var all = readSheets();
    var players = {}, colleges = {};
    all.attempts.forEach(function (a) {
      players[a.lead] = 1;
      var l = all.leads[a.lead];
      if (l && l.college) colleges[l.college.toLowerCase()] = 1;
    });
    counts = { players: Object.keys(players).length, colleges: Object.keys(colleges).length };
    cache.put('counts', JSON.stringify(counts), 60);
  }

  var referrer = null;
  if (ref) {
    var hit = cache.get('ref:' + ref);
    if (hit) {
      referrer = JSON.parse(hit);
    } else {
      var data = readSheets();
      var lead = data.leads[ref];
      if (lead) {
        var best = null;
        data.attempts.forEach(function (a) {
          if (a.lead !== ref) return;
          if (!best || a.score > best.score || (a.score === best.score && a.time < best.time)) best = a;
        });
        if (best) referrer = { name: lead.name, category: best.category, score: best.score,
                               time: isFinite(best.time) ? best.time : null };
      }
      cache.put('ref:' + ref, JSON.stringify(referrer), 60);
    }
  }

  return { ok: true, players: counts.players, colleges: counts.colleges, referrer: referrer };
}

// One read of both tabs, in the shapes the landing needs.
function readSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leads = {}, attempts = [];

  var ls = ss.getSheetByName('leads');
  if (ls && ls.getLastRow() > 1) {
    var iId = TABS.leads.indexOf('id'), iName = TABS.leads.indexOf('name'), iCol = TABS.leads.indexOf('city');
    ls.getRange(2, 1, ls.getLastRow() - 1, TABS.leads.length).getValues().forEach(function (r) {
      leads[String(r[iId])] = { name: displayName(r[iName]), college: String(r[iCol] || '').replace(/\s+/g, ' ').trim() };
    });
  }

  var as = ss.getSheetByName('attempts');
  if (as && as.getLastRow() > 1) {
    var iLead = TABS.attempts.indexOf('lead_id'), iCat = TABS.attempts.indexOf('category'),
        iScore = TABS.attempts.indexOf('score'), iTime = TABS.attempts.indexOf('time_seconds');
    as.getRange(2, 1, as.getLastRow() - 1, TABS.attempts.length).getValues().forEach(function (r) {
      var time = Number(r[iTime]);
      attempts.push({ lead: String(r[iLead]), category: String(r[iCat]),
                      score: Number(r[iScore]) || 0, time: time > 0 ? time : Infinity });
    });
  }
  return { leads: leads, attempts: attempts };
}

/**
 * Best attempt per player in one category, ranked by score then by time,
 * plus the same ranking rolled up by college. Only a display name, college,
 * score and time ever leave the sheet: this endpoint is public, so contact
 * details must never be part of the response.
 */
function leaderboard(category, leadId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var names = {}, colleges = {};
  var leads = ss.getSheetByName('leads');
  if (leads && leads.getLastRow() > 1) {
    var lrows = leads.getRange(2, 1, leads.getLastRow() - 1, TABS.leads.length).getValues();
    var iId = TABS.leads.indexOf('id'), iName = TABS.leads.indexOf('name'),
        iCollege = TABS.leads.indexOf('city');   // the column still carries its original header
    lrows.forEach(function (r) {
      var id = String(r[iId]);
      names[id] = displayName(r[iName]);
      colleges[id] = String(r[iCollege] || '').replace(/\s+/g, ' ').trim();
    });
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

  // Roll the same best-runs up by college. A college stands on its single
  // best run; player count is shown and only breaks exact ties.
  var byCollege = {};
  rows.forEach(function (r) {
    var college = colleges[r.id];
    if (!college) return;                       // the field is optional
    var key = college.toLowerCase();
    var c = byCollege[key];
    if (!c) c = byCollege[key] = { name: college, score: -1, time: Infinity, by: '', players: 0 };
    c.players++;
    if (r.score > c.score || (r.score === c.score && r.time < c.time)) {
      c.score = r.score; c.time = r.time; c.by = names[r.id] || 'Player';
    }
  });
  var crow = Object.keys(byCollege).map(function (k) { return byCollege[k]; });
  crow.sort(function (a, b) { return (b.score - a.score) || (a.time - b.time) || (b.players - a.players); });

  var myCollege = (colleges[leadId] || '').toLowerCase();
  var ctop = [], yours = null;
  crow.forEach(function (c, i) {
    var entry = {
      rank: i + 1, name: c.name, score: c.score,
      time: isFinite(c.time) ? c.time : null, by: c.by, players: c.players
    };
    if (i < LEADERBOARD_SIZE) ctop.push(entry);
    if (myCollege && c.name.toLowerCase() === myCollege) yours = entry;
  });

  return { ok: true, category: category, total: rows.length, top: top, you: you,
           colleges: ctop, your_college: yours };
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
