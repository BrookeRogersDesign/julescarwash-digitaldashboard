// Jules Campaign Hub — Google Apps Script backend
// Lives in the "Jules Campaign Hub" Google Sheet (Extensions → Apps Script).
// The portal's Campaign Hub reads and writes through this script.
//
// Setup:
//   1. Set PORTAL_CODE (what Marissa uses to log in) and TEAM_CODE (Brooke & Abby;
//      also logs into the portal, with editing turned on).
//   2. Deploy → New deployment → Web app. Execute as: Me. Who has access: Anyone.
//   3. Copy the web app URL into HUB_API in index.html.
// After editing this script later, use Deploy → Manage deployments → Edit → New version
// so the same URL keeps working.

const PORTAL_CODE = 'pick-a-client-passcode';
const TEAM_CODE = 'pick-a-passcode';

const TABS = {
  Wishes:    ['id', 'month', 'text', 'by', 'added'],
  Ideas:     ['id', 'month', 'text', 'by', 'added'],  // blank month = Idea Bank
  Campaigns: ['id', 'month', 'start', 'end', 'status', 'name', 'goal', 'leadTarget', 'offer', 'audience', 'channels', 'dates', 'content', 'notes', 'updated']
};
const TEXT_COLS = ['month', 'start', 'end'];  // kept as plain text so Sheets doesn't turn them into dates
const STATUSES = ['wishlist', 'planning', 'ready', 'live', 'done'];
const MAX_LEN = 3000;

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(TABS[name]);
    sh.setFrozenRows(1);
  }
  // Add any columns a newer version of this script expects
  const head = header_(sh);
  TABS[name].filter(h => !head.includes(h)).forEach(h => sh.getRange(1, sh.getLastColumn() + 1).setValue(h));
  const full = header_(sh);
  sh.getRange(1, 1, 1, full.length).setFontWeight('bold');
  TEXT_COLS.forEach(h => {
    const col = full.indexOf(h);
    if (col >= 0) sh.getRange(1, col + 1, sh.getMaxRows(), 1).setNumberFormat('@');
  });
  return sh;
}

function header_(sh) {
  return sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
}

// Writes obj into the sheet by column name, appending or replacing the row with the same id
function put_(name, obj) {
  const sh = sheet_(name);
  const head = header_(sh);
  const row = head.map(h => obj[h] == null ? '' : obj[h]);
  const idCol = head.indexOf('id') + 1;
  if (idCol && sh.getLastRow() > 1) {
    const ids = sh.getRange(2, idCol, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]));
    const i = ids.indexOf(String(obj.id));
    if (i >= 0) return sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
  }
  sh.appendRow(row);
}

function rows_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  const head = values.shift();
  return values
    .filter(r => r.some(v => v !== ''))
    .map(r => Object.fromEntries(head.map((h, i) => [h, cell_(h, r[i])])));
}

function cell_(head, v) {
  if (!(v instanceof Date)) return String(v);
  const tz = Session.getScriptTimeZone();
  if (head === 'month') return Utilities.formatDate(v, tz, 'yyyy-MM');
  if (head === 'start' || head === 'end') return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return v.toISOString();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clean_(v) {
  return String(v == null ? '' : v).trim().slice(0, MAX_LEN);
}

function isMonth_(m) {
  return /^\d{4}-(0[1-9]|1[0-2])$/.test(m);
}

function isDay_(d) {
  return /^\d{4}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$/.test(d);
}

// 'team', 'client', or throws
function role_(code) {
  if (TEAM_CODE === 'pick-a-passcode' || PORTAL_CODE === 'pick-a-client-passcode') {
    throw new Error('Set PORTAL_CODE and TEAM_CODE in the Apps Script first.');
  }
  const c = clean_(code);
  if (c && c === TEAM_CODE) return 'team';
  if (c && c === PORTAL_CODE) return 'client';
  throw new Error('Wrong passcode.');
}

function checkTeam_(code) {
  if (role_(code) !== 'team') throw new Error('That needs the team passcode.');
}

function deleteById_(name, id) {
  const sh = sheet_(name);
  if (sh.getLastRow() < 2) return;
  const idCol = header_(sh).indexOf('id') + 1;
  const ids = sh.getRange(2, idCol, sh.getLastRow() - 1, 1).getValues();
  const i = ids.findIndex(r => String(r[0]) === id);
  if (i >= 0) sh.deleteRow(i + 2);
}

// Everything goes through POST so passcodes never sit in a URL
function doGet() {
  return json_({ ok: false, error: 'Not available.' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const now = new Date();

    if (body.action === 'login') {
      return json_({ ok: true, role: role_(body.code) });

    } else if (body.action === 'list') {
      role_(body.code);
      return json_({ ok: true, wishes: rows_('Wishes'), ideas: rows_('Ideas'), campaigns: rows_('Campaigns') });

    } else if (body.action === 'addWish') {
      role_(body.code);
      const month = clean_(body.month), text = clean_(body.text);
      if (!isMonth_(month) || !text) throw new Error('Pick a month and write what you want.');
      put_('Wishes', { id: Utilities.getUuid(), month, text, by: clean_(body.by) || 'Marissa', added: now });

    } else if (body.action === 'addIdea') {
      role_(body.code);
      const text = clean_(body.text), month = clean_(body.month);
      if (!text) throw new Error('Write the idea first.');
      if (month && !isMonth_(month)) throw new Error('Pick a month or the Idea Bank.');
      put_('Ideas', { id: Utilities.getUuid(), month, text, by: clean_(body.by), added: now });

    } else if (body.action === 'saveCampaign') {
      checkTeam_(body.code);
      const start = clean_(body.start), end = clean_(body.end) || start;
      if (start && !isDay_(start)) throw new Error('Start date looks wrong.');
      if (end && !isDay_(end)) throw new Error('End date looks wrong.');
      if (start && end < start) throw new Error('End date is before the start date.');
      // A campaign with dates lives in the month it starts
      const month = start ? start.slice(0, 7) : clean_(body.month);
      if (!isMonth_(month)) throw new Error('Pick a month or a start date.');
      const c = { id: clean_(body.id) || Utilities.getUuid(), month, start, end: start ? end : '',
                  status: STATUSES.includes(body.status) ? body.status : 'planning', updated: now };
      ['name', 'goal', 'leadTarget', 'offer', 'audience', 'channels', 'dates', 'content', 'notes'].forEach(k => c[k] = clean_(body[k]));
      put_('Campaigns', c);

    } else if (body.action === 'deleteWish') {
      checkTeam_(body.code);
      deleteById_('Wishes', clean_(body.id));

    } else if (body.action === 'deleteCampaign') {
      checkTeam_(body.code);
      deleteById_('Campaigns', clean_(body.id));

    } else if (body.action === 'deleteIdea') {
      checkTeam_(body.code);
      deleteById_('Ideas', clean_(body.id));

    } else {
      throw new Error('Unknown action.');
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}
