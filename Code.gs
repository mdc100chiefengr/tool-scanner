/**
 * Tools Inventory — backend API for the hosted scanner page (index.html)
 *
 * Sheets required in this spreadsheet:
 *   Tools        : ToolID | Name | ... | Status | CurrentHolder | CheckedOutAt | DueBack | ... | CalibrationDue
 *   Log          : Time | ToolID | Action | User | Notes
 *   Technicians  : Name | PIN        (format the PIN column as Plain text so leading zeros survive)
 *
 * Deploy: Deploy > New deployment > Web app
 *   Execute as: Me        Who has access: Anyone
 * Paste the /exec URL into API_URL in index.html.
 * After any change to this file, use Deploy > Manage deployments > Edit > New version.
 */

const SS = SpreadsheetApp.getActive();
const TOOLS = SS.getSheetByName('Tools');
const LOG = SS.getSheetByName('Log');
const TECHS = SS.getSheetByName('Technicians');

const DUE_HOURS = 24;

// ---------- helpers ----------
function norm_(h) { return String(h).trim().toLowerCase().replace(/\s+/g, ''); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function colIndex_(headers) {
  const m = {};
  headers.forEach((h, i) => m[norm_(h)] = i);
  return m;
}

function need_(m, name) {
  const i = m[name.toLowerCase()];
  if (i === undefined) throw new Error('Missing column header: ' + name);
  return i;
}

function checkPin_(name, pin) {
  const rows = TECHS.getDataRange().getValues().slice(1);
  return rows.some(r =>
    String(r[0]).trim() === String(name).trim() &&
    String(r[1]).trim() === String(pin).trim() &&
    String(r[1]).trim() !== '');
}

// ---------- API ----------
function doGet(e) {
  try {
    const action = (e.parameter.action || 'init');
    if (action !== 'init') return json_({ error: 'Unknown action' });

    const data = TOOLS.getDataRange().getValues();
    const m = colIndex_(data[0]);
    const iId = need_(m, 'ToolID'), iName = need_(m, 'Name'),
          iStatus = need_(m, 'Status'), iHolder = need_(m, 'CurrentHolder');

    const tools = data.slice(1)
      .filter(r => r[iId] !== '')
      .map(r => ({
        id: String(r[iId]).trim(),
        name: String(r[iName]),
        status: String(r[iStatus] || 'In'),
        holder: String(r[iHolder] || '')
      }));

    const technicians = TECHS.getDataRange().getValues().slice(1)
      .map(r => String(r[0]).trim()).filter(Boolean);

    return json_({ technicians, tools });
  } catch (err) {
    return json_({ error: err.message });
  }
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.action !== 'batch') return json_({ error: 'Unknown action' });
    if (!checkPin_(req.tech, req.pin)) return json_({ error: 'BAD_PIN' });
    return json_({ results: submitBatch_(req.tech, req.mode, req.ids || [], req.notes || '') });
  } catch (err) {
    return json_({ error: err.message });
  }
}

function submitBatch_(user, action, ids, notes) {
  if (['CHECKOUT', 'RETURN', 'ISSUE'].indexOf(action) < 0) throw new Error('Bad mode: ' + action);
  if (ids.length > 50) throw new Error('Too many tools in one batch (max 50)');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // prevents two scans racing
  try {
    const data = TOOLS.getDataRange().getValues();
    const m = colIndex_(data[0]);
    const cId = need_(m, 'ToolID'), cName = need_(m, 'Name'),
          cStatus = need_(m, 'Status'), cHolder = need_(m, 'CurrentHolder'),
          cOutAt = need_(m, 'CheckedOutAt'), cDue = need_(m, 'DueBack');

    const results = [];
    const seen = {};

    ids.forEach(rawId => {
      const toolId = String(rawId).trim();
      const i = data.findIndex((r, n) => n > 0 &&
        String(r[cId]).trim().toUpperCase() === toolId.toUpperCase());
      if (i < 1) { results.push({ id: toolId, name: '', ok: false, message: 'Unknown tool' }); return; }
      if (seen[i]) return;
      seen[i] = true;

      const row = i + 1;
      const name = String(data[i][cName]);
      const status = String(data[i][cStatus]);
      const now = new Date();

      try {
        if (action === 'CHECKOUT') {
          if (status === 'Out') throw new Error('Already out with ' + data[i][cHolder]);
          if (status === 'Needs Repair') throw new Error('Marked Needs Repair');
          TOOLS.getRange(row, cStatus + 1).setValue('Out');
          TOOLS.getRange(row, cHolder + 1).setValue(user);
          TOOLS.getRange(row, cOutAt + 1).setValue(now);
          TOOLS.getRange(row, cDue + 1).setValue(new Date(now.getTime() + DUE_HOURS * 3600 * 1000));
        } else if (action === 'RETURN') {
          if (status !== 'Out') throw new Error('Was not checked out');
          TOOLS.getRange(row, cStatus + 1).setValue('In');
          TOOLS.getRange(row, cHolder + 1).clearContent();
          TOOLS.getRange(row, cOutAt + 1).clearContent();
          TOOLS.getRange(row, cDue + 1).clearContent();
        } else { // ISSUE
          TOOLS.getRange(row, cStatus + 1).setValue('Needs Repair');
        }
        LOG.appendRow([now, data[i][cId], action, user, notes]);
        results.push({ id: toolId, name, ok: true, message: 'OK' });
      } catch (err) {
        results.push({ id: toolId, name, ok: false, message: err.message });
      }
    });
    return results;
  } finally {
    lock.releaseLock();
  }
}

// ---------- daily report (unchanged) ----------
function dailyCheck() {
  const rows = TOOLS.getDataRange().getValues().slice(1);
  const now = new Date();
  const overdue = rows.filter(r => r[4] === 'Out' && r[7] && new Date(r[7]) < now);
  const needsCal = rows.filter(r => r[9] && new Date(r[9]) < new Date(now.getTime() + 14 * 864e5));

  if (overdue.length || needsCal.length) {
    let body = '';
    if (overdue.length) body += 'OVERDUE:\n' + overdue.map(r => `${r[0]} ${r[1]} → ${r[5]}`).join('\n') + '\n\n';
    if (needsCal.length) body += 'CALIBRATION DUE ≤14 DAYS:\n' + needsCal.map(r => `${r[0]} ${r[1]}`).join('\n');
    MailApp.sendEmail('you@yourdomain.com', 'Tool report', body);
  }
}
// Set a time-driven trigger: Triggers → dailyCheck → Day timer → 7–8 AM
