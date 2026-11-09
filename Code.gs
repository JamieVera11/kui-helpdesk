/**
 * KUI Employee Helpdesk – server side (Google Apps Script, free with a Gmail account).
 *
 * Where things are kept (all in the Google account that deploys this script):
 *  - a Google Sheet "KUI Helpdesk – data" (made automatically the first time):
 *      "Requests" tab: every request in plain columns, refreshed on each save (read, filter or download it; do not type in it)
 *      "kv" and "posts" tabs: the helpdesk's own data. Do not edit these by hand.
 *  - a Drive folder "KUI Helpdesk – memos and documents" for the files posted on the bulletin board (viewable by link).
 *  - a Drive folder "KUI Helpdesk – employee proofs" for payslips and other proof employees upload with a request.
 *      These files are NOT shared: they open only inside the helpdesk.
 *
 * Deploy: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone
 *   ("Anyone" is needed so the helpdesk page on GitHub Pages can reach it.)
 *
 * HELPDESK PASSCODE (required): Project Settings (gear icon) > Script Properties > Add script property
 *   Property: HELPDESK_PASSCODE    Value: the passcode you choose (at least 6 characters).
 *   Every request must carry it, so the link alone opens nothing. Change the value any time: everyone must type the new one.
 */
const META_KEY = 'kuiTrial.v2.meta';   // the helpdesk settings document (name kept for compatibility)
const CHUNK = 45000;                   // a Sheet cell holds at most 50,000 characters: long values are split over several cells
const PROPS = PropertiesService.getScriptProperties();
const FORM_NAMES = {LEAVE: 'Leave Application', CS: 'Change of Schedule', ME: 'Manual Time In / Out', OT: 'Overtime Authorization',
  OB: 'Official Business', SC: 'Salary Concern', SO: 'Temporary Stay-Out Pass', SR: 'Return from Stay-Out', LS: 'Luggage / Item Slip',
  GP: 'Overnight Permit / Gate Pass', PAYROLL: 'Payroll Request'};

function doGet() {
  // the helpdesk page lives on GitHub Pages; this link only answers the page, with the passcode
  return HtmlService.createHtmlOutput('<p style="font:15px Arial,sans-serif;margin:40px">KUI Employee Helpdesk: open the helpdesk from its own link.</p>');
}
function passOk_(pass) {
  const want = String(PROPS.getProperty('HELPDESK_PASSCODE') || '').trim();
  if (!want) throw new Error('WRONG_PASSCODE: no passcode is set yet. HR: add HELPDESK_PASSCODE in Project Settings > Script Properties.');
  if (String(pass || '') !== want) { Utilities.sleep(1500); throw new Error('WRONG_PASSCODE'); }
}

// ---------- storage helpers ----------
function ss_() {
  const id = PROPS.getProperty('SS');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const s = SpreadsheetApp.create('KUI Helpdesk – data');
  PROPS.setProperty('SS', s.getId());
  return s;
}
function sheet_(name) { const s = ss_(); return s.getSheetByName(name) || s.insertSheet(name); }
function values_(sh) { return sh.getLastRow() ? sh.getDataRange().getValues() : []; }
function ver_() { return Number(PROPS.getProperty('VER') || 0); }
function bump_() { const v = ver_() + 1; PROPS.setProperty('VER', String(v)); return v; }
function withLock_(f) { const lock = LockService.getScriptLock(); lock.waitLock(20000); try { return f(); } finally { lock.releaseLock(); } }
const yes_ = x => x === 1 || x === '1' || x === true;
// every stored text starts with "~" so the Sheet never turns it into a number, a date or a formula
const cell_ = s => '~' + s;
const uncell_ = x => String(x).slice(1);

function kvRead_() {
  const sh = sheet_('kv'), m = {};
  values_(sh).forEach((r, i) => {
    if (!r[0]) return;
    m[r[0]] = {row: i + 1, ver: Number(r[1]) || 0, del: yes_(r[2]), json: r.slice(3).filter(x => x !== '' && x != null).map(uncell_).join('')};
  });
  return {sh, m};
}
function revOf_(json) { try { return JSON.parse(json).rev || 0; } catch (e) { return -1; } }

// ---------- called by the page ----------
/** Everything, when the page opens. */
function kvAll() {
  const m = kvRead_().m, docs = {};
  for (const k in m) if (!m[k].del) docs[k] = m[k].json;
  return {ver: ver_(), docs: docs, posts: postsList_()};
}

/** What changed since version v (the page asks every few seconds). */
function kvSince(v) {
  const m = kvRead_().m, docs = {};
  for (const k in m) if (m[k].ver > v) docs[k] = m[k].del ? null : m[k].json;
  const pv = Number(PROPS.getProperty('PVER') || 0);
  return {ver: ver_(), docs: docs, posts: pv > v ? postsList_() : null};
}

/**
 * Save one batch. base = the revision the page started from (-1 = no check).
 * If someone else saved first, nothing is written and the page gets the current data back.
 */
function kvCommit(base, writes) {
  return withLock_(() => {
    const kv = kvRead_(), sh = kv.sh, m = kv.m;
    if (base >= 0) {
      const cur = m[META_KEY], srv = cur && !cur.del ? revOf_(cur.json) : 0;
      if (srv !== base) {
        const docs = {};
        for (const k in m) if (!m[k].del) docs[k] = m[k].json;
        return {conflict: true, ver: ver_(), docs: docs};
      }
    }
    const v = bump_();
    let tickets = false;
    Object.keys(writes).forEach(k => {
      const val = writes[k], cells = [k, v, val === null ? 1 : 0];
      if (val !== null) for (let i = 0; i < val.length; i += CHUNK) cells.push(cell_(val.slice(i, i + CHUNK)));
      if (/\.t\./.test(k)) tickets = true;
      const row = m[k] ? m[k].row : sh.getLastRow() + 1;
      if (m[k]) sh.getRange(row, 1, 1, sh.getMaxColumns()).clearContent();
      sh.getRange(row, 1, 1, cells.length).setValues([cells]);
      m[k] = {row: row, ver: v, del: val === null, json: val || ''};
    });
    if (tickets) { try { requestsSheet_(m); } catch (e) { console.error(e); } }
    return {ok: true, ver: v};
  });
}

/** The Google Sheet link, for the Administrator. */
function sheetUrl() { return ss_().getUrl(); }

// ---------- readable "Requests" tab ----------
function requestsSheet_(m) {
  let all = [];
  for (const k in m) if (/\.t\./.test(k) && !m[k].del) { try { all = all.concat(JSON.parse(m[k].json)); } catch (e) {} }
  all.sort((a, b) => (a.filed || 0) - (b.filed || 0));
  const tz = 'Asia/Manila', when = t => t ? Utilities.formatDate(new Date(t), tz, 'yyyy-MM-dd HH:mm') : '';
  const txt = x => { const s = x == null ? '' : String(x); return /^[=+\-@]/.test(s) ? "'" + s : s; };
  const head = ['Ticket', 'Filed (PH time)', 'Emp ID', 'Employee', 'Department', 'Request', 'Date(s)', 'Days', 'Reason',
    'Approval', 'Approved / rejected by', 'Approval date', 'Approval remark', 'Processed by (dept)', 'Processing', 'Processed / held by',
    'Processing date', 'Processing remark', 'HR status'];
  const rows = all.map(t => {
    const a = t.approval || {}, p = t.proc || {}, d = t.data || {}, e = t.emp || {};
    const from = t.from || d.dFrom || d.ddate || d.date || '', to = t.to || d.dTo || d.rdate || '';
    const req = t.code === 'LEAVE' ? 'Leave – ' + (t.typeText || t.type || '') : (FORM_NAMES[t.code] || t.code || '');
    const ap = t.code === 'SC' && a.status === 'Approved' ? 'Forwarded to Payroll by HR' : t.code === 'SC' && a.status === 'Rejected' ? 'Not forwarded to Payroll by HR'
      : {NotRequired: 'No approval needed', Counterpart: 'Waiting for counterpart', Declined: 'Declined by counterpart'}[a.status] || a.status || 'Pending';
    const pst = p.result ? 'Payroll: ' + String(p.result).toUpperCase() : (p.status || '');
    return [t.no, when(t.filed), "'" + (e.id || ''), e.name, t.dept || e.dept, req, from + (to && to !== from ? ' to ' + to : ''), t.days || d.days || '',
      t.reason || d.reason || '', ap, a.by || '', when(a.at), a.remark || '', p.dept || '', pst, p.by || '', when(p.at), p.remark || '', t.hrStatus || '']
      .map((x, i) => i === 2 ? x : txt(x));
  });
  const sh = sheet_('Requests');
  sh.clearContents();
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  sh.setFrozenRows(1);
}

// ---------- bulletin board ----------
function postsRead_() {
  const sh = sheet_('posts'), m = {};
  values_(sh).forEach((r, i) => { if (r[0]) m[r[0]] = {row: i + 1, del: yes_(r[1]), json: r[2] ? uncell_(r[2]) : ''}; });
  return {sh, m};
}
function postsList_() {
  const m = postsRead_().m, out = [];
  for (const id in m) if (!m[id].del) { try { const p = JSON.parse(m[id].json); p.id = id; out.push(p); } catch (e) {} }
  return out;
}
function postSave(id, d) {
  return withLock_(() => {
    const r = postsRead_(), row = r.m[id] ? r.m[id].row : r.sh.getLastRow() + 1;
    r.sh.getRange(row, 1, 1, 3).setValues([[id, 0, cell_(JSON.stringify(d))]]);
    PROPS.setProperty('PVER', String(bump_()));
    return postsList_();
  });
}
function postDelete(id) {
  return withLock_(() => {
    const r = postsRead_(), p = r.m[id];
    if (!p) return postsList_();
    try { const d = JSON.parse(p.json); if (d.file && d.file.gid) DriveApp.getFileById(d.file.gid).setTrashed(true); } catch (e) {}
    r.sh.getRange(p.row, 1, 1, 3).setValues([[id, 1, '']]);
    PROPS.setProperty('PVER', String(bump_()));
    return postsList_();
  });
}
function folder_() {
  const id = PROPS.getProperty('FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('KUI Helpdesk – memos and documents');
  PROPS.setProperty('FOLDER', f.getId());
  return f;
}
/** Save an approved memo (PDF or image) in the Drive folder; anyone with the link can view it. */
function uploadFile(name, type, b64) {
  if (!/^(application\/pdf|image\/(png|jpeg|webp))$/.test(type)) throw new Error('Only PDF or image files.');
  const f = folder_().createFile(Utilities.newBlob(Utilities.base64Decode(b64), type, name));
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (e) { try { f.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (x) {} }
  return {gid: f.getId(), name: name, type: type, size: f.getSize()};
}

// ---------- employee proofs (payslips, chat screenshots): private, opened only through the helpdesk ----------
function proofFolder_() {
  const id = PROPS.getProperty('PROOFS');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('KUI Helpdesk – employee proofs');
  PROPS.setProperty('PROOFS', f.getId());
  return f;
}
/** Save a payslip or proof uploaded with a request. The file is not shared with anyone. */
function uploadProof(name, type, b64) {
  if (!/^(application\/pdf|image\/(png|jpeg|webp))$/.test(type)) throw new Error('Only a photo (JPG, PNG) or a PDF.');
  const bytes = Utilities.base64Decode(b64);
  if (bytes.length > 10 * 1048576) throw new Error('The file is over 10 MB.');
  const f = proofFolder_().createFile(Utilities.newBlob(bytes, type, name));
  return {gid: f.getId(), name: name, type: type, size: f.getSize()};
}
/** Send a helpdesk file (a proof or a posted memo) back to the page, so it opens inside the helpdesk. */
function fileData(gid) {
  const f = DriveApp.getFileById(gid), ok = [PROPS.getProperty('PROOFS'), PROPS.getProperty('FOLDER')];
  const ps = f.getParents(); let inside = false;
  while (ps.hasNext()) if (ok.indexOf(ps.next().getId()) >= 0) inside = true;
  if (!inside) throw new Error('This file is not part of the helpdesk.');
  const b = f.getBlob();
  return {type: b.getContentType(), b64: Utilities.base64Encode(b.getBytes())};
}

// ---------- connection for the page hosted on GitHub Pages ----------
// The GitHub page sends {fn, args} here (Deploy > Web app > Who has access: Anyone).
const API_FUNCS = {kvAll: kvAll, kvSince: kvSince, kvCommit: kvCommit, sheetUrl: sheetUrl, postSave: postSave, postDelete: postDelete, uploadFile: uploadFile, uploadProof: uploadProof, fileData: fileData};
function doPost(e) {
  let out;
  try {
    const q = JSON.parse(e.postData.contents);
    passOk_(q.pass);
    const f = API_FUNCS[q.fn];
    if (!f) throw new Error('Unknown request');
    out = {ok: true, r: f.apply(null, q.args || [])};
  } catch (err) {
    out = {ok: false, error: String((err && err.message) || err)};
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
