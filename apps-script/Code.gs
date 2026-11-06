/**
 * ASDD Re-verification – AC-39 RAJINDER NAGAR – Google Apps Script backend
 * Standalone script (not bound to the sheet, so any other script already attached to the sheet is left alone).
 * Works on the Google Sheet below, in its own tabs "ASDD_Reverification" and "Edit_Log" (created automatically).
 * The sheet's other tabs (AC-39 Supervisor-wise, Rejected Entries, Pending Today, ASDD Responses) are never touched.
 *
 * Columns of ASDD_Reverification (one row per elector, keyed by EPIC in column J):
 *  A Last Updated        B District Name   C AC No        D AC Name       E Part No (PS)   F Part Name   G BLO Name
 *  H S.No                I Name            J EPIC No      K Relation Type L Relation Name  M Old Part No N Old Serial No
 *  O Age                 P Gender          Q Status of Elector in the ASDD List
 *  R Status of ASDD Elector during Re-verification by BLOs (Residing/ASDD)        (old "Not Residing" = ASDD)
 *  S If found Residing & eligible for Enrolment, Form-6 Obtained or Not (Yes/No)
 *  T Remarks (old version – kept, no longer written)
 *  U If found Residing & eligible for Enrolment, but Form-6 Not obtained, Reason thereof
 *  V Date(s) of visit by BLO during Enumeration Phase
 *  W Date(s) of visit by BLO during Re-verification
 *  X If Found Residing, Reason for not submitting EF during Enumeration Phase
 *  Y If Form-6 Refused, Whether refused by Elector/Relative (Mention Name)
 *  Z If Form-6 Refused, Whether refused by Elector/Relative (Mention Mobile Number)
 *
 * Website answer array (8 values): [status, form6, enumDates, reverifDates, reasonEF, reasonNoForm6, refusedName, refusedMobile]
 * "get" also returns the old Remarks (column T) as a 9th value, shown to the BLO as a hint.
 */
const SSID  = '1LrL6dZJVUDZCArJ1xuH6UFp9BpLWKr_kCpZUQZLT7rg';
const SITE  = 'https://asdd-ac39.vercel.app';     // used only by prefillFromSite()
const DIST = 'CENTRAL', AC_NO = '39', AC_NAME = 'RAJINDER NAGAR';
const SHEET = 'ASDD_Reverification';
const LOG = 'Edit_Log';
const NCOL = 26;                 // A..Z
const C = { upd:1, ps:5, epic:10, R:18, S:19, T:20, U:21, V:22, W:23, X:24, Y:25, Z:26 };
const HEADERS = {
  18: 'Status of ASDD Elector during Re-verification by BLOs (Residing/ASDD)',
  19: 'If found Residing & eligible for Enrolment, Form-6 Obtained or Not (Yes/No)',
  20: 'Remarks (old version)',
  21: 'If found Residing & eligible for Enrolment, but Form-6 Not obtained, Reason thereof',
  22: 'Date(s) of visit by BLO during Enumeration Phase',
  23: 'Date(s) of visit by BLO during Re-verification',
  24: 'If Found Residing, Reason for not submitting EF during Enumeration Phase',
  25: 'If Form-6 Refused, Whether refused by Elector/Relative (Mention Name)',
  26: 'If Form-6 Refused, Whether refused by Elector/Relative (Mention Mobile Number)'
};
const MOB_RE = /^[6-9]\d{9}$/;

/* ------------------------------------------------------------------ helpers */
function out_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
let ss__ = null;
function ss_(){ return ss__ || (ss__ = SpreadsheetApp.openById(SSID)); }
const MAIN_HEAD = ['Last Updated','District Name','AC No','AC Name','Part No (PS)','Part Name','BLO Name','S.No','Name','EPIC No',
  'Relation Type','Relation Name','Old Part No','Old Serial No','Age','Gender','Status of Elector in the ASDD List',
  HEADERS[18], HEADERS[19], HEADERS[20], HEADERS[21], HEADERS[22], HEADERS[23], HEADERS[24], HEADERS[25], HEADERS[26]];
const LOG_HEAD = ['Time','PS','BLO','EPIC','Name','Old Status','Old Form-6','Old Details','New Status','New Form-6','New Details'];
// Returns the data tab, creating it (with headers) the first time
function sh_(){
  const ss = ss_(); let sh = ss.getSheetByName(SHEET);
  if(!sh){
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, NCOL).setValues([MAIN_HEAD]).setFontWeight('bold').setWrap(true);
    sh.setFrozenRows(1);
    sh.getRange('E:E').setNumberFormat('@'); sh.getRange('H:H').setNumberFormat('@'); sh.getRange('M:N').setNumberFormat('@');
    sh.getRange('U:Z').setNumberFormat('@');
  }
  return sh;
}
function log_(){
  const ss = ss_(); let lg = ss.getSheetByName(LOG);
  if(!lg){ lg = ss.insertSheet(LOG); lg.getRange(1, 1, 1, LOG_HEAD.length).setValues([LOG_HEAD]).setFontWeight('bold'); lg.setFrozenRows(1); }
  return lg;
}
function str_(v){ return v == null ? '' : String(v).trim(); }
function lastRow_(sh){ return Math.max(sh.getLastRow(), 1); }

// Clean one answer exactly like the website does (fields that don't apply are emptied)
function norm_(a){
  a = (a || []).slice(0, 8).map(str_); while(a.length < 8) a.push('');
  if(a[0] === 'Not Residing') a[0] = 'ASDD';
  if(a[0] !== 'Residing' && a[0] !== 'ASDD') a[0] = '';
  if(a[1] !== 'Yes' && a[1] !== 'No') a[1] = '';
  if(a[0] === 'ASDD' || !a[0]){ a[1] = a[4] = a[5] = a[6] = a[7] = ''; }
  else if(a[1] === 'Yes'){ a[5] = a[6] = a[7] = ''; }
  a[2] = firstDate_(a[2]); a[3] = firstDate_(a[3]);          // only one visit date per field
  a[4] = a[4].slice(0, 200); a[5] = a[5].slice(0, 200); a[6] = a[6].slice(0, 100); a[7] = a[7].replace(/\D/g, '').slice(0, 10);
  return a;
}
// Keep only the first date of a comma-separated list ("05-11-2026, 07-11-2026" -> "05-11-2026")
function firstDate_(v){ return str_(String(v || '').split(',')[0]).slice(0, 200); }
function complete_(a){
  if(!a[2] || !a[3]) return false;            // both visit-date fields are mandatory (ASDD and Residing)
  if(a[0] === 'ASDD') return true;
  if(a[0] !== 'Residing' || !a[4] || !a[1]) return false;
  if(a[1] === 'Yes') return true;
  return !!(a[5] && a[6] && MOB_RE.test(a[7]));
}
// Row values R..Z (9 cells, display strings) -> website answer (9 values)
function fromRow_(v){
  // Column U (reason Form-6 not obtained) is prefilled from the old Remarks (T) when it is still empty
  const a = norm_([v[0], v[1], v[4], v[5], v[6], str_(v[3]) || str_(v[2]), v[7], v[8]]);
  a.push(str_(v[2]));            // old Remarks (T)
  return a;
}

/* ------------------------------------------------------------------ GET */
function doGet(e){
  try{
    const p = (e && e.parameter) || {}, act = p.action || 'get';
    if(act === 'get')     return out_(getPS_(String(p.ps || '')));
    if(act === 'summary') return out_(summary_());
    if(act === 'resall')  return out_(resAll_());
    if(act === 'ping')    return out_({ ok:true, v:2, ac:AC_NO });
    return out_({ ok:false, error:'unknown action ' + act });
  }catch(err){ return out_({ ok:false, error:String(err && err.message || err) }); }
}

// Answers of one PS: { ok, data: { EPIC: [9 values] } }. Reads only that PS's rows.
// PS -> contiguous runs of rows [[startIndex, length]], kept in the script cache so loading a Part
// does not have to scan all 53,859 rows every time. Rebuilt when the row count changes.
function psIndex_(sh, n){
  const cache = CacheService.getScriptCache(), key = 'psidx_' + n, hit = cache.get(key);
  if(hit) return JSON.parse(hit);
  const psCol = sh.getRange(2, C.ps, n, 1).getDisplayValues(), idx = {};
  let cur = null, s = 0;
  for(let i = 0; i <= n; i++){
    const p = i < n ? String(psCol[i][0]).trim() : null;
    if(p !== cur){ if(cur) (idx[cur] || (idx[cur] = [])).push([s, i - s]); cur = p; s = i; }
  }
  try{ cache.put(key, JSON.stringify(idx), 21600); }catch(e){}
  return idx;
}
function getPS_(ps){
  if(!ps) return { ok:false, error:'ps missing' };
  const sh = sh_(), n = lastRow_(sh) - 1, data = {};
  if(n < 1) return { ok:true, data:data };
  const runs = psIndex_(sh, n)[ps] || [];
  runs.forEach(function(r){
    const ep = sh.getRange(r[0] + 2, C.epic, r[1], 1).getDisplayValues();
    const v  = sh.getRange(r[0] + 2, C.R, r[1], 9).getDisplayValues();
    for(let i = 0; i < r[1]; i++){ const k = str_(ep[i][0]); if(k) data[k] = fromRow_(v[i]); }
  });
  return { ok:true, data:data };
}

// Per-PS counts for the admin panel and Daily Report (cached for 60 s)
function summary_(){
  const cache = CacheService.getScriptCache(), hit = cache.get('summary_v2');
  if(hit) return JSON.parse(hit);
  const sh = sh_(), n = lastRow_(sh) - 1, data = {}, seen = {};
  if(n >= 1){
    const ps = sh.getRange(2, C.ps, n, 1).getDisplayValues();
    const ep = sh.getRange(2, C.epic, n, 1).getDisplayValues();
    const v  = sh.getRange(2, C.R, n, 9).getDisplayValues();
    for(let i = n - 1; i >= 0; i--){                  // last row of an EPIC wins
      const k = str_(ep[i][0]), p = str_(ps[i][0]); if(!k || !p || seen[k]) continue; seen[k] = 1;
      const a = fromRow_(v[i]); if(!a[0]) continue;
      const o = data[p] || (data[p] = { done:0, residing:0, notResiding:0, f6yes:0, f6no:0, started:0, frozen:false });
      o.started++;
      if(complete_(a)) o.done++;
      if(a[0] === 'Residing'){ o.residing++; if(a[1] === 'Yes') o.f6yes++; else if(a[1] === 'No') o.f6no++; }
      else o.notResiding++;
    }
  }
  const res = { ok:true, data:data };
  try{ cache.put('summary_v2', JSON.stringify(res), 60); }catch(e){}
  return res;
}

// Every Residing elector, PS-wise: { ok, resall:true, data: { ps: [[sno, name, epic, [answer]]] } }
function resAll_(){
  const sh = sh_(), n = lastRow_(sh) - 1, data = {}, seen = {};
  if(n >= 1){
    const ps = sh.getRange(2, C.ps, n, 1).getDisplayValues();
    const hij = sh.getRange(2, 8, n, 3).getDisplayValues();     // H S.No, I Name, J EPIC
    const v  = sh.getRange(2, C.R, n, 9).getDisplayValues();
    for(let i = n - 1; i >= 0; i--){
      const k = str_(hij[i][2]); if(!k || seen[k]) continue; seen[k] = 1;
      const a = fromRow_(v[i]); if(a[0] !== 'Residing') continue;
      const p = str_(ps[i][0]); (data[p] || (data[p] = [])).push([hij[i][0], hij[i][1], k, a.slice(0, 8)]);
    }
  }
  return { ok:true, resall:true, data:data };
}

/* ------------------------------------------------------------------ POST (save) */
// Body: { ps, blo, partName, v:2, rows: [[sn, epic, name, relType, relName, oldPart, oldSerial, age, gender, asddStatus, ...8 answer values]] }
// Rows from the old website (13 values) are refused with a 'please refresh' message.
function doPost(e){
  const lock = LockService.getScriptLock();
  try{
    if(!lock.tryLock(25000)) return out_({ ok:false, busy:true, error:'Server busy, try again' });
    const body = JSON.parse(e.postData.contents || '{}');
    if(body.action === 'freeze') return out_({ ok:false, error:'Freeze is disabled' });
    return out_(save_(body));
  }catch(err){
    return out_({ ok:false, error:String(err && err.message || err) });
  }finally{
    try{ lock.releaseLock(); }catch(e){}
  }
}

function save_(body){
  const ps = str_(body.ps), blo = str_(body.blo), part = str_(body.partName), list = body.rows || [];
  if(!ps || !list.length) return { ok:false, error:'nothing to save' };
  // An old copy of the website (before 02-Nov) is still open on some phone – make it refresh instead of wiping details
  if(body.v !== 2 || list.some(function(r){ return r.length < 18; })) return { ok:false, error:'Website updated – please refresh the page and submit again' };
  const sh = sh_(), n = lastRow_(sh) - 1, now = new Date();
  const epCol = n >= 1 ? sh.getRange(2, C.epic, n, 1).getDisplayValues() : [];
  const where = {};                                     // EPIC -> [sheet row numbers]
  for(let i = 0; i < epCol.length; i++){ const k = str_(epCol[i][0]); if(k) (where[k] || (where[k] = [])).push(i + 2); }

  const upd = {}, add = [], logs = [], skipped = [];
  list.forEach(function(r){
    const epic = str_(r[1]); if(!epic){ return; }
    const a = norm_(r.slice(10, 18));
    if(!a[0]){ skipped.push(epic); return; }
    const fixed = [now, DIST, AC_NO, AC_NAME, ps, part, blo, r[0], r[2], epic, r[3], r[4], r[5], r[6], r[7], r[8], r[9]]; // A..Q
    const ansCells = { R:a[0], S:a[1], U:a[5], V:a[2], W:a[3], X:a[4], Y:a[6], Z:a[7] };
    if(where[epic]) where[epic].forEach(function(row){ upd[row] = { fixed:fixed, ans:ansCells, epic:epic, name:r[2] }; });
    else add.push({ fixed:fixed, ans:ansCells, epic:epic, name:r[2] });
  });

  // Updates: write contiguous runs of rows in one go (column T – old remarks – is preserved)
  const rowsUpd = Object.keys(upd).map(Number).sort(function(x, y){ return x - y; });
  let i = 0;
  while(i < rowsUpd.length){
    let j = i; while(j + 1 < rowsUpd.length && rowsUpd[j + 1] === rowsUpd[j] + 1) j++;
    const start = rowsUpd[i], len = j - i + 1, rng = sh.getRange(start, 1, len, NCOL);
    const vals = rng.getValues(), disp = sh.getRange(start, C.R, len, 9).getDisplayValues();
    for(let k = 0; k < len; k++){
      const u = upd[start + k], oldA = fromRow_(disp[k]);
      const row = u.fixed.concat(vals[k].slice(17));      // keep R..Z, then overwrite answers
      row[C.R-1] = u.ans.R; row[C.S-1] = u.ans.S; row[C.U-1] = u.ans.U; row[C.V-1] = u.ans.V;
      row[C.W-1] = u.ans.W; row[C.X-1] = u.ans.X; row[C.Y-1] = u.ans.Y; row[C.Z-1] = u.ans.Z;
      vals[k] = row;
      const newA = [u.ans.R, u.ans.S, u.ans.V, u.ans.W, u.ans.X, u.ans.U, u.ans.Y, u.ans.Z];
      if(oldA[0] && oldA.slice(0, 8).join('|') !== newA.join('|'))
        logs.push([now, ps, blo, u.epic, u.name, oldA[0], oldA[1], details_(oldA), newA[0], newA[1], details_(newA)]);
    }
    sh.getRange(start, C.U, len, 6).setNumberFormat('@');     // keep dates / mobile numbers as text
    rng.setValues(vals);
    i = j + 1;
  }
  // New electors: append at the bottom
  if(add.length){
    const start = lastRow_(sh) + 1;
    const vals = add.map(function(u){
      const row = u.fixed.slice(); while(row.length < NCOL) row.push('');
      row[C.R-1] = u.ans.R; row[C.S-1] = u.ans.S; row[C.U-1] = u.ans.U; row[C.V-1] = u.ans.V;
      row[C.W-1] = u.ans.W; row[C.X-1] = u.ans.X; row[C.Y-1] = u.ans.Y; row[C.Z-1] = u.ans.Z;
      return row;
    });
    if(start + vals.length - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), start + vals.length - 1 - sh.getMaxRows());
    sh.getRange(start, C.U, vals.length, 6).setNumberFormat('@');
    sh.getRange(start, 1, vals.length, NCOL).setValues(vals);
  }
  if(logs.length){
    const lg = log_();
    lg.getRange(lg.getLastRow() + 1, 1, logs.length, logs[0].length).setValues(logs);
  }
  SpreadsheetApp.flush();
  try{ CacheService.getScriptCache().remove('summary_v2'); }catch(e){}
  return { ok:true, saved: list.length - skipped.length, skipped: skipped };
}
function details_(a){
  return ['Enum dates: ' + (a[2]||'-'), 'Re-verif dates: ' + (a[3]||'-'), 'Reason EF: ' + (a[4]||'-'),
          'Reason no F6: ' + (a[5]||'-'), 'Refused by: ' + (a[6]||'-') + ' ' + (a[7]||'')].join(' | ');
}

/* ------------------------------------------------------------------ one-time setup (run once from the editor) */
// Creates the two tabs, then fills ASDD_Reverification with all 53,859 electors of AC-39 (columns A..Q, answers empty),
// reading the lists from the website. Safe to run again: electors already in the sheet are skipped.
function prefillFromSite(){
  const sh = sh_(); log_();
  const n = lastRow_(sh) - 1, have = {};
  if(n >= 1) sh.getRange(2, C.epic, n, 1).getDisplayValues().forEach(function(r){ if(str_(r[0])) have[str_(r[0])] = 1; });
  const blos = JSON.parse(UrlFetchApp.fetch(SITE + '/data/blos.json').getContentText());
  const byPs = {}; blos.forEach(function(b){ byPs[String(b.ps)] = b; });
  const rows = [];
  for(let g = 1; g <= 8; g++){
    const grp = JSON.parse(UrlFetchApp.fetch(SITE + '/data/g' + g + '.json').getContentText());
    Object.keys(grp).sort(function(a, b){ return a - b; }).forEach(function(ps){
      const b = byPs[ps] || {};
      grp[ps].forEach(function(r){
        if(have[r[1]]) return;
        const row = ['', DIST, AC_NO, AC_NAME, ps, b.part || '', b.blo || '', r[0], r[2], r[1]].concat(r.slice(3, 10));
        while(row.length < NCOL) row.push('');
        rows.push(row);
      });
    });
  }
  if(!rows.length){ Logger.log('Nothing to add – all electors are already in the sheet.'); return; }
  const start = lastRow_(sh) + 1, need = start + rows.length - 1;
  if(need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  if(sh.getMaxColumns() > NCOL) sh.deleteColumns(NCOL + 1, sh.getMaxColumns() - NCOL);
  for(let i = 0; i < rows.length; i += 5000)
    sh.getRange(start + i, 1, Math.min(5000, rows.length - i), NCOL).setValues(rows.slice(i, i + 5000));
  SpreadsheetApp.flush();
  try{ CacheService.getScriptCache().remove('summary_v2'); }catch(e){}
  Logger.log('Added ' + rows.length + ' electors. Total rows now: ' + (lastRow_(sh) - 1));
}

// One-time repair (06-Nov-2026): the first prefill wrote EPIC into column I and Name into column J.
// Swaps them back for rows where column I looks like an EPIC number. Safe to run again.
function fixNameEpicColumns(){
  const sh = sh_(), n = lastRow_(sh) - 1; if(n < 1) return;
  const rng = sh.getRange(2, 9, n, 2), v = rng.getValues(); let fixed = 0;
  for(let i = 0; i < n; i++){
    if(/^[A-Z]{2,4}\d{5,8}$/.test(str_(v[i][0])) && !/^[A-Z]{2,4}\d{5,8}$/.test(str_(v[i][1]))){ const t = v[i][0]; v[i][0] = v[i][1]; v[i][1] = t; fixed++; }
  }
  rng.setValues(v); SpreadsheetApp.flush();
  try{ CacheService.getScriptCache().remove('summary_v2'); }catch(e){}
  Logger.log('Swapped Name/EPIC in ' + fixed + ' rows');
}
