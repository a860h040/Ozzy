const SHEET_NAME = 'ICU Workups';
const FIELDS = [
 ['recordId','Record ID'],['updatedAt','Updated At'],['patientCode','Room Number'],
 ['cc','CC'],['hpi','HPI'],['pmh','PMH'],['whyIcu','Why ICU'],
 ['cpot','CPOT'],['rass','RASS'],['camIcu','CAM-ICU'],['sedationMeds','Sedation Meds'],['seizureMeds','Anti-seizure Meds'],['painMeds','Pain Meds'],
 ['respSupport','RA / NC / Vent'],['fio2','FiO2'],['peep','PEEP'],['ph','pH'],['paco2','PaCO2'],['pao2','PaO2'],['hco3','HCO3'],['respMeds','Respiratory Meds'],['abgInterpretation','ABG Interpretation'],
 ['sbp','SBP'],['map','MAP'],['mapTrends','MAP Trends'],['hr','HR'],['echo','Echo'],['ef','EF'],['bpMeds','BP Meds'],['pressors','Pressors'],
 ['bg','BG'],['insulin','Insulin'],['hgb','Hgb'],['hct','Hct'],['platelets','Platelets'],['scr','SCr'],['crcl','CrCl (mL/min)'],['rrt','RRT'],['hd','HD'],['crrt','CRRT'],['fluid24','Fluid Balance 24h'],['fluidTotal','Fluid Balance Total'],['acMeds','AC Meds'],['diuretics','Diuretics'],
 ['dvt','DVT PPx'],['sup','SUP'],['wbc','WBC'],['mrsa','MRSA PCR'],['fever','Febrile/Afebrile'],['cultures','Cultures'],['antibiotics','Antibiotics'],['therapyDay','Day of Therapy'],['endDate','End Date'],['hdReceived','HD Received'],['hdSchedule','HD Schedule']
];
function doGet() { sheet_(); return HtmlService.createHtmlOutputFromFile('Index').setTitle('ICU Workup'); }
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Bind this script to an approved Google Sheet.');
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  // Migrate older worksheets by removing the retired Workup Date column.
  if (sh.getLastRow() > 0 && sh.getLastColumn() >= 4 &&
      sh.getRange(1, 4).getDisplayValue() === 'Workup Date') {
    sh.deleteColumn(4);
  }
  const headers = FIELDS.map(f => f[1]);
  if (sh.getLastRow() > 0) {
    const oldHeaders = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
    if (!oldHeaders.includes('CrCl (mL/min)') && oldHeaders.includes('SCr')) {
      const scrColumn = oldHeaders.indexOf('SCr') + 1;
      sh.insertColumnAfter(scrColumn);
      sh.getRange(1,scrColumn+1).setValue('CrCl (mL/min)');
    }
  }
  if (sh.getLastRow() === 0) { sh.appendRow(headers); sh.setFrozenRows(1); }
  else {
    const existing = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
    const prefixMatches = existing.every((heading,i)=>heading===headers[i]);
    if (!prefixMatches) throw new Error('Sheet columns changed. Restore original headers.');
    if (existing.length < headers.length) sh.getRange(1,existing.length+1,1,headers.length-existing.length).setValues([headers.slice(existing.length)]);
  }
  return sh;
}
function clean_(v) { const s = String(v ?? '').slice(0,10000); return /^[=+@\-\t\r]/.test(s) ? "'"+s : s; }
function saveWorkup(data) {
  if (!data) throw new Error('Workup data is required.');
  const sh=sheet_(); const id=String(data.recordId||Utilities.getUuid());
  const record={...data,recordId:id,updatedAt:new Date().toISOString()};
  const row=FIELDS.map(([key])=>clean_(record[key]));
  const last=sh.getLastRow(); let target=0;
  if(last>1){const ids=sh.getRange(2,1,last-1,1).getValues().flat();const idx=ids.indexOf(id);if(idx>=0)target=idx+2;}
  if(target)sh.getRange(target,1,1,row.length).setValues([row]); else sh.appendRow(row);
  return {recordId:id,updatedAt:record.updatedAt};
}
function listWorkups(){
  const sh=sheet_(),last=sh.getLastRow();if(last<2)return [];
  const rows=sh.getRange(2,1,last-1,FIELDS.length).getDisplayValues();
  return rows.map(r=>Object.fromEntries(FIELDS.map(([k],i)=>[k,r[i]]))).map(r=>({recordId:r.recordId,patientCode:r.patientCode,updatedAt:r.updatedAt})).reverse();
}
function getWorkup(id){
  const sh=sheet_(),last=sh.getLastRow();if(last<2)throw new Error('No records found.');
  const rows=sh.getRange(2,1,last-1,FIELDS.length).getDisplayValues();
  const r=rows.find(row=>row[0]===String(id));if(!r)throw new Error('Record not found.');
  return Object.fromEntries(FIELDS.map(([k],i)=>[k,r[i]]));
}

// Free NLM RxTerms medication name autocomplete; never sends patient data.
function searchMedications(query) {
  const q = String(query || '').trim().slice(0, 80);
  if (q.length < 3) return [];
  const cache = CacheService.getScriptCache();
  const key = 'rxterms_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, q.toLowerCase())).slice(0,35);
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);
  const url = 'https://clinicaltables.nlm.nih.gov/api/rxterms/v3/search?terms=' + encodeURIComponent(q) + '&maxList=12';
  const response = UrlFetchApp.fetch(url, {muteHttpExceptions:true,followRedirects:true});
  if (response.getResponseCode() !== 200) throw new Error('Medication search is temporarily unavailable.');
  const data = JSON.parse(response.getContentText());
  const names = Array.isArray(data[3]) ? [...new Set(data[3].map(row => Array.isArray(row) ? String(row[0]) : String(row)).filter(Boolean))].slice(0,12) : [];
  cache.put(key,JSON.stringify(names),3600);
  return names;
}
/**
 * Receives a form POST from the GitHub Pages version.
 * Requires a newly deployed Apps Script web-app version.
 * Never include patient data in the acknowledgement.
 */
function doPost(e) {
  let reply;
  try {
    if (!e || !e.parameter || e.parameter.action !== 'saveWorkup') throw new Error('Invalid request');
    const data = JSON.parse(e.parameter.payload || '{}');
    if (!data || typeof data !== 'object' || Array.isArray(data)) throw new Error('Invalid workup');
    const saved = saveWorkup(data);
    reply = {type:'icu-workup-save-result', ok:true, recordId:saved.recordId};
  } catch (err) {
    reply = {type:'icu-workup-save-result', ok:false, error:String(err.message || err).slice(0,200)};
  }
  const html = '<!doctype html><html><body><script>parent.postMessage(' +
    JSON.stringify(reply).replace(/</g, '\\u003c') +
    ', "*");</scr' + 'ipt></body></html>';
  return HtmlService.createHtmlOutput(html);
}
