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
function doGet() {
  const html = `<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><title>Saved ICU Workups</title>
  <style>body{font:14px Arial,sans-serif;color:#203244;margin:20px;background:#f5f6f7}main{max-width:950px;margin:auto;background:white;padding:20px;border:1px solid #ddd}h1{font-size:23px}button{background:#263f52;color:white;border:0;border-radius:4px;padding:9px 12px;margin:4px;cursor:pointer}input{padding:9px;border:1px solid #bbb}table{width:100%;border-collapse:collapse}td,th{padding:10px;border-bottom:1px solid #ddd;text-align:left;vertical-align:top}pre{white-space:pre-wrap;overflow-wrap:anywhere}#details{margin-top:20px} @media print{button,input,#list{display:none}body{background:white}main{border:0}}</style></head><body><main><h1>Saved ICU Workups</h1><p>Access is controlled by your Google Apps Script deployment permissions.</p><button onclick="refresh()">Refresh from Google Sheets</button><input id="q" placeholder="Search room number" oninput="render()"><button onclick="download()">Download CSV</button><span id="status"></span><div id="list"></div><div id="details"></div></main>
  <script>
  let records=[],selected=null;
  const esc=x=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',[String.fromCharCode(39)]:'&#39;'}[c]));
  const status=x=>document.getElementById('status').textContent=x;
  function refresh(){status('Loading...');google.script.run.withSuccessHandler(rows=>{records=rows;render();status(rows.length+' records');}).withFailureHandler(e=>status(e.message)).listWorkups();}
  function render(){const q=document.getElementById('q').value.toLowerCase();const rows=records.filter(r=>String(r.patientCode||'').toLowerCase().includes(q));document.getElementById('list').innerHTML='<table><thead><tr><th>Room</th><th>Last updated</th><th>View</th></tr></thead><tbody>'+rows.map(r=>'<tr><td>'+esc(r.patientCode)+'</td><td>'+esc(r.updatedAt)+'</td><td><button data-id="'+esc(r.recordId)+'" onclick="openRecord(this.dataset.id)">View</button></td></tr>').join('')+'</tbody></table>';}
  function showRecord(){
    if(!selected)return;
    document.getElementById('details').innerHTML='<h2>Room '+esc(selected.patientCode)+'</h2><button onclick="editRecord()">Edit / Complete Missing Information</button><button onclick="printWorksheet()">Download PDF / Print Worksheet</button><table>'+Object.entries(selected).filter(([k,v])=>v!==''&&v!=null).map(([k,v])=>'<tr><th>'+esc(k)+'</th><td>'+esc(v)+'</td></tr>').join('')+'</table>';
  }
  function openRecord(id){status('Loading room...');google.script.run.withSuccessHandler(data=>{selected=data;showRecord();status('Record loaded');}).withFailureHandler(e=>status(e.message)).getWorkup(id);}
  function editRecord(){
    if(!selected)return;
    const editable=Object.entries(selected).filter(([k])=>k!=='recordId'&&k!=='updatedAt');
    document.getElementById('details').innerHTML='<h2>Edit Room '+esc(selected.patientCode)+'</h2><p>Update missing or incorrect information, then save to the same Google Sheet record.</p><button onclick="saveEditedRecord()">Save Changes</button><button onclick="showRecord()">Cancel</button><table>'+editable.map(([k,v])=>'<tr><th><label for="edit_'+esc(k)+'">'+esc(k)+'</label></th><td><textarea id="edit_'+esc(k)+'" data-key="'+esc(k)+'" rows="2" style="width:100%;min-height:50px;padding:7px;border:1px solid #aaa">'+esc(v)+'</textarea></td></tr>').join('')+'</table><button onclick="saveEditedRecord()">Save Changes</button>';
  }
  function saveEditedRecord(){
    if(!selected)return;
    const updated={...selected};
    document.querySelectorAll('#details textarea[data-key]').forEach(el=>updated[el.dataset.key]=el.value);
    status('Saving changes...');
    google.script.run.withSuccessHandler(result=>{
      google.script.run.withSuccessHandler(data=>{selected=data;showRecord();refresh();status('Changes saved to Google Sheets.');}).withFailureHandler(e=>status('Saved, but refresh failed: '+e.message)).getWorkup(result.recordId);
    }).withFailureHandler(e=>status('Update failed: '+e.message)).updateExistingWorkup(updated);
  }
  function printWorksheet(){
    if(!selected){status('Select a room first.');return;}
    const d=selected;
    const groups=[
      {name:'Patient',rows:[['Room','patientCode'],['CC','cc'],['HPI','hpi'],['PMH','pmh'],['Why ICU?','whyIcu']]},
      {name:'Neuro',rows:[['CPOT','cpot'],['RASS','rass'],['CAM-ICU','camIcu'],['Sedation','sedationMeds'],['Anti-seizure','seizureMeds'],['Pain','painMeds']]},
      {name:'Resp',rows:[['RA / NC / Vent','respSupport'],['FiO₂','fio2'],['PEEP','peep'],['ABG pH','ph'],['PaCO₂','paco2'],['PaO₂','pao2'],['HCO₃','hco3'],['ABG interpretation','abgInterpretation'],['Resp meds','respMeds']]},
      {name:'Cards',rows:[['SBP','sbp'],['MAP','map'],['MAP trends','mapTrends'],['HR','hr'],['Echo','echo'],['EF','ef'],['BP meds','bpMeds'],['Pressors','pressors']]},
      {name:'Heme / Renal',rows:[['BG','bg'],['Insulin','insulin'],['Hgb','hgb'],['Hct','hct'],['Platelets','platelets'],['SCr','scr'],['CrCl','crcl'],['RRT','rrt'],['HD','hd'],['CRRT','crrt'],['Fluid balance 24h','fluid24'],['Fluid balance total','fluidTotal'],['AC','acMeds'],['Diuretics','diuretics'],['HD received','hdReceived'],['HD schedule','hdSchedule']]},
      {name:'PPx',rows:[['DVT','dvt'],['SUP','sup']]},
      {name:'ID',rows:[['WBC','wbc'],['MRSA PCR','mrsa'],['Febrile / Afebrile','fever'],['Cultures','cultures'],['Antibiotics','antibiotics'],['Day of therapy','therapyDay'],['End date','endDate']]}
    ];
    const cols=groups.map(g=>'<section class="system"><h2>'+esc(g.name)+'</h2><div class="system-fields">'+g.rows.map(([label,key])=>'<div class="item"><div class="label">'+esc(label)+'</div><div class="value">'+esc(d[key]||'')+'</div></div>').join('')+'</div></section>').join('');
    const w=window.open('','_blank');
    if(!w){status('Please allow popups to print.');return;}
    const css='@page{size:landscape;margin:12mm}*{box-sizing:border-box}body{font:11px Arial,sans-serif;color:#111;margin:0;padding:8px}header{display:flex;justify-content:space-between;align-items:center;border-bottom:2px solid #111;padding-bottom:9px;margin-bottom:14px}h1{font-size:19px;margin:0}h2{font-size:15px;border-bottom:2px solid #263f52;background:#edf1f4;margin:0 0 10px;padding:8px}.sheet{display:grid;grid-template-columns:1fr 1fr;gap:14px;align-items:start}.system{border:1px solid #a9b1b8;border-radius:3px;padding:0 10px 10px;break-inside:avoid;page-break-inside:avoid}.system-fields{display:grid;grid-template-columns:1fr 1fr;gap:8px 14px}.item{margin:0;break-inside:avoid}.label{font-weight:700;font-size:10px;margin-bottom:4px}.value{min-height:22px;white-space:pre-wrap;overflow-wrap:anywhere;border-bottom:1px solid #ddd;padding-bottom:5px}.controls{margin:10px 0}button{background:#263f52;color:white;border:0;padding:10px;border-radius:4px}@media print{.controls{display:none}body{padding:0}.system{break-inside:avoid}}';
    w.document.open();w.document.write('<!doctype html><html><head><meta charset="utf-8"><title>ICU Room '+esc(d.patientCode)+'</title><style>'+css+'</style></head><body><header><h1>ICU Patient Workup — Room '+esc(d.patientCode)+'</h1><span>Updated: '+esc(d.updatedAt||'')+'</span></header><div class="controls"><button onclick="window.print()">Print / Save as PDF</button></div><div class="sheet">'+cols+'</div></body></html>');w.document.close();
  }
  function download(){status('Preparing export...');google.script.run.withSuccessHandler(rows=>{const csv=rows.map(r=>r.map(x=>'"'+String(x??'').replace(/"/g,'""')+'"').join(',')).join('\\r\\n');const blob=new Blob(['\\uFEFF'+csv],{type:'text/csv;charset=utf-8'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download='icu-workups.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),3000);status('CSV downloaded');}).withFailureHandler(e=>status(e.message)).exportWorkups();}
  refresh();
  </scr`+`ipt></body></html>`;
  return HtmlService.createHtmlOutput(html).setTitle('Saved ICU Workups').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function exportWorkups() {
  const sh=sheet_();
  return sh.getDataRange().getDisplayValues();
}
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
// Update only an existing saved record. Never append a new row during editing.
function updateExistingWorkup(data) {
  if (!data || !data.recordId) throw new Error('Select a saved workup before editing.');
  const sh=sheet_(), id=String(data.recordId), last=sh.getLastRow();
  if(last<2) throw new Error('No saved workups found.');
  const ids=sh.getRange(2,1,last-1,1).getDisplayValues().flat();
  const idx=ids.indexOf(id);
  if(idx<0) throw new Error('This record no longer exists. Refresh the saved workups.');
  const current=sh.getRange(idx+2,1,1,FIELDS.length).getDisplayValues()[0];
  const previous=Object.fromEntries(FIELDS.map(([key],i)=>[key,current[i]]));
  if(data.updatedAt && previous.updatedAt !== String(data.updatedAt)) throw new Error('This workup was updated by another user. Refresh before editing.');
  const record={...previous,...data,recordId:id,updatedAt:new Date().toISOString()};
  sh.getRange(idx+2,1,1,FIELDS.length).setValues([FIELDS.map(([key])=>clean_(record[key]))]);
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
