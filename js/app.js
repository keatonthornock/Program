// app.js
// Simple app: loads config.json, fetches Admin CSV from Google Sheets, parses Key/Value pairs,
// and renders the basic fields to the page.
//
// IMPORTANT: For the CSV export to work, your Google sheet must be shared publicly (Anyone with link -> Viewer).
// See README for instructions.

const $ = sel => document.querySelector(sel);
const cfgPath = '/config.json';

async function loadConfig(){
  try {
    const r = await fetch(cfgPath);
    if(!r.ok) throw new Error('Missing config.json or not reachable');
    return await r.json();
  } catch(err){
    showError(`Cannot load config.json. Create config.json from config.json.example and fill sheetId & adminGid.\n${err.message}`);
    throw err;
  }
}

function buildCsvUrl(sheetId, gid){
  // CSV export: ensure sheet is shared publicly
  return `https://docs.google.com/spreadsheets/d/e/2PACX-1vRc9q9Jjs8H2JPQGZMDAAluqiJ2IOjjs2stLdZghdmA2qGS2sWnce9fyuTAeejWHeF4sU9rB4pzFNmG/pubhtml`;
}

function parseCSV(text){
  // Basic CSV parser: returns array of rows (array of strings). Handles quoted fields and CRLF.
  const rows = [];
  let cur = [''];
  let inQuotes = false;
  for(let i=0;i<text.length;i++){
    const ch = text[i];
    if(ch === '"' ){
      if(inQuotes && text[i+1] === '"'){ // escaped quote
        cur[cur.length-1] += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }
    if(!inQuotes && (ch === ',')){
      cur.push('');
      continue;
    }
    if(!inQuotes && (ch === '\n' || ch === '\r')){
      // handle CRLF
      if(ch === '\r' && text[i+1] === '\n'){ /* skip, loop will skip next\n */ }
      rows.push(cur);
      cur = [''];
      // Skip possible following \n in CRLF
      if(ch === '\r' && text[i+1] === '\n') i++;
      continue;
    }
    cur[cur.length-1] += ch;
  }
  // finish last
  if(cur.length && (cur.some(c => c !== ''))) rows.push(cur);
  return rows;
}

function normalizeKey(k){
  return String(k||'').trim().toLowerCase().replace(/\s+/g,' ');
}

function formatDateIfPossible(s){
  if(!s) return '';
  // try parse known formats; if parse fails, return original
  const d = new Date(s);
  if(!isNaN(d)) {
    return d.toLocaleDateString(undefined, { weekday:'long', month:'long', day:'numeric', year:'numeric' });
  }
  // else try mm/dd/yyyy manually:
  const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if(m){
    const dd = new Date(`${m[3]}-${m[1].padStart(2,'0')}-${m[2].padStart(2,'0')}`);
    if(!isNaN(dd)) return dd.toLocaleDateString(undefined, { month:'long', day:'numeric', year:'numeric' });
  }
  return s;
}

function showError(msg){
  const n = $('#notice');
  n.hidden = false;
  n.textContent = msg;
}

function clearError(){ $('#notice').hidden = true; $('#notice').textContent = ''; }

async function run(){
  clearError();
  let config;
  try { config = await loadConfig(); } catch(e){ return; }
  const sheetId = config.sheet_id;
  const adminGid = config.admin_gid;
  if(!sheetId || !adminGid){
    showError('config.json is missing sheet_id or admin_gid. Open config.json.example for instructions.');
    return;
  }

  const csvUrl = buildCsvUrl(sheetId, adminGid);
  $('#sheet-link').href = `https://docs.google.com/spreadsheets/d/${sheetId}/edit#gid=${adminGid}`;

  try {
    const resp = await fetch(csvUrl);
    if(!resp.ok) throw new Error(`Sheet fetch failed (${resp.status}) — check sharing settings.`);
    const text = await resp.text();
    const rows = parseCSV(text);
    if(!rows || rows.length < 2) {
      showError('CSV returned but no rows found. Ensure the Administrative sheet has header row "Key,Value" and data rows.');
      return;
    }
    // Build key/value map
    const headers = rows[0].map(h=> String(h||'').trim().toLowerCase());
    // Prefer columns 'key' and 'value' in any order
    const keyCol = headers.indexOf('key') >= 0 ? headers.indexOf('key') : 0;
    const valCol = headers.indexOf('value') >= 0 ? headers.indexOf('value') : 1;
    const map = {};
    for(let i=1;i<rows.length;i++){
      const r = rows[i];
      const k = r[keyCol] ? r[keyCol].trim() : '';
      const v = (r[valCol] !== undefined) ? r[valCol].trim() : '';
      if(k) map[normalizeKey(k)] = v;
    }

    // Render into page
    const title = map['title'] || map['church title'] || map['the church of jesus christ of latter-day saints'] || '';
    const ward  = map['ward'] || '';
    const stake = map['stake'] || '';
    const dateRaw = map['upcoming sunday date'] || map['upcoming sunday'] ||  '';
    const presiding = map['presiding'] || '';
    const conducting = map['conducting'] || '';
    const meetingType = map['meeting type'] || map['meeting type (is_meeting)'] || '';

    // Header/hero
    $('#title').textContent = title || 'The Church of Jesus Christ of Latter-day Saints';
    $('#submeta').textContent = [ward, stake, dateRaw ? formatDateIfPossible(dateRaw) : ''].filter(Boolean).join(' · ');

    // Program meta
    $('#presiding').textContent = presiding || '—';
    $('#conducting').textContent = conducting || '—';
    $('#meeting-type').querySelector('span').textContent = meetingType || '—';

    // Admin value table
    $('#val-title').textContent = title || '—';
    $('#val-ward').textContent = ward || '—';
    $('#val-stake').textContent = stake || '—';
    $('#val-date').textContent = dateRaw ? formatDateIfPossible(dateRaw) : '—';
    $('#val-presiding').textContent = presiding || '—';
    $('#val-conducting').textContent = conducting || '—';
    $('#val-meetingtype').textContent = meetingType || '—';

    clearError();
  } catch(err){
    console.error(err);
    showError('Failed to fetch the sheet CSV. Two common fixes:\n1) Make the sheet public (Share → Anyone with the link → Viewer), then try again.\n2) Or use "Publish to web" and use the published CSV link.\n\nError: ' + err.message);
  }
}

run();
