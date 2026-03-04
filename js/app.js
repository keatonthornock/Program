// app.js (diagnostic build)
// Loads config.json, fetches Admin CSV from Google Sheets (or admin_csv_url),
// parses Key/Value pairs, and renders the basic fields to the page.

// Utility
const $ = sel => document.querySelector(sel);
const cfgPath = window.location.pathname.includes('/Program/')
  ? '/Program/config.json'
  : './config.json';

async function loadConfig(){
  try {
    const r = await fetch(cfgPath);
    if(!r.ok) throw new Error('Missing config.json or not reachable');
    const cfg = await r.json();
    console.log('[app] loaded config.json:', cfg);
    return cfg;
  } catch(err){
    showError(`Cannot load config.json. Create config.json from config.json.example and fill sheet_id & admin_gid.\n${err.message}`);
    throw err;
  }
}

function buildCsvUrl(sheetId, gid){
  return `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}`;
}

function stripBOM(s){
  if(!s) return s;
  return s.charCodeAt(0) === 0xFEFF ? s.slice(1) : s;
}

function parseCSV(text){
  // remove BOM
  text = stripBOM(text);
  const rows = [];
  let cur = [''];
  let inQuotes = false;
  for(let i=0;i<text.length;i++){
    const ch = text[i];
    if(ch === '"'){
      if(inQuotes && text[i+1] === '"'){ // escaped quote
        cur[cur.length-1] += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }
    if(!inQuotes && ch === ','){
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

  // Trim empty trailing rows
  while(rows.length && rows[rows.length-1].every(c => c === '')) rows.pop();
  return rows;
}

function normalizeKey(k){
  return String(k||'').trim().toLowerCase().replace(/\s+/g,' ');
}

function formatDateIfPossible(s){
  if(!s) return '';
  const d = new Date(s);
  if(!isNaN(d)) {
    return d.toLocaleDateString(undefined, { weekday:'long', month:'long', day:'numeric', year:'numeric' });
  }
  // mm/dd/yyyy fallback
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
  console.warn('[app] user error:', msg);
}

function clearError(){ $('#notice').hidden = true; $('#notice').textContent = ''; }

// run
async function run(){
  clearError();
  let config;
  try { config = await loadConfig(); } catch(e){ return; }

  // Build CSV URL - support direct admin_csv_url fallback
  let csvUrl = null;
  if(config.admin_csv_url){
    csvUrl = config.admin_csv_url;
    console.log('[app] using admin_csv_url from config:', csvUrl);
  } else {
    const sheetId = config.sheet_id;
    const adminGid = config.admin_gid;
    if(!sheetId || !adminGid){
      showError('config.json is missing sheet_id or admin_gid. Open config.json.example for instructions.');
      return;
    }
    csvUrl = buildCsvUrl(sheetId, adminGid);
    const sheetLink = $('#sheet-link');
    if (sheetLink) sheetLink.href = `https://docs.google.com/spreadsheets/d/${sheetId}/edit#gid=${adminGid}`;
    console.log('[app] built csvUrl from sheet_id/admin_gid:', csvUrl);
  }

  try {
    const resp = await fetch(csvUrl, { cache: "no-store" });
    console.log('[app] fetch response:', resp.status, resp.headers.get('content-type'));
    const text = await resp.text();

    // Quick heuristic: if the response contains HTML tags, it may be a redirect/login page
    const first500 = text.slice(0,500).replace(/\s+/g,' ');
    console.log('[app] response first chars:', first500);

    if(/<html|doctype html|<script|login/i.test(first500)){
      // Likely an HTML page, not CSV (login, permission or publish page)
      showError('The sheet returned HTML (likely a permission/login page). Ensure the sheet is published or set to "Anyone with the link - Viewer".\nIf you used "Publish to web", use the published CSV link in config as admin_csv_url.');
      return;
    }

    if(!resp.ok) throw new Error(`Sheet fetch failed (${resp.status}) — check sharing settings.`);

    const rows = parseCSV(text);
    console.log('[app] parsed rows:', rows.length, rows.slice(0,6));

    if(!rows || rows.length < 2) {
      showError('CSV returned but no rows found. Ensure the Administrative sheet has header row "Key,Value" and data rows.');
      return;
    }

    // Build key/value map
    const headers = rows[0].map(h=> String(h||'').trim().toLowerCase());
    const keyCol = headers.indexOf('key') >= 0 ? headers.indexOf('key') : 0;
    const valCol = headers.indexOf('value') >= 0 ? headers.indexOf('value') : 1;
    const map = {};
    for(let i=1;i<rows.length;i++){
      const r = rows[i];
      const k = r[keyCol] ? r[keyCol].trim() : '';
      const v = (r[valCol] !== undefined) ? r[valCol].trim() : '';
      if(k) map[normalizeKey(k)] = v;
    }
    console.log('[app] map keys:', Object.keys(map));

    // Render into page
    const title = map['title'] || map['church title'] || map['the church of jesus christ of latter-day saints'] || '';
    const ward  = map['ward'] || '';
    const stake = map['stake'] || '';
    const dateRaw = map['upcoming sunday date'] || map['upcoming sunday'] ||  '';
    const presiding = map['presiding'] || '';
    const conducting = map['conducting'] || '';
    const meetingType = map['meeting type'] || map['meeting type (is_meeting)'] || '';

    $('#title').textContent = title || 'The Church of Jesus Christ of Latter-day Saints';
    $('#submeta').textContent = [ward, stake, dateRaw ? formatDateIfPossible(dateRaw) : ''].filter(Boolean).join(' · ');

    $('#presiding').textContent = presiding || '—';
    $('#conducting').textContent = conducting || '—';
    const mtSpan = $('#meeting-type').querySelector('span');
    if(mtSpan) mtSpan.textContent = meetingType || '—';

    $('#val-title').textContent = title || '—';
    $('#val-ward').textContent = ward || '—';
    $('#val-stake').textContent = stake || '—';
    $('#val-date').textContent = dateRaw ? formatDateIfPossible(dateRaw) : '—';
    $('#val-presiding').textContent = presiding || '—';
    $('#val-conducting').textContent = conducting || '—';
    $('#val-meetingtype').textContent = meetingType || '—';

    clearError();
  } catch(err){
    console.error('[app] error fetching/parsing sheet:', err);
    showError('Failed to fetch the sheet CSV. Two common fixes:\n1) Make the sheet public (Share → Anyone with the link → Viewer), then try again.\n2) Or use "Publish to web" and use the published CSV link in config as admin_csv_url.\n\nError: ' + err.message);
  }
}

run();
