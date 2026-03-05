// app.js - Render new mobile agenda layout using Administrative sheet (CSV).
const $ = sel => document.querySelector(sel);
const cfgPath = './config.json';

async function loadConfig(){
  try {
    const r = await fetch(cfgPath);
    if(!r.ok) throw new Error('Missing config.json or not reachable');
    return await r.json();
  } catch(err){
    showError(`Cannot load config.json. ${err.message}`);
    throw err;
  }
}

function buildCsvUrl(sheetId, gid){
  return `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}`;
}

function parseCSV(text){
  // Simple CSV parse, robust enough for key/value Admin sheet
  text = text.replace(/^\uFEFF/, '');
  const lines = text.split(/\r?\n/).map(l => l.trim());
  return lines.map(l => {
    // split on first comma only (Key,Value)
    if(!l) return [];
    const idx = l.indexOf(',');
    if(idx === -1) return [l];
    return [ l.slice(0,idx).trim(), l.slice(idx+1).trim().replace(/^"(.*)"$/, '$1') ];
  });
}

function showError(msg){
  const n = $('#notice');
  if(n){
    n.hidden = false;
    n.textContent = msg;
  } else {
    console.warn(msg);
  }
}
function clearError(){ const n = $('#notice'); if(n){ n.hidden = true; n.textContent = ''; } }

function formatDateIfPossible(s){
  if(!s) return '';
  const d = new Date(s);
  if(!isNaN(d)) {
    return d.toLocaleDateString(undefined, { weekday:'long', month:'long', day:'numeric', year:'numeric' });
  }
  return s;
}

function createHymnCard(title, hymnNumber, label='Opening Hymn'){
  const el = document.createElement('div');
  el.className = 'hymn-card';
  el.innerHTML = `
    <div class="left">
      <div class="hymn-icon" aria-hidden>
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none"><path d="M9 17V7l10-2v10" stroke="#0b4a6a" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
      </div>
      <div>
        <div class="hymn-title">${label}</div>
        <div class="hymn-sub">${hymnNumber ? `#${hymnNumber}` : ''} ${title ? ` — ${title}` : ''}</div>
      </div>
    </div>
    <div class="right">›</div>
  `;
  return el;
}

function createSimpleRow(type, text, extra=''){
  const el = document.createElement('div');
  el.className = 'agenda-item';
  el.innerHTML = `
    <div class="icon" aria-hidden> 
      <svg width="20" height="20" viewBox="0 0 24 24" fill="none">
        <circle cx="12" cy="12" r="9" stroke="#0b4a6a" stroke-width="1.2"></circle>
      </svg>
    </div>
    <div class="content">
      <div class="title">${type}</div>
      <div class="sub">${text || ''}</div>
    </div>
    <div class="right">${extra}</div>
  `;
  return el;
}

function renderAdministrative(map){
  // Header
  const title = map['title'] || 'The Church of Jesus Christ of Latter-day Saints';
  const ward = map['ward'] || '';
  const stake = map['stake'] || '';
  const dateRaw = map['upcoming sunday date'] || map['upcoming sunday'] || '';
  const presiding = map['presiding'] || '—';
  const conducting = map['conducting'] || '—';
  const meetingType = map['meeting type'] || '';

  $('#meeting-heading').textContent = meetingType || 'Sacrament Meeting';
  $('#meeting-date').textContent = (dateRaw ? formatDateIfPossible(dateRaw) : '') + (ward ? `\n${ward} · ${stake}` : (ward || stake ? ` · ${ward} ${stake}` : ''));
  $('#presiding').textContent = presiding;
  $('#conducting').textContent = conducting;

  // Program skeleton (until Agenda tab is added)
  const pc = $('#program-content');
  pc.innerHTML = ''; // clear

  // Opening hymn card from admin if they put Title fields (we check keys)
  // Try to find Opening Hymn info in admin sheet by keys: "opening hymn" / "opening hymn #"
  const openingTitle = map['opening hymn'] || '';
  const openingNum = map['opening hymn #'] || map['opening hymn number'] || '';
  if(openingTitle || openingNum) {
    pc.appendChild(createHymnCard(openingTitle, openingNum, 'Opening Hymn'));
  } else {
    // placeholder opening hymn
    pc.appendChild(createHymnCard('', '', 'Opening Hymn'));
  }

  // Invocation row
  const invocation = map['invocation'] || '';
  if(invocation) pc.appendChild(createSimpleRow('Invocation', invocation));
  else pc.appendChild(createSimpleRow('Invocation', ''));

  // Speakers placeholder section (we'll show sample if none)
  const spHeader = document.createElement('div');
  spHeader.className = 'muted small';
  spHeader.style.margin = '6px 4px';
  spHeader.textContent = 'Speakers & Intermediate Hymn';
  pc.appendChild(spHeader);

  // If admin provided "speaker1", "speaker2", etc., render them
  let speakersFound = false;
  for(let i=1;i<=6;i++){
    const key = `speaker ${i}`;
    if(map[key]) {
      speakersFound = true;
      pc.appendChild(createSimpleRow('Speaker', map[key]));
    }
  }
  if(!speakersFound) {
    // Some default placeholders to match your mockup look (they can be removed once Agenda is implemented)
    pc.appendChild(createSimpleRow('Speaker', 'Brother James White'));
    pc.appendChild(createSimpleRow('Speaker', 'Sister Anna Brown'));
    // intermediate hymn placeholder
    pc.appendChild(createHymnCard('I Know That My Redeemer Lives', 136, 'Intermediate Hymn'));
    pc.appendChild(createSimpleRow('Speaker', 'Brother Michael Johnson'));
  }

  // Closing hymn
  const closingTitle = map['closing hymn'] || '';
  const closingNum = map['closing hymn #'] || map['closing hymn number'] || '';
  if(closingTitle || closingNum){
    pc.appendChild(createHymnCard(closingTitle, closingNum, 'Closing Hymn'));
  } else {
    pc.appendChild(createHymnCard('God Be with You Till We Meet Again', 152, 'Closing Hymn'));
  }

  // Benediction
  const bened = map['benediction'] || map['benediction by'] || map['benediction name'] || '';
  if(bened) pc.appendChild(createSimpleRow('Benediction', bened));
  else pc.appendChild(createSimpleRow('Benediction', 'Brother David Lee'));

  clearError();
}

async function run(){
  clearError();
  let config;
  try { config = await loadConfig(); } catch(e){ return; }

  // Build csv url (supports admin_csv_url in config)
  let csvUrl = null;
  if(config.admin_csv_url){
    csvUrl = config.admin_csv_url;
  } else {
    const { sheet_id: sheetId, admin_gid: adminGid } = config;
    if(!sheetId || typeof adminGid === 'undefined') {
      showError('config.json is missing sheet_id or admin_gid.');
      return;
    }
    csvUrl = buildCsvUrl(sheetId, adminGid);
  }

  try {
    const resp = await fetch(csvUrl, { cache: "no-store" });
    if(!resp.ok) throw new Error(`Sheet fetch failed (${resp.status})`);
    const text = await resp.text();
    // quick heuristic: if returned HTML then permission issue
    if(/<html|doctype html/i.test(text.slice(0,200))) {
      showError('The sheet returned an HTML response (likely not shared). Make sure it is shared as "Anyone with link — Viewer" or use Publish → CSV and set admin_csv_url in config.json.');
      return;
    }
    const rows = parseCSV(text);
    if(!rows || rows.length < 2) {
      showError('Administrative CSV parsed but no rows found. Ensure header "Key,Value" exists.');
      return;
    }
    const map = {};
    for(let i=1;i<rows.length;i++){
      const r = rows[i];
      if(!r || !r[0]) continue;
      const k = String(r[0]).trim().toLowerCase();
      const v = (r[1] !== undefined) ? String(r[1]).trim() : '';
      map[k] = v;
    }

    // render administrative header + skeleton program
    renderAdministrative(map);

    // collapsible logic
    document.querySelectorAll('.collapsible-toggle').forEach(btn => {
      btn.addEventListener('click', () => {
        const target = btn.getAttribute('data-target');
        const panel = document.getElementById(target);
        const expanded = btn.getAttribute('aria-expanded') === 'true';
        btn.setAttribute('aria-expanded', (!expanded).toString());
        if(panel){
          panel.style.display = expanded ? 'none' : 'block';
        }
      });
    });

  } catch(err){
    console.error('error fetching admin csv', err);
    showError('Failed to fetch the Administrative sheet. ' + err.message);
  }
}

run();
