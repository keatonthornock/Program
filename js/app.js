// js/app.js
// Enhanced: load Administrative + Agenda CSVs, render agenda items (hymns, speakers, etc.)
const $ = s => document.querySelector(s);
const cfgPath = './config.json';

async function loadConfig(){
  try {
    const r = await fetch(cfgPath);
    if(!r.ok) throw new Error('Missing config.json');
    const cfg = await r.json();
    console.log('[app] config', cfg);
    return cfg;
  } catch(e){
    showError('Cannot load config.json. ' + e.message);
    throw e;
  }
}

function buildCsvUrl(sheetId, gid){
  return `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}`;
}

function stripBOM(s){ return s && s.charCodeAt(0) === 0xFEFF ? s.slice(1) : s; }

function parseCSVtoRows(text){
  text = stripBOM(text || '');
  // Split into lines and use simple CSV split on first two commas (we only expect 3 columns)
  const lines = text.split(/\r?\n/);
  const rows = [];
  for(const ln of lines){
    if(!ln) continue;
    // We will split on commas but handle quoted values simply:
    // Use a basic CSV split (good enough for simple sheets).
    const parts = [];
    let cur = '', inQ = false;
    for(let i=0;i<ln.length;i++){
      const ch = ln[i];
      if(ch === '"' ) { inQ = !inQ; continue; }
      if(ch === ',' && !inQ){
        parts.push(cur); cur = '';
        continue;
      }
      cur += ch;
    }
    parts.push(cur);
    rows.push(parts.map(p => (p||'').trim()));
  }
  return rows;
}

// slug helper
function slugify(text){
  return (text || '').toString().trim()
    .toLowerCase()
    .replace(/^[0-9]+\.\s*/,'')            // strip leading "123. "
    .replace(/[’'"]/g,'')                  // remove smart quotes
    .replace(/[^a-z0-9\s-]/g,'')           // remove non-safe chars
    .replace(/\s+/g,'-')
    .replace(/-+/g,'-')
    .replace(/^-|-$/g,'');
}

// hymn URL generator
function getHymnUrl(title, hymnNumber, extraInfo){
  // prefer numeric hymnNumber if present:
  const n = Number(hymnNumber);
  if(!isNaN(n) && n > 0){
    if(n <= 341){
      return `https://www.churchofjesuschrist.org/music/library/hymns/${n}`;
    }
    if(n >= 1000){
      // hymns for home and church path uses slug based on title
      const slug = slugify(title);
      if(slug) return `https://www.churchofjesuschrist.org/study/music/hymns-for-home-and-church/${slug}?lang=eng`;
      // fallback: sometimes site uses numeric id; try search fallback
      return `https://www.churchofjesuschrist.org/study/music/hymns-for-home-and-church?lang=eng`;
    }
    // fallback
    return `https://www.churchofjesuschrist.org/music/library/hymns/${n}`;
  }

  // If no numeric hymn, try to infer from extraInfo label "Hymns for Home and Church"
  if((extraInfo||'').toLowerCase().includes('hymns for home')) {
    const slug = slugify(title);
    if(slug) return `https://www.churchofjesuschrist.org/study/music/hymns-for-home-and-church/${slug}?lang=eng`;
  }

  // last resort: try search by title
  const q = encodeURIComponent(title || '');
  if(q) return `https://www.churchofjesuschrist.org/search?q=${q}`;
  return null;
}

function createElemFromHTML(html){
  const div = document.createElement('div');
  div.innerHTML = html.trim();
  return div.firstChild;
}

function createHymnCard(title, hymnNumber, label='Opening Hymn', url=null){
  const el = document.createElement('div');
  el.className = 'hymn-card';
  el.innerHTML = `
    <div class="left">
      <div class="hymn-icon" aria-hidden>
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none"><path d="M9 17V7l10-2v10" stroke="#0b4a6a" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
      </div>
      <div>
        <div class="hymn-title">${label}</div>
        <div class="hymn-sub">${hymnNumber ? `#${hymnNumber}` : ''}${title ? (hymnNumber ? ` — ${title}` : title) : ''}</div>
      </div>
    </div>
    <div class="right">${url? '›' : ''}</div>
  `;
  if(url){
    const a = document.createElement('a');
    a.href = url;
    a.target = '_blank';
    a.rel = 'noopener';
    a.className = 'hymn-link';
    // move contents into the anchor
    a.appendChild(el);
    return a;
  }
  return el;
}

function createRow(typeLabel, name, extra){
  const el = document.createElement('div');
  el.className = 'agenda-item';
  el.innerHTML = `
    <div class="icon" aria-hidden>
      <svg width="20" height="20" viewBox="0 0 24 24" fill="none">
        <circle cx="12" cy="12" r="9" stroke="#0b4a6a" stroke-width="1.2"></circle>
      </svg>
    </div>
    <div class="content">
      <div class="title">${typeLabel}</div>
      <div class="sub">${name || ''}</div>
    </div>
    <div class="right">${extra || ''}</div>
  `;
  return el;
}

function showError(msg){
  const n = $('#notice'); if(n) { n.hidden = false; n.textContent = msg; } else console.warn(msg);
}
function clearError(){ const n = $('#notice'); if(n) { n.hidden = true; n.textContent = ''; } }

function normalizeItemKey(s){ return (s||'').toString().trim().toLowerCase(); }

async function run(){
  clearError();
  let config;
  try { config = await loadConfig(); } catch(e){ return; }

  // determine URLs
  let adminCsvUrl = config.admin_csv_url || (config.sheet_id && config.admin_gid ? buildCsvUrl(config.sheet_id, config.admin_gid) : null);
  let agendaCsvUrl = config.agenda_csv_url || (config.sheet_id && config.agenda_gid ? buildCsvUrl(config.sheet_id, config.agenda_gid) : null);

  if(!adminCsvUrl){
    showError('No admin CSV URL available. Set admin_gid or admin_csv_url in config.json');
    return;
  }
  if(!agendaCsvUrl){
    showError('No agenda CSV URL available. Set agenda_gid or agenda_csv_url in config.json');
    return;
  }

  // fetch admin and agenda in parallel
  try {
    const [admResp, agResp] = await Promise.all([ fetch(adminCsvUrl), fetch(agendaCsvUrl) ]);
    if(!admResp.ok) throw new Error('Admin sheet fetch failed: ' + admResp.status);
    if(!agResp.ok) throw new Error('Agenda sheet fetch failed: ' + agResp.status);

    const admText = await admResp.text();
    const agText = await agResp.text();

    // Quick HTML detection
    if(/<html|doctype html/i.test(admText.slice(0,200))) { showError('Admin sheet returned HTML (not public)'); return; }
    if(/<html|doctype html/i.test(agText.slice(0,200))) { showError('Agenda sheet returned HTML (not public)'); return; }

    const admRows = parseCSVtoRows(admText);
    const agRows = parseCSVtoRows(agText);

    // build admin map
    const adminMap = {};
    for(let i=0;i<admRows.length;i++){
      const r = admRows[i];
      if(!r || !r[0]) continue;
      adminMap[(r[0]||'').toString().trim().toLowerCase()] = (r[1]||'').toString().trim();
    }

    // render header
    renderHeaderFromAdmin(adminMap);

    // parse and render agenda rows in order
    const container = $('#program-content');
    container.innerHTML = '';
    let any = false;

    for (let i = 0; i < agRows.length; i++) {
      const r = agRows[i];
      if (!r || !r[0]) continue; // skip completely blank rows
    
      // Normalize each column (trim and coerce to string)
      const colA = (r[0] || '').toString().trim(); // Item
      const colB = (r[1] || '').toString().trim(); // Name / Title / Hymn
      const colC = (r[2] || '').toString().trim(); // Extra Info
    
      // Detect and skip header row (tolerant to capitalization and small variations)
      const aKey = colA.toLowerCase();
      const bKey = colB.toLowerCase();
      const cKey = colC.toLowerCase();
      const looksLikeHeader =
        aKey === 'item' ||
        bKey === 'name' ||
        cKey === 'extra info' ||
        (aKey.includes('item') && bKey.includes('name')); // extra safety
      if (looksLikeHeader) {
        console.log(`[app] skipping header-like row ${i+1}: ${colA} | ${colB} | ${colC}`);
        continue;
      }
    
      // Now assign the meaningful variables the rest of the code expects
      const item = colA;   // A: Item (Opening Hymn, Invocation, Speaker, etc)
      const name = colB;   // B: Name or hymn title with number
      const extra = colC;  // C: Extra Info
    
      // Skip any row where the Name column (B) is empty — prevents rendering empty optional rows
      if (!name) {
        console.log(`[app] skipping empty agenda row ${i+1} (no name/title in column B)`);
        continue;
      }
    
      const key = normalizeItemKey(item);

      // handle hymns which may be in the "Name" column with leading number
      if(key.includes('hymn')) {
        // detect hymn number from name: leading number or trailing
        let hymnNumber = null;
        let hymnTitle = name;
        const m = name.match(/^\s*([0-9]{1,4})\s*[\.\-:]?\s*(.+)$/);
        if(m){
          hymnNumber = m[1];
          hymnTitle = m[2] || '';
        } else {
          // another pattern "1024. I have faith..." possibly includes dot + title
          const m2 = name.match(/([0-9]{3,4})/);
          if(m2) hymnNumber = m2[1];
        }

        const hymnUrl = getHymnUrl(hymnTitle, hymnNumber, extra);
        container.appendChild(createHymnCard(hymnTitle, hymnNumber, item, hymnUrl));
        any = true;
        continue;
      }

      // handle "Speaker" or "Speaker (Optional)" and similar
      if(key.startsWith('speaker') || key === 'testimony' || key.includes('testimon')) {
        // For "Speaker (Optional)" or "Speaker" just show as Speaker line
        container.appendChild(createRow('Speaker', name, ''));
        any = true;
        continue;
      }

      // Invocation, Benediction, Closing Prayer, Musical Number etc.
      if(key.includes('invocation') || key.includes('opening prayer') || key.includes('closing prayer') || key.includes('benediction') || key.includes('closing')) {
        container.appendChild(createRow(item, name, ''));
        any = true;
        continue;
      }

      if(key.includes('musical')) {
        container.appendChild(createRow('Musical Number', name, extra));
        any = true;
        continue;
      }

      // catch-all: render generic row
      container.appendChild(createRow(item, name, extra));
      any = true;
    }

    if(!any){
      container.innerHTML = '<div class="placeholder"><p class="muted">No agenda items found in Agenda sheet.</p></div>';
    }

    clearError();

    // collapsible toggles for Activities and Leadership remain same as before
    document.querySelectorAll('.collapsible-toggle').forEach(btn => {
      btn.addEventListener('click', () => {
        const target = btn.getAttribute('data-target');
        const panel = document.getElementById(target);
        const expanded = btn.getAttribute('aria-expanded') === 'true';
        btn.setAttribute('aria-expanded', (!expanded).toString());
        if(panel) panel.style.display = expanded ? 'none' : 'block';
      });
    });

    // optionally load activities and leadership from other sheets if you add them later

  } catch(err){
    console.error('[app] error', err);
    showError('Failed to fetch sheets: ' + err.message);
  }
}

// render header from admin map (keeps look)
function renderHeaderFromAdmin(map){
  const title = map['title'] || 'The Church of Jesus Christ of Latter-day Saints';
  const ward = map['ward'] || '';
  const stake = map['stake'] || '';
  const dateRaw = map['upcoming sunday date'] || map['upcoming sunday'] || '';
  const presiding = map['presiding'] || '—';
  const conducting = map['conducting'] || '—';
  const meetingType = map['meeting type'] || 'Sacrament Meeting';

  $('#meeting-heading').textContent = meetingType;
  $('#meeting-date').textContent = (dateRaw ? new Date(dateRaw).toLocaleDateString(undefined, { weekday:'long', month:'long', day:'numeric', year:'numeric' }) : '') + (ward ? `\n${ward} · ${stake}` : '');
  $('#presiding').textContent = presiding;
  $('#conducting').textContent = conducting;
}

run();
