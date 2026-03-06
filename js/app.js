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


function getAgendaIcon(type){

  if(type === "hymn"){
    return `<img src="./icons/hymn.svg" class="agenda-icon" alt="">`;
  }

  if(type === "speaker"){
    return `<img src="./icons/speaker.svg" class="agenda-icon" alt="">`;
  }

  if(type === "prayer"){
    return `<img src="./icons/prayer.svg" class="agenda-icon" alt="">`;
  }

  return "";
}


// slug helper
function slugify(text){
  if(!text) return '';
  return text.toString()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')   // remove accents
    .trim()
    .replace(/^[0-9]+\.\s*/,'')                        // strip leading "123. "
    .replace(/[’'"\.:,;!?\(\)\[\]\/]/g,'')             // remove punctuation
    .replace(/[^a-zA-Z0-9\s-]/g,'')                   // remove other non-safe chars
    .toLowerCase()
    .replace(/\s+/g,'-')
    .replace(/-+/g,'-')
    .replace(/^-|-$/g,'');
}

/**
 * title: hymn title (string)
 * hymnNumber: numeric or string (e.g. "193" or "1024")
 * extraInfo: contents of column C (hints like "Hymns", "Hymns for Home and Church", "Children's Songbook")
 * slugOverride: optional explicit slug from column D (if present)
 */
function getHymnUrl(title, hymnNumber, extraInfo, slugOverride){
  const extra = (extraInfo || '').toString().toLowerCase();
  const t = (title || '').toString().trim();
  const titleSlug = slugOverride ? String(slugOverride).trim() : slugify(t);
  const n = Number((hymnNumber !== undefined && hymnNumber !== null) ? String(hymnNumber).replace(/[^\d]/g,'') : NaN);

  // 1) If explicit slug provided in sheet (column D), use that with sensible path guessing
  if(titleSlug) {
    // prefer specific collections if extra hints them
    if(extra.includes('child') || extra.includes('songbook')) {
      return `https://www.churchofjesuschrist.org/study/manual/childrens-songbook/${titleSlug}?lang=eng`;
    }
    if(extra.includes('hymns for home') || extra.includes('home and church')) {
      return `https://www.churchofjesuschrist.org/study/music/hymns-for-home-and-church/${titleSlug}?lang=eng`;
    }
    // Classic hymn book
    return `https://www.churchofjesuschrist.org/study/manual/hymns/${titleSlug}?lang=eng`;
  }

  // 2) If numeric hymn is present, use it to choose general path
  if(!isNaN(n) && n > 0){
    if(n <= 341){
      // classic hymns: map number -> manual/hymns slug is reliable if we have title, but
      // if we only have number attempt canonical manual link via search by number (search fallback)
      // Prefer constructing study/manual/hymns/<slug> if title exists; otherwise, link to the hymns index anchor via search.
      // We'll return the manual path with the number if all else fails by searching the manual page.
      // Best experience: include title in sheet or add slug in column D.
      return `https://www.churchofjesuschrist.org/study/manual/hymns/${n}?lang=eng`;
    }
    if(n >= 1000){
      // treat as Hymns-for-Home numeric id: prefer titleSlug if present, otherwise link to collection anchor via search
      return `https://www.churchofjesuschrist.org/study/music/hymns-for-home-and-church?lang=eng#${n}`;
    }
    // default fallback
    return `https://www.churchofjesuschrist.org/search?q=${encodeURIComponent(String(n))}`;
  }

  // 3) If extra info points to a collection but no slug/number, build slug from title and try appropriate path
  if(extra.includes('child') || extra.includes('songbook')){
    if(titleSlug) return `https://www.churchofjesuschrist.org/study/manual/childrens-songbook/${titleSlug}?lang=eng`;
  }
  if(extra.includes('hymns for home') || extra.includes('home and church')){
    if(titleSlug) return `https://www.churchofjesuschrist.org/study/music/hymns-for-home-and-church/${titleSlug}?lang=eng`;
  }

  // 4) Last resort: if we have a title at all, do a site search for it
  if(t) return `https://www.churchofjesuschrist.org/search?q=${encodeURIComponent(t)}`;

  // nothing to link
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
      ${getAgendaIcon("hymn")}
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
    a.appendChild(el);
    return a;
  }
  return el;
}

function createRow(typeLabel, name, extra){
  const el = document.createElement('div');
  el.className = 'agenda-item';
  // pick icon type based on the label (you can map more rules here)
  let iconType = 'default';
  const tkey = (typeLabel||'').toString().toLowerCase();
  if(tkey.includes('hymn') || tkey.includes('sacrament')) iconType = 'hymn';
  else if(tkey.includes('speaker') || tkey.includes('testimon')) iconType = 'speaker';
  else if(tkey.includes('invocation') || tkey.includes('benediction') || tkey.includes('prayer')) iconType = 'prayer';
  // build
  el.innerHTML = `
    <div class="icon">${getAgendaIcon(iconType)}</div>
    
    <div class="content">
      <div class="title">${typeLabel}</div>
      <div class="sub">${name || ''}</div>
    </div>
    <div class="right">${extra || ''}</div>
  `;
  return el;
}


function createDivider(label){
  const el = document.createElement('div');
  el.className = 'agenda-divider';

  el.innerHTML = `
    <div class="divider-line"></div>
    <div class="divider-text">${label}</div>
    <div class="divider-line"></div>
  `;

  return el;
}


/**
 * Render the leadership rows into #leaders-list.
 * Expects rows in the format: Column A = Role/Key, Column B = Name, Column C = Contact
 */
/**
 * Render the leadership rows into #leaders-list as a 3-column table:
 * Column A = Role/Key, Column B = Name, Column C = Contact
 * No headers, no avatars.
 */
function renderLeadership(rows){
  const container = document.getElementById('leaders-list');
  if(!container) return;
  container.innerHTML = '';

  // skip header-like first row if it's a header
  let start = 0;
  if(rows[0] && rows[0][0]){
    const h = rows[0][0].toString().toLowerCase();
    if(h.includes('key') || h.includes('role') || h.includes('name') || h.includes('contact')) start = 1;
  }

  // create table
  const table = document.createElement('table');
  table.className = 'leadership-table';
  const tbody = document.createElement('tbody');

  for(let i = start; i < rows.length; i++){
    const r = rows[i];
    if(!r) continue;

    // If entire row is empty skip it
    const hasAny = (r[0]||'').toString().trim() || (r[1]||'').toString().trim() || (r[2]||'').toString().trim();
    if(!hasAny) continue;

    const role = (r[0]||'').toString().trim();
    const name = (r[1]||'').toString().trim();
    const contact = (r[2]||'').toString().trim();

    const tr = document.createElement('tr');

    // Role cell (first column)
    const tdRole = document.createElement('td');
    tdRole.className = 'lead-col-role';
    tdRole.textContent = role || '';
    tr.appendChild(tdRole);

    // Name cell (second column)
    const tdName = document.createElement('td');
    tdName.className = 'lead-col-name';
    tdName.textContent = name || '';
    tr.appendChild(tdName);

    // Contact cell (third column) — use formatContactLink to produce safe markup
    const tdContact = document.createElement('td');
    tdContact.className = 'lead-col-contact';
    // formatContactLink returns an HTML string (tel/mailto or text).
    // Use innerHTML intentionally because the string can include <a> tags.
    tdContact.innerHTML = contact ? formatContactLink(contact) : '';
    tr.appendChild(tdContact);

    tbody.appendChild(tr);
  }

  table.appendChild(tbody);
  container.appendChild(table);
}

/** Format contact value: tel: for phones, mailto: for emails, otherwise plain text */
function formatContactLink(contact){
  if(!contact) return '';
  const digits = contact.replace(/[^\d+]/g,'');
  const digitCount = (digits.match(/\d/g)||[]).length;
  if(digitCount >= 7){
    const tel = contact.replace(/[^\d+]/g,'');
    return `<a href="tel:${tel}" class="muted small">${contact}</a>`;
  }
  if(contact.includes('@')){
    return `<a href="mailto:${contact}" class="muted small">${contact}</a>`;
  }
  return `<span class="muted small">${contact}</span>`;
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
  let leadershipCsvUrl = config.leadership_csv_url || (config.sheet_id && config.leadership_gid ? buildCsvUrl(config.sheet_id, config.leadership_gid) : null);

  if(!adminCsvUrl){
    showError('No admin CSV URL available. Set admin_gid or admin_csv_url in config.json');
    return;
  }
  if(!agendaCsvUrl){
    showError('No agenda CSV URL available. Set agenda_gid or agenda_csv_url in config.json');
    return;
  }

  // fetch admin, agenda, and leadership in parallel
  try {
    // Build list of fetch promises, include leadership only if a URL is available
    const fetches = [ fetch(adminCsvUrl), fetch(agendaCsvUrl) ];
    if (leadershipCsvUrl) fetches.push(fetch(leadershipCsvUrl));

    const responses = await Promise.all(fetches);
    // responses[0] -> admin, [1] -> agenda, [2] -> leadership (if provided)
    const admResp = responses[0];
    const agResp = responses[1];
    const leadResp = responses[2] || null;

    if(!admResp.ok) throw new Error('Admin sheet fetch failed: ' + admResp.status);
    if(!agResp.ok) throw new Error('Agenda sheet fetch failed: ' + agResp.status);
    if(leadResp && !leadResp.ok) throw new Error('Leadership sheet fetch failed: ' + leadResp.status);

    const admText = await admResp.text();
    const agText = await agResp.text();
    const leadText = leadResp ? await leadResp.text() : null;

    // Quick HTML detection (permission problems often return HTML)
    if(/<html|doctype html/i.test(admText.slice(0,200))) { showError('Admin sheet returned HTML (not public)'); return; }
    if(/<html|doctype html/i.test(agText.slice(0,200))) { showError('Agenda sheet returned HTML (not public)'); return; }
    if(leadText && /<html|doctype html/i.test(leadText.slice(0,200))) { showError('Leadership sheet returned HTML (not public)'); return; }

    // Parse CSV rows
    const admRows = parseCSVtoRows(admText);
    const agRows = parseCSVtoRows(agText);
    const leadRows = leadText ? parseCSVtoRows(leadText) : null;

    // build admin map
    const adminMap = {};
    for(let i=0;i<admRows.length;i++){
      const r = admRows[i];
      if(!r || !r[0]) continue;
      adminMap[(r[0]||'').toString().trim().toLowerCase()] = (r[1]||'').toString().trim();
    }

    // render header
    renderHeaderFromAdmin(adminMap);

    // parse and render agenda rows in order (your existing loop)
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
      const colD = (r[3] || '').toString().trim(); // Optional explicit hymn slug (column D)
    
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
      const itemRaw = colA;                // raw Item text from sheet (may include "(Optional)")
      // Remove trailing " (Optional)" (case-insensitive, flexible spacing) for display
      const displayItem = itemRaw.replace(/\s*\(\s*optional\s*\)\s*$/i, '').trim();
      const item = displayItem;            // use this cleaned label for UI
      const name = colB;                   // B: Name or hymn title with number
      const extra = colC;                  // C: Extra Info
      const slugOverride = colD;           // D: optional explicit slug (use when site slug differs)

      // Special divider for "Administration of the Sacrament"
      if(item.toLowerCase().includes('administration of the sacrament')){
        container.appendChild(createDivider(item));
        any = true;
        continue;
      }
      
      // Skip any row where the Name column (B) is empty — prevents rendering empty optional rows
      if (!name) {
        console.log(`[app] skipping empty agenda row ${i+1} (no name/title in column B)`);
        continue;
      }
    
      // Use cleaned item when building the match key so "Speaker (Optional)" matches "speaker"
      const key = normalizeItemKey(item);
    
      // handle hymns which may be in the "Name" column with leading number
      if (key.includes('hymn')) {
        // detect hymn number from name: leading number or trailing
        let hymnNumber = null;
        let hymnTitle = name;
        const m = name.match(/^\s*([0-9]{1,4})\s*[\.\-:]?\s*(.+)$/);
        if (m) {
          hymnNumber = m[1];
          hymnTitle = m[2] || '';
        } else {
          const m2 = name.match(/([0-9]{3,4})/);
          if (m2) hymnNumber = m2[1];
        }
    
        // Pass slugOverride into getHymnUrl so explicit slugs in column D are used when present
        const hymnUrl = getHymnUrl(hymnTitle, hymnNumber, extra, slugOverride);
        container.appendChild(createHymnCard(hymnTitle, hymnNumber, item, hymnUrl)); // use cleaned item label
        any = true;
        continue;
      }
    
      // handle "Speaker" or "Speaker (Optional)" and similar
      if (key.startsWith('speaker') || key === 'testimony' || key.includes('testimon')) {
        container.appendChild(createRow('Speaker', name, ''));
        any = true;
        continue;
      }
    
      // Invocation, Benediction, Closing Prayer, Musical Number etc.
      if (key.includes('invocation') || key.includes('opening prayer') || key.includes('closing prayer') || key.includes('benediction') || key.includes('closing')) {
        container.appendChild(createRow(item, name, '')); // use cleaned item label for display
        any = true;
        continue;
      }
    
      if (key.includes('musical')) {
        container.appendChild(createRow('Musical Number', name, extra));
        any = true;
        continue;
      }
    
      // catch-all: render generic row using cleaned label
      container.appendChild(createRow(item, name, extra));
      any = true;
    }

    if(!any){
      container.innerHTML = '<div class="placeholder"><p class="muted">No agenda items found in Agenda sheet.</p></div>';
    }

    // Render leadership list (if the leadership CSV was provided)
    if (leadRows && leadRows.length) {
      renderLeadership(leadRows);
    } else {
      // clear or show a friendly message in the leaders list
      const ll = document.getElementById('leaders-list');
      if (ll) ll.innerHTML = '<div class="muted small">No leadership data found.</div>';
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
