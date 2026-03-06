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
  const lines = text.split(/\r?\n/);
  const rows = [];
  for(const ln of lines){
    if(!ln) continue;
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


// slug helper (unchanged)
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

  if(titleSlug) {
    if(extra.includes('child') || extra.includes('songbook')) {
      return `https://www.churchofjesuschrist.org/study/manual/childrens-songbook/${titleSlug}?lang=eng`;
    }
    if(extra.includes('hymns for home') || extra.includes('home and church')) {
      return `https://www.churchofjesuschrist.org/study/music/hymns-for-home-and-church/${titleSlug}?lang=eng`;
    }
    return `https://www.churchofjesuschrist.org/study/manual/hymns/${titleSlug}?lang=eng`;
  }

  if(!isNaN(n) && n > 0){
    if(n <= 341){
      return `https://www.churchofjesuschrist.org/study/manual/hymns/${n}?lang=eng`;
    }
    if(n >= 1000){
      return `https://www.churchofjesuschrist.org/study/music/hymns-for-home-and-church?lang=eng#${n}`;
    }
    return `https://www.churchofjesuschrist.org/search?q=${encodeURIComponent(String(n))}`;
  }

  if(extra.includes('child') || extra.includes('songbook')){
    if(titleSlug) return `https://www.churchofjesuschrist.org/study/manual/childrens-songbook/${titleSlug}?lang=eng`;
  }
  if(extra.includes('hymns for home') || extra.includes('home and church')){
    if(titleSlug) return `https://www.churchofjesuschrist.org/study/music/hymns-for-home-and-church/${titleSlug}?lang=eng`;
  }

  if(t) return `https://www.churchofjesuschrist.org/search?q=${encodeURIComponent(t)}`;

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
    <div class="right">
      ${url ? `
      <svg class="hymn-arrow" viewBox="0 0 24 24" aria-hidden="true">
        <path d="M9 6l6 6-6 6"/>
      </svg>
      ` : ''}
    </div>
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

function createRow(typeLabel, name, extra, iconType){
  const el = document.createElement('div');
  el.className = 'agenda-item';
  let resolvedIcon = iconType || 'default';
  const tkey = (typeLabel||'').toString().toLowerCase();
  if(!iconType){
    if(tkey.includes('hymn') || tkey.includes('sacrament')) resolvedIcon = 'hymn';
    else if(tkey.includes('speaker') || tkey.includes('testimon')) resolvedIcon = 'speaker';
    else if(tkey.includes('invocation') || tkey.includes('benediction') || tkey.includes('prayer')) resolvedIcon = 'prayer';
  }

  el.innerHTML = `
    <div class="icon">${getAgendaIcon(resolvedIcon)}</div>
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

function createTestimonyBanner(){
  const el = document.createElement('div');
  el.className = 'testimony-banner';
  el.innerHTML = `<div class="testimony-text">Testimonies of the Congregation</div>`;
  return el;
}

function createMeetingPlaceholder(meetingTypeLabel){
  const el = document.createElement('div');
  el.className = 'meeting-placeholder';
  el.innerHTML = `<div class="placeholder-text">${meetingTypeLabel}</div>`;
  return el;
}

/**
 * Render leadership as a 3-column table (unchanged from previous)
 */
function renderLeadership(rows){
  const container = document.getElementById('leaders-list');
  if(!container) return;
  container.innerHTML = '';

  let start = 0;
  if(rows[0] && rows[0][0]){
    const h = rows[0][0].toString().toLowerCase();
    if(h.includes('key') || h.includes('role') || h.includes('name') || h.includes('contact')) start = 1;
  }

  const table = document.createElement('table');
  table.className = 'leadership-table';
  const tbody = document.createElement('tbody');

  for(let i = start; i < rows.length; i++){
    const r = rows[i];
    if(!r) continue;
    const hasAny = (r[0]||'').toString().trim() || (r[1]||'').toString().trim() || (r[2]||'').toString().trim();
    if(!hasAny) continue;

    const role = (r[0]||'').toString().trim();
    const name = (r[1]||'').toString().trim();
    const contact = (r[2]||'').toString().trim();

    const tr = document.createElement('tr');

    const tdRole = document.createElement('td');
    tdRole.className = 'lead-col-role';
    tdRole.textContent = role || '';
    tr.appendChild(tdRole);

    const tdName = document.createElement('td');
    tdName.className = 'lead-col-name';
    tdName.textContent = name || '';
    tr.appendChild(tdName);

    const tdContact = document.createElement('td');
    tdContact.className = 'lead-col-contact';
    tdContact.innerHTML = contact ? formatContactLink(contact) : '';
    tr.appendChild(tdContact);

    tbody.appendChild(tr);
  }

  table.appendChild(tbody);
  container.appendChild(table);
}

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

  try {
    const fetches = [ fetch(adminCsvUrl), fetch(agendaCsvUrl) ];
    if (leadershipCsvUrl) fetches.push(fetch(leadershipCsvUrl));

    const responses = await Promise.all(fetches);
    const admResp = responses[0];
    const agResp = responses[1];
    const leadResp = responses[2] || null;

    if(!admResp.ok) throw new Error('Admin sheet fetch failed: ' + admResp.status);
    if(!agResp.ok) throw new Error('Agenda sheet fetch failed: ' + agResp.status);
    if(leadResp && !leadResp.ok) throw new Error('Leadership sheet fetch failed: ' + leadResp.status);

    const admText = await admResp.text();
    const agText = await agResp.text();
    const leadText = leadResp ? await leadResp.text() : null;

    if(/<html|doctype html/i.test(admText.slice(0,200))) { showError('Admin sheet returned HTML (not public)'); return; }
    if(/<html|doctype html/i.test(agText.slice(0,200))) { showError('Agenda sheet returned HTML (not public)'); return; }
    if(leadText && /<html|doctype html/i.test(leadText.slice(0,200))) { showError('Leadership sheet returned HTML (not public)'); return; }

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

    // Determine meeting type for special behavior
    const meetingTypeRaw = (adminMap['meeting type'] || adminMap['meeting'] || '').toString().trim();
    const meetingTypeKey = meetingTypeRaw.toLowerCase();

    const container = $('#program-content');
    container.innerHTML = '';
    let any = false;

    // If it's a non-sacrament and non-testimony meeting => render centered placeholder and skip rest
    const isSacrament = meetingTypeKey.includes('sacrament');
    const isTestimony = meetingTypeKey.includes('testimony');

    if(!isSacrament && !isTestimony){
      // Render the big centered placeholder and stop
      container.appendChild(createMeetingPlaceholder(meetingTypeRaw || 'Meeting'));
      any = true;
    } else {
      // Normal or Testimony flow: iterate agenda rows but with filters for testimony
      for (let i = 0; i < agRows.length; i++) {
        const r = agRows[i];
        if (!r || !r[0]) continue;

        const colA = (r[0] || '').toString().trim();
        const colB = (r[1] || '').toString().trim();
        const colC = (r[2] || '').toString().trim();
        const colD = (r[3] || '').toString().trim();

        const aKey = colA.toLowerCase();
        const bKey = colB.toLowerCase();
        const cKey = colC.toLowerCase();
        const looksLikeHeader = aKey === 'item' || bKey === 'name' || cKey === 'extra info' || (aKey.includes('item') && bKey.includes('name'));
        if (looksLikeHeader) continue;

        const itemRaw = colA;
        const displayItem = itemRaw.replace(/\s*\(\s*optional\s*\)\s*$/i, '').trim();
        const item = displayItem;
        const name = colB;
        const extra = colC;
        const slugOverride = colD;

        // If this row is the "Administration of the Sacrament" divider -- handle always in sacram/testimony as needed
        if(item.toLowerCase().includes('administration of the sacrament')){
          container.appendChild(createDivider(item));
          any = true;
          // For testimony meetings insert the "Testimonies..." banner immediately after the divider
          if(isTestimony){
            container.appendChild(createTestimonyBanner());
          }
          continue;
        }

        // Skip rows with no name (common for optional rows etc.)
        if(!name) continue;

        const key = normalizeItemKey(item);

        // If meeting type is 'Testimony' we only allow a small set of item types:
        if(isTestimony){
          // If it's a hymn, allow only if it is NOT an intermediate hymn.
          if(key.includes('hymn')){
            // skip intermediate or "special" or "musical hymn"
            if(key.includes('intermediate') || key.includes('special') || key.includes('musical')) {
              continue;
            }
            // Accept opening, sacrament, closing, or generic 'hymn' labels
            const allowedHymnMatch = /opening|sacrament|closing|hymn/;
            if(!allowedHymnMatch.test(key)) continue;
            // proceed to render hymn card
            let hymnNumber = null;
            let hymnTitle = name;
            const m = name.match(/^\s*([0-9]{1,4})\s*[\.\-:]?\s*(.+)$/);
            if (m) { hymnNumber = m[1]; hymnTitle = m[2] || ''; }
            else { const m2 = name.match(/([0-9]{3,4})/); if(m2) hymnNumber = m2[1]; }
            const hymnUrl = getHymnUrl(hymnTitle, hymnNumber, extra, slugOverride);
            container.appendChild(createHymnCard(hymnTitle, hymnNumber, item, hymnUrl));
            any = true;
            continue;
          }

          // allow invocation/opening prayer/benediction/closing
          if(key.includes('invocation') || key.includes('opening prayer') || key.includes('benediction') || key.includes('closing') ){
            container.appendChild(createRow(item, name, '','prayer'));
            any = true;
            continue;
          }

          // explicitly skip musical, speaker, testimony (individual), and other items
          // (so nothing else is rendered)
          continue;
        }

        // Standard Sacrament meeting flow (render everything as you had before, including musical numbers)
        // Handle hymn
        if (key.includes('hymn')) {
          let hymnNumber = null;
          let hymnTitle = name;
          const m = name.match(/^\s*([0-9]{1,4})\s*[\.\-:]?\s*(.+)$/);
          if (m) { hymnNumber = m[1]; hymnTitle = m[2] || ''; }
          else { const m2 = name.match(/([0-9]{3,4})/); if(m2) hymnNumber = m2[1]; }
          const hymnUrl = getHymnUrl(hymnTitle, hymnNumber, extra, slugOverride);
          container.appendChild(createHymnCard(hymnTitle, hymnNumber, item, hymnUrl));
          any = true;
          continue;
        }

        // Speakers
        if (key.startsWith('speaker') || key === 'testimony' || key.includes('testimon')) {
          container.appendChild(createRow('Speaker', name, ''));
          any = true;
          continue;
        }

        // Invocation, Benediction, Closing Prayer, Musical Number etc.
        if (key.includes('invocation') || key.includes('opening prayer') || key.includes('closing prayer') || key.includes('benediction') || key.includes('closing')) {
          container.appendChild(createRow(item, name, ''));
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
      } // end agRows loop
    } // end sacram/testimony block

    if(!any){
      container.innerHTML = '<div class="placeholder"><p class="muted">No agenda items found in Agenda sheet.</p></div>';
    }

    // Render leadership list (if provided)
    if (leadRows && leadRows.length) {
      renderLeadership(leadRows);
    } else {
      const ll = document.getElementById('leaders-list');
      if (ll) ll.innerHTML = '<div class="muted small">No leadership data found.</div>';
    }

    clearError();

    document.querySelectorAll('.collapsible-toggle').forEach(btn => {
      btn.addEventListener('click', () => {
        const target = btn.getAttribute('data-target');
        const panel = document.getElementById(target);
        const expanded = btn.getAttribute('aria-expanded') === 'true';
        btn.setAttribute('aria-expanded', (!expanded).toString());
        if(panel) panel.style.display = expanded ? 'none' : 'block';
      });
    });

  } catch(err){
    console.error('[app] error', err);
    showError('Failed to fetch sheets: ' + err.message);
  }
}

// render header from admin map (keeps look) — unchanged behavior
function renderHeaderFromAdmin(map){
  const title = map['title'] || 'The Church of Jesus Christ of Latter-day Saints';
  const ward = map['ward'] || '';
  const stake = map['stake'] || '';
  const dateRaw = map['upcoming sunday date'] || map['upcoming sunday'] || '';
  const presiding = map['presiding'] || '—';
  const conducting = map['conducting'] || '—';
  const meetingType = map['meeting type'] || 'Sacrament Meeting';
  const chorister = map['chorister'] || '';
  const organist = map['organist'] || '';

  $('#meeting-heading').textContent = meetingType;
  $('#meeting-date').textContent = (dateRaw ? new Date(dateRaw).toLocaleDateString(undefined, { weekday:'long', month:'long', day:'numeric', year:'numeric' }) : '') + (ward ? `\n${ward} · ${stake}` : '');

  // Presiding & Conducting lines
  const presEl = $('#presiding');
  const condEl = $('#conducting');
  if(presEl) presEl.textContent = presiding || '';
  if(condEl) condEl.textContent = conducting || '';

  // Chorister & Organist lines (these elements are expected in HTML with ids chorister/organist)
  const chorEl = $('#chorister');
  const orgEl = $('#organist');
  if(chorEl){
    if(chorister) { chorEl.textContent = chorister; chorEl.style.display = ''; }
    else { chorEl.style.display = 'none'; }
  }
  if(orgEl){
    if(organist) { orgEl.textContent = organist; orgEl.style.display = ''; }
    else { orgEl.style.display = 'none'; }
  }
}

run();
