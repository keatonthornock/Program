// js/app.js
// Enhanced: load Administrative + Agenda CSVs, render agenda items, and handle conference events
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
    if(ln === undefined || ln === null) continue;
    if(ln.trim() === '') continue;
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

/* ---------- hymn/url helpers ---------- */
function slugify(text){
  if(!text) return '';
  return text.toString()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .trim()
    .replace(/^[0-9]+\.\s*/,'')
    .replace(/[’'"\.:,;!?\(\)\[\]\/]/g,'')
    .replace(/[^a-zA-Z0-9\s-]/g,'')
    .toLowerCase()
    .replace(/\s+/g,'-')
    .replace(/-+/g,'-')
    .replace(/^-|-$/g,'');
}

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

/* ---------- small UI helpers ---------- */
function showError(msg){
  const n = $('#notice'); if(n) { n.hidden = false; n.textContent = msg; } else console.warn(msg);
}
function clearError(){ const n = $('#notice'); if(n) { n.hidden = true; n.textContent = ''; } }

function normalizeItemKey(s){ return (s||'').toString().trim().toLowerCase(); }

function shouldRenderAgendaItem(key, meetingType){

  const isTestimony = meetingType.includes('testimony');
  const isSacrament = meetingType.includes('sacrament') || meetingType === '' || meetingType === 'sacrament meeting';

  // If it's neither sacrament nor testimony, agenda items are not rendered
  if(!isSacrament && !isTestimony){
    return false;
  }

  // Testimony meeting rules
  if(isTestimony){

    // remove speakers
    if(key.includes('speaker')) return false;

    // remove musical numbers
    if(key.includes('musical')) return false;

    // only allow opening / sacrament / closing hymns
    if(key.includes('hymn') && !/opening|sacrament|closing/i.test(key)){
      return false;
    }

  }

  return true;
}

/* ---------- icon/image helpers ---------- */
function getAgendaIcon(type){
  if(type === "hymn") return `<img src="./icons/hymn.svg" class="agenda-icon" alt="">`;
  if(type === "speaker") return `<img src="./icons/speaker.svg" class="agenda-icon" alt="">`;
  if(type === "prayer") return `<img src="./icons/prayer.svg" class="agenda-icon" alt="">`;
  return "";
}

/* ---------- rendering helpers ---------- */
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
    <div class="right">${url? `
      <svg class="hymn-arrow" viewBox="0 0 24 24">
        <path d="M9 6l6 6-6 6"/>
      </svg>` : ''}</div>
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

function createRow(typeLabel, name, extra, iconType = 'default'){
  const el = document.createElement('div');
  el.className = 'agenda-item';
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

/* ---------- Conference event parsing + rendering ---------- */
function parseConferenceEvents(admRows){
  if(!Array.isArray(admRows)) return [];
  const headerNames = ['event','location','date','time','address'];
  for(let i=0;i<admRows.length;i++){
    const row = admRows[i].map(c => (c||'').toString().trim().toLowerCase());
    // require row[0] === 'event' or includes 'event' to be robust
    if(row[0] && row[0].includes('event')){
      // check next columns for at least 'location' or 'date' to be confident
      if(row[1] && (row[1].includes('location') || row[1].includes('loc'))) {
        const events = [];
        for(let j=i+1;j<admRows.length;j++){
          const r = admRows[j];
          if(!r || !r[0] || r[0].toString().trim() === '') break;
          const ev = {
            event: (r[0]||'').toString().trim(),
            location: (r[1]||'').toString().trim(),
            date: (r[2]||'').toString().trim(),
            time: (r[3]||'').toString().trim(),
            address: (r[4]||'').toString().trim()
          };
          if(ev.event || ev.location || ev.date || ev.time || ev.address) events.push(ev);
        }
        return events;
      }
    }
  }
  return [];
}

function createEventCard(ev){
  const el = document.createElement('div');
  el.className = 'event-card';
  const mapHref = ev.address ? `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(ev.address)}` : '';
  el.innerHTML = `
    <div class="event-left">
      <div class="event-title">${ev.event || ''}</div>
      <div class="event-meta">
        ${ev.location ? `<div class="event-loc">${ev.location}</div>` : ''}
        ${ev.date || ev.time ? `<div class="event-time">${ev.date ? ev.date : ''}${ev.date && ev.time ? ' · ' : ''}${ev.time ? ev.time : ''}</div>` : ''}
        ${ev.address ? `<div class="event-address"><a href="${mapHref}" target="_blank" rel="noopener">${ev.address}</a></div>` : ''}
      </div>
    </div>
    <div class="event-right">
      <svg class="hymn-arrow" viewBox="0 0 24 24"><path d="M9 6l6 6-6 6"/></svg>
    </div>
  `;
  if(mapHref){
    el.addEventListener('click', (e) => {
      // let native anchor clicks behave normally
      const a = e.target.closest('a');
      if(a) return;
      window.open(mapHref, '_blank', 'noopener');
    });
    el.style.cursor = 'pointer';
  }
  return el;
}

/* ---------- Leadership rendering ---------- */
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
    const tdRole = document.createElement('td'); tdRole.className='lead-col-role'; tdRole.textContent = role || ''; tr.appendChild(tdRole);
    const tdName = document.createElement('td'); tdName.className='lead-col-name'; tdName.textContent = name || ''; tr.appendChild(tdName);
    const tdContact = document.createElement('td'); tdContact.className='lead-col-contact'; tdContact.innerHTML = contact ? formatContactLink(contact) : ''; tr.appendChild(tdContact);
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

/* ---------- Header rendering + conference logic ---------- */
function renderHeaderFromAdmin(map, admRows){
  const title = map['title'] || 'The Church of Jesus Christ of Latter-day Saints';
  const ward = map['ward'] || '';
  const stake = map['stake'] || '';
  const dateRaw = map['upcoming sunday date'] || map['upcoming sunday'] || '';
  const presiding = map['presiding'] || '';
  const conducting = map['conducting'] || '';
  const meetingTypeRaw = (map['meeting type'] || '').toString();
  const meetingType = meetingTypeRaw.trim();
  const chorister = map['chorister'] || '';
  const organist = map['organist'] || '';

  $('#meeting-heading').textContent = meetingType || 'Sacrament Meeting';
  $('#meeting-date').textContent = (dateRaw ? new Date(dateRaw).toLocaleDateString(undefined, { weekday:'long', month:'long', day:'numeric', year:'numeric' }) : '') + (ward ? `\n${ward} · ${stake}` : '');

  // special handling for stake/general conferences: hide meta box entirely
  const isStakeConference = meetingType.toLowerCase().includes('stake conference') || meetingType.toLowerCase().includes('stake meeting') || meetingType.toLowerCase() === 'stake conference' || meetingType.toLowerCase() === 'stake';
  const isGeneralConference = meetingType.toLowerCase().includes('general conference') || meetingType.toLowerCase() === 'general conference' || meetingType.toLowerCase().includes('general');

  if(isStakeConference || isGeneralConference){
    const mb = document.querySelector('.meta-box'); if(mb) mb.style.display = 'none';
    const me = document.querySelector('.meta-extra'); if(me) me.style.display = 'none';
  } else {
    const mb = document.querySelector('.meta-box'); if(mb) mb.style.display = '';
    const me = document.querySelector('.meta-extra'); if(me) me.style.display = '';

    // Presiding
    if(presiding){
      const pelt = document.getElementById('presiding'); if(pelt) pelt.textContent = presiding;
      const pLine = document.getElementById('presiding-line'); if(pLine) pLine.style.display = '';
    } else {
      const pLine = document.getElementById('presiding-line'); if(pLine) pLine.style.display = 'none';
    }

    // Conducting placed below Presiding (left-aligned)
    if(conducting){
      const celt = document.getElementById('conducting'); if(celt) celt.textContent = conducting;
      const cLine = document.getElementById('conducting-line'); if(cLine) cLine.style.display = '';
    } else {
      const cLine = document.getElementById('conducting-line'); if(cLine) cLine.style.display = 'none';
    }

    // Chorister & Organist lines (outside the meta-box, left & right)
    if(chorister){
      const ch = document.getElementById('chorister'); if(ch) ch.textContent = chorister;
      const chLine = document.getElementById('chorister-line'); if(chLine) chLine.style.display = '';
    } else {
      const chLine = document.getElementById('chorister-line'); if(chLine) chLine.style.display = 'none';
    }

    if(organist){
      const og = document.getElementById('organist'); if(og) og.textContent = organist;
      const ogLine = document.getElementById('organist-line'); if(ogLine) ogLine.style.display = '';
    } else {
      const ogLine = document.getElementById('organist-line'); if(ogLine) ogLine.style.display = 'none';
    }
  }

  // If Stake conference, render events from Administrative rows (place into program-content as first child)
  if (isStakeConference) {
    const events = parseConferenceEvents(admRows);

    let container = document.getElementById('conference-events');
    if (!container) {
      container = document.createElement('div');
      container.id = 'conference-events';
      container.className = 'conference-events-body';
      const program = document.getElementById('program');
      const progContent = program ? program.querySelector('#program-content') : null;
      if (progContent) {
        // Prefer placing the events immediately *after* the meeting-type placeholder (so cards are under "Stake Conference")
        const placeholder = progContent.querySelector('.meeting-placeholder');
        if (placeholder) {
          // insert after placeholder
          placeholder.after(container);
        } else {
          // otherwise append to the end of the program content (right above the agenda items)
          progContent.appendChild(container);
        }
      } else {
        // fallback: append directly beneath the program card
        if (program && program.parentNode) program.parentNode.insertBefore(container, program.nextSibling);
        else document.querySelector('.app').appendChild(container);
      }
    }

    container.innerHTML = '';
    if (events && events.length) {
      events.forEach(ev => container.appendChild(createEventCard(ev)));
    } else {
      container.innerHTML = `<div class="muted small">No stake conference events found in Administrative sheet.</div>`;
    }
  } else {
    const container = document.getElementById('conference-events');
    if (container && container.parentNode) container.parentNode.removeChild(container);
  }

  // store meeting type on body dataset for other logic if needed
  document.body.dataset.meetingType = meetingType.toLowerCase();
}

/* ---------- main run flow ---------- */
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
      if(!r) continue;
      if(!r[0]) continue;
      adminMap[(r[0]||'').toString().trim().toLowerCase()] = (r[1]||'').toString().trim();
    }

    // parse and render agenda rows in order
    const container = $('#program-content');
    container.innerHTML = '';
    let any = false;

    const meetingType = (adminMap['meeting type'] || '').toString().toLowerCase();
    const isTestimony = meetingType.includes('testimony');
    const isSacrament = meetingType.includes('sacrament') || meetingType === '' || meetingType === 'sacrament meeting';

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

      // Administration divider special
      if(item.toLowerCase().includes('administration of the sacrament')){
        if(isSacrament || isTestimony){
          container.appendChild(createDivider(item));
          any = true;
        }
        if(isTestimony){
          const tb = document.createElement('div');
          tb.className = 'testimony-banner';
          tb.innerHTML = `<div class="testimony-text">Testimonies of the Congregation</div>`;
          container.appendChild(tb);
          any = true;
        }
        continue;
      }

      if(!name) continue;

      const key = normalizeItemKey(item);

      if(!shouldRenderAgendaItem(key, meetingType)){
        continue;
      }

      // hymn handling
      if(key.includes('hymn')){
        let hymnNumber = null;
        let hymnTitle = name;
        const m = name.match(/^\s*([0-9]{1,4})\s*[\.\-:]?\s*(.+)$/);
        if (m) { hymnNumber = m[1]; hymnTitle = m[2] || ''; }
        else {
          const m2 = name.match(/([0-9]{3,4})/);
          if(m2) hymnNumber = m2[1];
        }
        const hymnUrl = getHymnUrl(hymnTitle, hymnNumber, extra, slugOverride);
        container.appendChild(createHymnCard(hymnTitle, hymnNumber, item, hymnUrl));
        any = true;
        continue;
      }

      if(key.startsWith('speaker') || key === 'testimony' || key.includes('testimon')){
        container.appendChild(createRow('Speaker', name, '', 'speaker'));
        any = true;
        continue;
      }

      if(key.includes('invocation') || key.includes('opening prayer') || key.includes('closing prayer') || key.includes('benediction') || key.includes('closing')){
        container.appendChild(createRow(item, name, '', 'prayer'));
        any = true;
        continue;
      }

      if(key.includes('musical')){
        if(!isTestimony){
          container.appendChild(createRow('Musical Number', name, extra));
          any = true;
        }
        continue;
      }

      // fallback generic
      container.appendChild(createRow(item, name, extra));
      any = true;
    }

    // If meeting type is neither sacrament nor testimony -> show centered placeholder with meeting-type text
    if(!isSacrament && !isTestimony){
      const pc = document.getElementById('program-content');
      if(pc){
        // If it's a Stake Conference, render the title + events in a single wrapper inside the program card
        const isStakeConference = (adminMap['meeting type'] || '').toString().toLowerCase().includes('stake');
        if(isStakeConference){
          // clear any existing content
          pc.innerHTML = '';
    
          // wrapper that contains the title and the events (keeps them inside the same card)
          const stakeWrapper = document.createElement('div');
          stakeWrapper.className = 'stake-wrapper';
    
          // Title (small padded area only for the text, not a giant box)
          const titleDiv = document.createElement('div');
          titleDiv.className = 'stake-title';
          titleDiv.textContent = (adminMap['meeting type']||'Stake Conference').toString();
          stakeWrapper.appendChild(titleDiv);
    
          // events container (cards will be appended here)
          const eventsContainer = document.createElement('div');
          eventsContainer.id = 'conference-events';
          eventsContainer.className = 'conference-events-body';
    
          // parse events from admin rows and render them here
          const events = parseConferenceEvents(admRows);
          if(events && events.length){
            events.forEach(ev => eventsContainer.appendChild(createEventCard(ev)));
          } else {
            const no = document.createElement('div');
            no.className = 'muted small';
            no.textContent = 'No stake conference events found in Administrative sheet.';
            eventsContainer.appendChild(no);
          }
    
          stakeWrapper.appendChild(eventsContainer);
    
          // append wrapper to program content (so title sits above the cards, both inside the same card)
          pc.appendChild(stakeWrapper);
          any = true;
        } else {
          // generic non-sacrament non-testimony placeholder (unchanged)
          pc.innerHTML = `<div class="meeting-placeholder"><div class="placeholder-text">${(adminMap['meeting type']||'').toString()}</div></div>`;
          any = true;
        }
      }
    }

    if(!any){
      container.innerHTML = '<div class="placeholder"><p class="muted">No agenda items found in Agenda sheet.</p></div>';
    }

    // Leadership
    if (leadRows && leadRows.length) {
      renderLeadership(leadRows);
    } else {
      const ll = document.getElementById('leaders-list');
      if (ll) ll.innerHTML = '<div class="muted small">No leadership data found.</div>';
    }

    // Now render header and conference events AFTER the program content has been generated.
    // This prevents inserted conference-events from being wiped by the agenda rendering.
    renderHeaderFromAdmin(adminMap, admRows);

    clearError();

    // wire collapsibles
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

run();
