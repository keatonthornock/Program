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

// Helper: create a compact display string for a URL (origin + truncated path)
function getDisplayUrl(rawUrl, maxLen = 48){
  try {
    const u = new URL(rawUrl);
    const origin = u.hostname.replace(/^www\./,'');
    let path = (u.pathname === '/' ? '' : u.pathname);
    if(u.search) path += u.search;
    // normalize
    path = path.replace(/\/$/,'');
    let display = origin + (path ? path : '');
    if(display.length <= maxLen) return display;
    // truncate: keep origin + leading part of path
    const remaining = maxLen - origin.length - 1; // leave room for slash/ellipsis
    if(remaining <= 6) return origin;
    const truncated = path.slice(0, Math.max(remaining-1, 0));
    return origin + (truncated ? truncated + '…' : '…');
  } catch(e) {
    // fallback: trim and ellipsize raw
    return rawUrl.length > maxLen ? rawUrl.slice(0, maxLen-1) + '…' : rawUrl;
  }
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

/* ---------- ensure external links have a scheme so the browser doesn't treat them as relative ---------- */
function normalizeHref(href){
  if(!href) return '';
  href = href.toString().trim();
  // Allow protocol-relative (//), mailto:, tel:, and already-schemed urls
  if (/^(\/\/|[a-z][a-z0-9+.-]*:)/i.test(href)) return href;
  return 'https://' + href;
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

// small html-escape helper for safe insertion of text
function escapeHtml(s){
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// find executive secretary row from leadership rows array (rows = parseCSVtoRows result)
function findExecutiveSecretaryFromRows(rows){
  if(!Array.isArray(rows)) return null;
  let start = 0;
  if(rows[0] && rows[0][0]){
    const h = rows[0][0].toString().toLowerCase();
    if(h.includes('role') || h.includes('key') || h.includes('name') || h.includes('contact')) start = 1;
  }
  for(let i = start; i < rows.length; i++){
    const r = rows[i];
    if(!r || !r[0]) continue;
    const role = (r[0]||'').toString().trim().toLowerCase();
    if(/executive secretary|exec(utive)? sec|ward executive/i.test(role)){
      const name = (r[1]||'').toString().trim();
      const contact = (r[2]||'').toString().trim();
      return { name, contact };
    }
  }
  return null;
}

function normalizeTel(contact){
  if(!contact) return null;
  const digits = contact.replace(/[^\d+]/g,'');
  if((digits.match(/\d/g)||[]).length >= 7) return digits;
  return null;
}

function updateAppointmentsInfoBox(exec){
  const info = document.getElementById('info-note');
  if(!info) return;
  if(exec && (exec.name || exec.contact)){
    const nameHtml = exec.name ? `<strong>${escapeHtml(exec.name)}</strong>` : '';
    const tel = normalizeTel(exec.contact);
    const phoneHtml = exec.contact ? (tel ? `<a href="tel:${tel}">${escapeHtml(exec.contact)}</a>` : `<span>${escapeHtml(exec.contact)}</span>`) : '';
    info.hidden = false;
    info.innerHTML = `Appointments with the Bishop can be scheduled with the Ward Executive Secretary: ${nameHtml} ${phoneHtml}`.trim();
  } else {
    // hide if no data
    info.hidden = true;
    info.innerHTML = '';
  }
}

// small html-escape helper for safe insertion of text
function escapeHtml(s){
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// find executive secretary row from leadership rows array (rows = parseCSVtoRows result)
function findExecutiveSecretaryFromRows(rows){
  if(!Array.isArray(rows)) return null;
  let start = 0;
  if(rows[0] && rows[0][0]){
    const h = rows[0][0].toString().toLowerCase();
    if(h.includes('role') || h.includes('key') || h.includes('name') || h.includes('contact')) start = 1;
  }
  for(let i = start; i < rows.length; i++){
    const r = rows[i];
    if(!r || !r[0]) continue;
    const role = (r[0]||'').toString().trim().toLowerCase();
    if(/executive secretary|exec(utive)? sec|ward executive/i.test(role)){
      const name = (r[1]||'').toString().trim();
      const contact = (r[2]||'').toString().trim();
      return { name, contact };
    }
  }
  return null;
}

function normalizeTel(contact){
  if(!contact) return null;
  const digits = contact.replace(/[^\d+]/g,'');
  if((digits.match(/\d/g)||[]).length >= 7) return digits;
  return null;
}

function updateAppointmentsInfoBox(exec){
  const info = document.getElementById('info-note');
  if(!info) return;
  if(exec && (exec.name || exec.contact)){
    const nameHtml = exec.name ? `<strong>${escapeHtml(exec.name)}</strong>` : '';
    const tel = normalizeTel(exec.contact);
    const phoneHtml = exec.contact ? (tel ? `<a href="tel:${tel}">${escapeHtml(exec.contact)}</a>` : `<span>${escapeHtml(exec.contact)}</span>`) : '';
    info.hidden = false;
    info.innerHTML = `Appointments with the Bishop can be scheduled with the Ward Executive Secretary: ${nameHtml} ${phoneHtml}`.trim();
  } else {
    // hide if no data
    info.hidden = true;
    info.innerHTML = '';
  }
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

/* ---------- share button + menu (Email, QR Code, Link/native share) ---------- */
function initShare(){
  const shareBtn = document.getElementById('share-btn');
  const menu = document.getElementById('share-menu');
  const emailBtn = document.getElementById('share-email');
  const qrBtn = document.getElementById('share-qr');
  const linkBtn = document.getElementById('share-link');
  const qrImg = document.getElementById('share-qr-img');
  const status = document.getElementById('share-status');

  if(!shareBtn || !menu) return;

  const pageUrl = () => (location.href || window.location.toString());
  const pageTitle = () => (document.title || '');

  const hideMenu = () => {
    menu.setAttribute('aria-hidden','true');
    shareBtn.setAttribute('aria-expanded','false');
    menu.setAttribute('aria-qr','false');
    menu.style.display = 'none';
    status.style.display = 'none';
  };
  const showMenu = () => {
    menu.setAttribute('aria-hidden','false');
    shareBtn.setAttribute('aria-expanded','true');
    menu.style.display = 'flex';
  };

  shareBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    const hidden = menu.getAttribute('aria-hidden') === 'true';
    if(hidden) showMenu(); else hideMenu();
  });

  // EMAIL handler (first item)
  emailBtn.addEventListener('click', (e) => {
    const url = pageUrl();
    const subject = pageTitle() || 'Link';
    const body = `${pageTitle()}\n\n${url}`;
    // Use mailto to populate subject & body
    window.location.href = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
    // keep menu open briefly so mobile clients have time to honor the navigation
    setTimeout(hideMenu, 500);
  });

  // QR CODE handler (second item)
  qrBtn.addEventListener('click', (e) => {
    const url = pageUrl();
    const qrSrc = `https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=${encodeURIComponent(url)}`;
    qrImg.src = qrSrc;
    menu.setAttribute('aria-qr', 'true');
    status.textContent = 'Scan this code to open the page';
    status.style.display = 'block';
  });

  // LINK handler (native share if available)
  linkBtn.addEventListener('click', async (e) => {
    const url = pageUrl();
    const title = pageTitle();
    // Try native share first (mobile)
    if(navigator.share){
      try{
        await navigator.share({ title: title, url: url });
        hideMenu();
        return;
      }catch(err){
        // user cancelled or failed — fall through to fallback
      }
    }

    // Fallback: copy to clipboard and show "Copied" status (desktop fallback)
    try {
      if(navigator.clipboard && navigator.clipboard.writeText){
        await navigator.clipboard.writeText(url);
        status.textContent = 'Link copied to clipboard';
        status.style.display = 'block';
        setTimeout(()=> status.style.display = 'none', 2000);
      } else {
        const ta = document.createElement('textarea');
        ta.value = url;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        ta.remove();
        status.textContent = 'Link copied to clipboard';
        status.style.display = 'block';
        setTimeout(()=> status.style.display = 'none', 2000);
      }
    } catch(err) {
      status.textContent = 'Unable to share link';
      status.style.display = 'block';
      setTimeout(()=> status.style.display = 'none', 2000);
    }
  });

  // click outside closes
  document.addEventListener('click', (ev) => {
    if(!menu || !shareBtn) return;
    const target = ev.target;
    if(target === shareBtn || shareBtn.contains(target) || menu.contains(target)) return;
    hideMenu();
  });

  // close on escape
  document.addEventListener('keydown', (ev) => {
    if(ev.key === 'Escape' || ev.key === 'Esc'){
      hideMenu();
    }
  });

  // hide when focus leaves the menu entirely
  menu.addEventListener('focusout', () => {
    setTimeout(() => {
      if(!menu.contains(document.activeElement)) hideMenu();
    }, 0);
  });
}

/* ---------- icon/image helpers ---------- */
function getAgendaIcon(type){
  if(type === "hymn") return `<img src="./icons/hymn.svg" class="agenda-icon" alt="">`;
  if(type === "speaker") return `<img src="./icons/speaker.svg" class="agenda-icon" alt="">`;
  if(type === "prayer") return `<img src="./icons/prayer.svg" class="agenda-icon" alt="">`;
  if(type === "music") return `<img src="./icons/musicnumber.svg" class="agenda-icon" alt="">`;
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

/* ---------- Announcements parsing + rendering ---------- */
function parseAnnouncements(admRows){
  // returns [{text: '...', url: '...'}, ...] or [] if none
  if(!Array.isArray(admRows)) return [];
  for(let i=0;i<admRows.length;i++){
    const row = (admRows[i] || []).map(c => (c||'').toString().trim().toLowerCase());
    if(!row[0]) continue;
    // find header row that includes "announcement" in first column
    if(row[0].includes('announcement') || row[0].includes('announcements')){
      // confirm second column is header-ish (optional)
      // collect following rows until blank first cell
      const out = [];
      for(let j = i + 1; j < admRows.length; j++){
        const r = admRows[j];
        if(!r || !r[0] || r[0].toString().trim() === '') break;
        const txt = (r[0]||'').toString().trim();
        const url = (r[1]||'').toString().trim();
        if(txt) out.push({ text: txt, url: url || null });
      }
      return out;
    }
  }
  return [];
}

function renderAnnouncements(admRows){
  // parse
  const announcements = parseAnnouncements(admRows);
  if(!announcements || announcements.length === 0) {
    // if none, remove any existing announcements section to keep UI tidy
    const old = document.getElementById('announcements-section');
    if(old && old.parentNode) old.parentNode.removeChild(old);
    return;
  }

  // remove existing to avoid duplicates
  const existing = document.getElementById('announcements-section');
  if(existing) existing.remove();

  // build section (match structure used by Activities / Ward Leadership)
  const section = document.createElement('section');
  section.className = 'card collapsible';
  section.id = 'announcements-section';

  // toggle button (match markup/style of other toggles)
  const toggle = document.createElement('button');
  toggle.className = 'collapsible-toggle';
  toggle.setAttribute('data-target','announcements-panel');
  toggle.setAttribute('aria-expanded','false');
  toggle.type = 'button';
  toggle.innerHTML = `
    <span>Announcements</span>
    <svg class="chev" viewBox="0 0 24 24" width="18" height="18" aria-hidden>
      <path fill="none" stroke="#fff" stroke-width="2" d="M6 9l6 6 6-6"/>
    </svg>
  `;
  section.appendChild(toggle);

  // panel
  const panel = document.createElement('div');
  panel.className = 'collapsible-panel';
  panel.id = 'announcements-panel';
  panel.style.display = 'none'; // keep closed by default (global wiring will toggle)
  panel.style.padding = '12px';

  // populate announcements inside the panel
  announcements.forEach((a, idx) => {
    const item = document.createElement('div');
    item.className = 'announcement-item';
    item.style.padding = '8px 0';
    item.style.display = 'flex';
    item.style.flexDirection = 'column';
    item.style.gap = '8px';

    const p = document.createElement('p');
    p.className = 'muted';
    p.style.margin = '0';
    p.style.color = '#0f1724';
    p.style.fontSize = '14px';
    p.textContent = a.text;
    item.appendChild(p);

    if (a.url) {
      const normalized = normalizeHref(a.url);
    
      const aWrap = document.createElement('div');
      aWrap.style.marginTop = '6px';
    
      const btn = document.createElement('a');
      btn.className = 'announcement-link';
      btn.href = normalized;
      btn.target = '_blank';
      btn.rel = 'noopener noreferrer';
      btn.title = normalized;
    
      btn.style.display = 'inline-flex';
      btn.style.alignItems = 'center';
      btn.style.gap = '8px';
      btn.style.padding = '8px 12px';
      btn.style.borderRadius = '8px';
      btn.style.background = 'var(--accent-2)';
      btn.style.color = 'var(--accent)';
      btn.style.fontWeight = '700';
      btn.style.textDecoration = 'none';
    
      // favicon
      const fav = document.createElement('img');
      try {
        const u = new URL(normalized);
        fav.src = `https://www.google.com/s2/favicons?sz=64&domain=${u.origin}`;
      } catch (e) {
        fav.src = './icons/link.svg';
      }
      fav.alt = '';
      fav.style.width = '18px';
      fav.style.height = '18px';
      fav.style.borderRadius = '4px';
      fav.style.background = '#fff';
      fav.style.padding = '2px';
      fav.style.flex = '0 0 auto';
      btn.appendChild(fav);
    
      // shortened display text
      const text = document.createElement('span');
      text.textContent = getDisplayUrl(normalized);
      text.style.whiteSpace = 'nowrap';
      text.style.overflow = 'hidden';
      text.style.textOverflow = 'ellipsis';
      text.style.maxWidth = 'calc(100% - 64px)';
      btn.appendChild(text);
    
      // arrow glyph
      const arrow = document.createElement('svg');
      arrow.className = 'hymn-arrow';
      arrow.setAttribute('viewBox', '0 0 24 24');
      arrow.innerHTML = '<path d="M9 6l6 6-6 6"/>';
      arrow.style.width = '18px';
      arrow.style.height = '18px';
      arrow.style.flex = '0 0 auto';
      btn.appendChild(arrow);
    
      aWrap.appendChild(btn);
      item.appendChild(aWrap);
    }

    panel.appendChild(item);

    if(idx < announcements.length - 1){
      const div = document.createElement('div');
      div.style.height = '1px';
      div.style.background = '#e5e7eb';
      div.style.margin = '8px 0';
      panel.appendChild(div);
    }
  });

  section.appendChild(panel);

  // Insert BEFORE the activities section (top of the three)
  const activities = document.getElementById('activities-section');
  if(activities && activities.parentNode){
    activities.parentNode.insertBefore(section, activities);
  } else {
    // fallback: append after program card
    const program = document.getElementById('program');
    if(program && program.parentNode){
      program.parentNode.insertBefore(section, program.nextSibling);
    } else {
      document.querySelector('.app').appendChild(section);
    }
  }

  // IMPORTANT:
  // Do NOT attach a local click handler here for the toggle — the global wiring in run()
  // (document.querySelectorAll('.collapsible-toggle')...) will attach the handler.
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
/* ---------- Header rendering + conference logic (replace previous renderHeaderFromAdmin) ---------- */

function findUtcForTimeZoneLocal(year, month, day, hour, minute = 0, timeZone = 'America/Denver'){
  // Return a Date object (UTC instant) that maps to the requested wall-clock in the given timeZone.
  function partsFor(date, tz){
    const fmt = new Intl.DateTimeFormat('en-US', {
      timeZone: tz,
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', hour12: false
    });
    const parts = {};
    for(const p of fmt.formatToParts(date)){
      if(p.type !== 'literal') parts[p.type] = p.value;
    }
    return parts;
  }

  const utcBase = Date.UTC(year, month - 1, day, hour, minute, 0);
  // search +/- 12 hours for the instant that matches the requested wall-clock in the timezone
  for(let offset = -12; offset <= 12; offset++){
    const cand = new Date(utcBase + offset * 3600 * 1000);
    const p = partsFor(cand, timeZone);
    const py = Number(p.year), pm = Number(p.month), pd = Number(p.day), ph = Number(p.hour), pmin = Number(p.minute);
    if(py === year && pm === month && pd === day && ph === hour && pmin === minute){
      return cand;
    }
  }
  return new Date(utcBase); // fallback
}

function formatLocalForInstant(date /* Date object */, locales){
  // short time and abbreviated weekday+month/day
  try{
    return new Intl.DateTimeFormat(locales || undefined, {
      weekday: 'short', month:'short', day:'numeric',
      hour:'numeric', minute:'2-digit'
    }).format(date);
  }catch(e){
    return date.toLocaleString();
  }
}

function renderGeneralConference(adminMap, admRows){
  const pc = document.getElementById('program-content');
  if(!pc) return;

  // build or replace GC wrapper (we use stake-wrapper styles for consistent look)
  // remove any existing GC wrapper first to avoid duplicates
  const existingGc = pc.querySelector('.gc-wrapper');
  if(existingGc) existingGc.remove();

  const wrapper = document.createElement('div');
  wrapper.className = 'stake-wrapper gc-wrapper';

  // Title
  const titleDiv = document.createElement('div');
  titleDiv.className = 'stake-title';
  titleDiv.textContent = 'GENERAL CONFERENCE';
  wrapper.appendChild(titleDiv);

  // Determine the saturday & sunday to display:
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate()); // midnight local
  const day = today.getDay(); // 0 = Sunday ... 6 = Saturday
  let satDate, sunDate;

  // If adminMap provided explicit keys use them (same logic as earlier)
  const gcSatKey = Object.keys(adminMap).find(k => k.includes('general') && k.includes('saturday') && k.includes('date'));
  const gcSunKey = Object.keys(adminMap).find(k => k.includes('general') && k.includes('sunday') && k.includes('date'));
  const gcAnyKey = Object.keys(adminMap).find(k => k.includes('general') && k.includes('date') && !k.includes('time'));

  function parseAdminDate(val){
    if(!val) return null;
    const d = new Date(val);
    if(!isNaN(d)) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    return null;
  }

  if(gcSatKey) satDate = parseAdminDate(adminMap[gcSatKey]);
  if(gcSunKey) sunDate = parseAdminDate(adminMap[gcSunKey]);
  if(!satDate && !sunDate && gcAnyKey){
    const base = parseAdminDate(adminMap[gcAnyKey]);
    if(base){
      // derive saturday and sunday around that base date
      const bDay = base.getDay();
      // move to nearest saturday of that week
      const saturday = new Date(base);
      saturday.setDate(base.getDate() + ((6 - bDay + 7) % 7));
      const sunday = new Date(saturday); sunday.setDate(saturday.getDate() + 1);
      satDate = saturday;
      sunDate = sunday;
    }
  }

  // If admin didn't give dates, compute the "conference weekend" as the upcoming Saturday & Sunday,
  // but special-case Sunday so that Sunday shows the Saturday from the same weekend (yesterday).
  if(!satDate || !sunDate){
    if(day === 0){ // Sunday -> show yesterday (sat) and today (sun)
      satDate = new Date(today); satDate.setDate(today.getDate() - 1);
      sunDate = new Date(today);
    } else {
      const daysUntilSat = (6 - day + 7) % 7; // 0..6 (if today saturday => 0)
      satDate = new Date(today); satDate.setDate(today.getDate() + daysUntilSat);
      sunDate = new Date(satDate); sunDate.setDate(satDate.getDate() + 1);
    }
  }

  // Build schedule table (no Saturday Evening session)
  const table = document.createElement('table');
  table.className = 'gc-schedule';
  table.style.width = '100%';
  table.style.borderCollapse = 'collapse';
  table.style.marginBottom = '8px';

  const thead = document.createElement('thead');
  const htr = document.createElement('tr');
  ['Date','Session','Time'].forEach(h=>{
    const th = document.createElement('th');
    th.textContent = h;
    th.style.textAlign = 'left';
    th.style.padding = '8px 6px';
    th.style.color = 'var(--muted)';
    th.style.fontSize = '13px';
    th.style.fontWeight = 700;
    htr.appendChild(th);
  });
  thead.appendChild(htr);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');

  const sessions = [
    { session: 'Saturday Morning Session', date: satDate, hour: 10 },
    { session: 'Saturday Afternoon Session', date: satDate, hour: 14 },
    { session: 'Sunday Morning Session', date: sunDate, hour: 10 },
    { session: 'Sunday Afternoon Session', date: sunDate, hour: 14 }
  ];

  const userLocales = navigator.languages && navigator.languages.length ? navigator.languages[0] : navigator.language || 'en-US';

  sessions.forEach(s => {
    const tr = document.createElement('tr');
    tr.style.borderBottom = '1px solid rgba(0,0,0,0.04)';

    const tdDate = document.createElement('td');
    tdDate.style.padding = '10px 6px';
    if(s.date && s.date instanceof Date){
      tdDate.textContent = new Intl.DateTimeFormat(userLocales, { weekday:'short', month:'short', day:'numeric' }).format(s.date);
    } else {
      tdDate.textContent = '';
    }

    const tdSession = document.createElement('td');
    tdSession.style.padding = '10px 6px';
    tdSession.textContent = s.session;

    const tdTime = document.createElement('td');
    tdTime.style.padding = '10px 6px';
    if(s.date){
      const y = s.date.getFullYear();
      const m = s.date.getMonth() + 1;
      const d = s.date.getDate();
      const instant = findUtcForTimeZoneLocal(y, m, d, s.hour, 0, 'America/Denver');
      const userTimeString = new Intl.DateTimeFormat(userLocales, { hour:'numeric', minute:'2-digit', timeZoneName:'short' }).format(instant);
      const mountainTimeString = new Intl.DateTimeFormat('en-US', { hour:'numeric', minute:'2-digit', timeZone:'America/Denver', timeZoneName:'short' }).format(instant);
      tdTime.textContent = `${userTimeString} (${mountainTimeString})`;
    } else {
      tdTime.textContent = '—';
    }

    tr.appendChild(tdDate);
    tr.appendChild(tdSession);
    tr.appendChild(tdTime);
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  wrapper.appendChild(table);

  // subtitle & watch cards (reuse same platform list as before)
  const subtitle = document.createElement('div');
  subtitle.style.fontWeight = 700;
  subtitle.style.color = 'var(--muted)';
  subtitle.style.margin = '6px 4px 6px 4px';
  subtitle.textContent = 'Where can I watch General Conference?';
  wrapper.appendChild(subtitle);

  const platforms = [
    { name: 'Church Broadcasts', url: 'https://www.churchofjesuschrist.org/media/broadcasts?lang=eng' },
    { name: 'General Conference', url: 'https://www.churchofjesuschrist.org/study/general-conference?lang=eng' },
    { name: 'Ways to Watch', url: 'https://www.churchofjesuschrist.org/learn/ways-to-watch-general-conference?lang=eng' },
    { name: 'YouTube', url: 'https://www.youtube.com/churchofjesuschrist' },
    { name: 'Gospel Stream / Apps', url: 'https://www.churchofjesuschrist.org/learn/gospel-stream-mobile-and-ott-app?lang=eng' },
    { name: 'BYUtv', url: 'https://www.byutv.org/live' },
    { name: 'Deseret News', url: 'https://www.deseret.com/faith/general-conference/' },
    { name: 'KSL', url: 'https://www.ksl.com/news/general-conference' },
    { name: 'LDS Living', url: 'https://www.ldsliving.com/tag/general-conference' },
    { name: 'Church News', url: 'https://www.thechurchnews.com/general-conference/' }
  ];

  const cardsContainer = document.createElement('div');
  cardsContainer.id = 'gc-watch-cards';
  cardsContainer.style.display = 'flex';
  cardsContainer.style.flexDirection = 'column';
  cardsContainer.style.gap = '12px';
  cardsContainer.style.marginTop = '6px';

  platforms.forEach(p => {
    const el = document.createElement('div');
    el.className = 'event-card';
    el.style.cursor = 'pointer';
    el.style.alignItems = 'center';

    const left = document.createElement('div');
    left.style.display = 'flex';
    left.style.gap = '12px';
    left.style.alignItems = 'center';

    const img = document.createElement('img');
    try{
      const urlObj = new URL(p.url);
      img.src = `https://www.google.com/s2/favicons?sz=64&domain=${urlObj.origin}`;
    }catch(e){
      img.src = './icons/link.svg';
    }
    img.alt = '';
    img.style.width = '36px';
    img.style.height = '36px';
    img.style.borderRadius = '8px';
    img.style.background = '#fff';
    img.style.padding = '6px';
    img.style.objectFit = 'contain';
    left.appendChild(img);

    const txt = document.createElement('div');
    txt.style.display = 'flex';
    txt.style.flexDirection = 'column';
    txt.style.gap = '4px';

    const nameEl = document.createElement('div');
    nameEl.style.fontWeight = 700;
    nameEl.style.color = '#0f1724';
    nameEl.textContent = p.name;
    txt.appendChild(nameEl);

    const urlEl = document.createElement('div');
    const a = document.createElement('a');
    a.href = p.url;
    a.target = '_blank';
    a.rel = 'noopener';
    a.className = 'gc-url';                // CSS class we'll style
    a.textContent = getDisplayUrl(p.url);  // shortened display
    a.title = p.url;                       // full URL on hover
    urlEl.appendChild(a);
    txt.appendChild(urlEl);

    left.appendChild(txt);

    const right = document.createElement('div');
    right.innerHTML = `<svg class="hymn-arrow" viewBox="0 0 24 24"><path d="M9 6l6 6-6 6"/></svg>`;
    right.style.marginLeft = '8px';

    el.appendChild(left);
    el.appendChild(right);

    el.addEventListener('click', (e) => {
      const a = e.target.closest('a');
      if(a) return;
      window.open(p.url, '_blank', 'noopener');
    });

    cardsContainer.appendChild(el);
  });

  wrapper.appendChild(cardsContainer);

  // Insert wrapper at top of program content (so title + table + subtitle + cards appear inside the program card)
  pc.insertBefore(wrapper, pc.firstChild);
}

/* New renderHeaderFromAdmin that calls stake/general handlers and also hides meta when appropriate */
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

  // --- new: Meeting Time line (from Admin "Meeting Time" key) ---
  const meetingTime = (map['meeting time'] || '').toString().trim();
  let meetingTimeEl = document.getElementById('meeting-time');
  if (!meetingTimeEl) {
    meetingTimeEl = document.createElement('div');
    meetingTimeEl.id = 'meeting-time';
    meetingTimeEl.className = 'meeting-time';
    // Insert before the visitor-line (if present) so it appears above the visitor CTA
    const heroText = document.querySelector('.hero-text');
    const visitorLine = document.getElementById('visitor-line');
    if (heroText) {
      if (visitorLine) heroText.insertBefore(meetingTimeEl, visitorLine);
      else heroText.appendChild(meetingTimeEl);
    }
  }
  if (meetingTime) {
    meetingTimeEl.textContent = `Weekly Meeting Time: ${meetingTime}`;
    meetingTimeEl.style.display = '';
  } else {
    meetingTimeEl.style.display = 'none';
  }

  const mtLower = meetingType.toLowerCase();
  const isStakeConference = mtLower.includes('stake conference') || mtLower.includes('stake meeting') || mtLower === 'stake conference' || mtLower === 'stake';
  const isGeneralConference = mtLower.includes('general conference') || mtLower === 'general conference' || mtLower.includes('general');

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

    // Conducting
    if(conducting){
      const celt = document.getElementById('conducting'); if(celt) celt.textContent = conducting;
      const cLine = document.getElementById('conducting-line'); if(cLine) cLine.style.display = '';
    } else {
      const cLine = document.getElementById('conducting-line'); if(cLine) cLine.style.display = 'none';
    }

    // Chorister & Organist
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

  // Stake conference behavior
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
        // Only remove any existing conference-events container to avoid duplicates.
        // Don't remove other wrappers (meeting-placeholder, stake-wrapper, etc.) — they may
        // contain other important content like Activities or Ward Leadership.
        const existingEvents = progContent.querySelector('#conference-events');
        if (existingEvents) existingEvents.remove();
  
        const placeholder = progContent.querySelector('.meeting-placeholder');
        if (placeholder) {
          placeholder.after(container);
        } else {
          progContent.appendChild(container);
        }
      } else if (program && program.parentNode) {
        program.parentNode.insertBefore(container, program.nextSibling);
      } else {
        document.querySelector('.app').appendChild(container);
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
  
  // General conference behavior (place schedule + watch cards inside the program card)
  if (isGeneralConference) {
    // Only remove an existing GC wrapper to avoid duplicates.
    // Do NOT touch other wrappers (meeting-placeholder / stake-wrapper) so Activities / Leadership stay.
    const pc = document.getElementById('program-content');
    if (pc) {
      const existingGC = pc.querySelector('.gc-wrapper');
      if (existingGC) existingGC.remove();
    }
  
    // call renderer (renderGeneralConference should create .gc-wrapper inside #program-content)
    renderGeneralConference(map, admRows);
  } else {
    // remove any general-conference wrapper if present
    const pc = document.getElementById('program-content');
    if (pc) {
      const existingGC = pc.querySelector('.gc-wrapper');
      if (existingGC) existingGC.remove();
    }
  }
  
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
  let announcementsCsvUrl = config.announcements_csv_url || (config.sheet_id && config.announcements_gid ? buildCsvUrl(config.sheet_id, config.announcements_gid) : null);

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
    if (announcementsCsvUrl) fetches.push(fetch(announcementsCsvUrl));
    const responses = await Promise.all(fetches);

    const admResp = responses[0];
    const agResp = responses[1];

    // derive indices for optional responses
    const hasLead = Boolean(leadershipCsvUrl);
    const hasAnn  = Boolean(announcementsCsvUrl);
    const leadResp = hasLead ? responses[2] : null;
    const annResp  = hasAnn  ? responses[2 + (hasLead ? 1 : 0)] : null;

    if(!admResp.ok) throw new Error('Admin sheet fetch failed: ' + admResp.status);
    if(!agResp.ok) throw new Error('Agenda sheet fetch failed: ' + agResp.status);
    if(leadResp && !leadResp.ok) throw new Error('Leadership sheet fetch failed: ' + leadResp.status);

    const admText = await admResp.text();
    const agText = await agResp.text();
    const leadText = leadResp ? await leadResp.text() : null;
    const annText  = annResp  ? await annResp.text()  : null;

    if(/<html|doctype html/i.test(admText.slice(0,200))) { showError('Admin sheet returned HTML (not public)'); return; }
    if(/<html|doctype html/i.test(agText.slice(0,200))) { showError('Agenda sheet returned HTML (not public)'); return; }
    if(leadText && /<html|doctype html/i.test(leadText.slice(0,200))) { showError('Leadership sheet returned HTML (not public)'); return; }
    if(annText && /<html|doctype html/i.test(annText.slice(0,200))) { showError('Announcements sheet returned HTML (not public)'); return; }

    const admRows = parseCSVtoRows(admText);
    const agRows = parseCSVtoRows(agText);
    const leadRows = leadText ? parseCSVtoRows(leadText) : null;
    const annRows  = annText  ? parseCSVtoRows(annText)  : null;

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
          container.appendChild(createRow('Musical Number', name, extra, 'music'));
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
    
      // find Exec Secretary and populate the appointments info line
      const exec = findExecutiveSecretaryFromRows(leadRows);
      updateAppointmentsInfoBox(exec);
    } else {
      const ll = document.getElementById('leaders-list');
      if (ll) ll.innerHTML = '<div class="muted small">No leadership data found.</div>';
      // no leadership -> hide info
      updateAppointmentsInfoBox(null);
    }
    
    // Now render header and conference events AFTER the program content has been generated.
    renderHeaderFromAdmin(adminMap, admRows);

     // render announcements (top of the three sections)
    renderAnnouncements(annRows || admRows);

    // Move visitor into a bottom hero controls container, ensure share button + menu exist
function moveHeroControlsAndEnsureShare() {
  const heroInner = document.querySelector('.hero-inner') || document.querySelector('.hero');
  if (!heroInner) return;

  // create controls container if missing
  let controls = document.querySelector('.hero-controls');
  if (!controls) {
    controls = document.createElement('div');
    controls.className = 'hero-controls';
    heroInner.appendChild(controls);
  }

  // move visitor-line into controls (if present)
  const visitor = document.getElementById('visitor-line');
  if (visitor && !controls.contains(visitor)) {
    visitor.style.marginTop = '0';
    controls.appendChild(visitor);
  }

  // Ensure the share button exists and has the id/class initShare expects
  let shareBtn = document.getElementById('share-btn') || document.querySelector('.share-btn');
  if (!shareBtn) {
    shareBtn = document.createElement('button');
    shareBtn.id = 'share-btn';        // matches initShare() expectations
    shareBtn.className = 'share-btn';
    shareBtn.type = 'button';
    // Use your share.svg in images/
    shareBtn.innerHTML = '<img src="images/share.svg" alt="Share">';
    controls.appendChild(shareBtn);
  } else {
    // ensure it has the expected id/class and is inside controls
    if (!shareBtn.id) shareBtn.id = 'share-btn';
    if (!shareBtn.classList.contains('share-btn')) shareBtn.classList.add('share-btn');
    if (!controls.contains(shareBtn)) controls.appendChild(shareBtn);
    shareBtn.style.background = 'transparent';
    shareBtn.style.border = 'none';
    shareBtn.style.padding = '0';
  }

  // Create share-menu DOM if missing (so initShare() finds #share-menu, #share-email, #share-qr, #share-link, #share-qr-img, #share-status)
  if (!document.getElementById('share-menu')) {
    const menu = document.createElement('div');
    menu.id = 'share-menu';
    menu.className = 'share-menu';
    menu.setAttribute('aria-hidden', 'true');
    menu.setAttribute('aria-qr', 'false');

    menu.innerHTML = `
      <div id="share-email" class="share-action" role="button" tabindex="0">
        <img src="icons/email.svg" alt="" aria-hidden>
        <div>Email</div>
      </div>

      <div id="share-qr" class="share-action" role="button" tabindex="0">
        <img src="icons/qrcode.svg" alt="" aria-hidden>
        <div>QR Code</div>
      </div>

      <div id="share-link" class="share-action" role="button" tabindex="0">
        <img src="icons/link.svg" alt="" aria-hidden>
        <div>Link</div>
      </div>

      <div class="qr-preview" aria-hidden="true">
        <img id="share-qr-img" src="" alt="QR code">
      </div>

      <div id="share-status" class="muted-small" style="display:none"></div>
    `;

    // Append the menu to heroInner so its absolute positioning relative to hero works.
    heroInner.appendChild(menu);
  } else {
    // ensure menu is a child of heroInner (so absolute position works) and has default aria attrs
    const existingMenu = document.getElementById('share-menu');
    if (!heroInner.contains(existingMenu)) heroInner.appendChild(existingMenu);
    existingMenu.setAttribute('aria-hidden', existingMenu.getAttribute('aria-hidden') || 'true');
    existingMenu.setAttribute('aria-qr', 'false');
  }
}
    
    initShare();

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
