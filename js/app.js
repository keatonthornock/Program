// js/app.js
// Enhanced: load Administrative + Agenda CSVs, render agenda items, and handle conference events
const $ = s => document.querySelector(s);
const cfgPath = './config.json';

async function loadConfig(){
  try {
    const r = await fetch(cfgPath);
    if(!r.ok) throw new Error(`Missing ${cfgPath}. Edit this file in your repository and add your Google Sheets settings.`);
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

function sleep(ms){
  return new Promise(resolve => setTimeout(resolve, ms));
}

// Fetch CSV with timeout + small retry budget to reduce transient first-load failures.
// Uses no-store and a cache-busting param per attempt to avoid stale cached responses.
async function fetchCsvText(url, label, { timeoutMs = 8000, retries = 2 } = {}){
  let lastErr = null;
  for(let attempt = 0; attempt <= retries; attempt++){
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const u = new URL(url, window.location.href);
      u.searchParams.set('_cb', `${Date.now()}-${attempt}`);

      const resp = await fetch(u.toString(), {
        cache: 'no-store',
        signal: controller.signal
      });
      if(!resp.ok) throw new Error(`${label} sheet fetch failed: HTTP ${resp.status}`);

      const text = await resp.text();
      if(/<html|doctype html/i.test((text || '').slice(0, 200))){
        throw new Error(`${label} sheet returned HTML instead of CSV`);
      }

      clearTimeout(timeoutId);
      return text;
    } catch(err){
      clearTimeout(timeoutId);
      lastErr = err;
      const isLastAttempt = attempt >= retries;
      if(isLastAttempt) break;
      await sleep(300 * (attempt + 1));
    }
  }

  throw new Error(`${label} sheet failed after ${retries + 1} attempt(s): ${lastErr ? lastErr.message : 'Unknown error'}`);
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
    .replace(/^[0-9]{1,4}[a-z]?\s*\.\s*/i,'')
    .replace(/[’'"\.:,;!?\(\)\[\]\/]/g,'')
    .replace(/[^a-zA-Z0-9\s-]/g,'')
    .toLowerCase()
    .replace(/\s+/g,'-')
    .replace(/-+/g,'-')
    .replace(/^-|-$/g,'');
}

function normalizeHymnCollection(extraInfo){
  const raw = (extraInfo || '').toString();
  const normalized = raw
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/[—–-]/g, ' ')
    .replace(/&/g, ' and ')
    .replace(/\s+/g, ' ')
    .trim();

  if(!normalized) return null;

  if(
    normalized.includes("children's songbook") ||
    normalized.includes('childrens songbook') ||
    normalized.includes('children songbook') ||
    normalized.includes('children song book') ||
    (normalized.includes('child') && normalized.includes('songbook'))
  ){
    return 'childrens_songbook';
  }

  if(
    normalized.includes('hymns for home and church') ||
    normalized.includes('hymns for homes and church') ||
    normalized.includes('for home and church')
  ){
    return 'hymns_for_home_and_church';
  }

  if(normalized === 'hymns' || normalized.includes(' hymn')){
    return 'hymns';
  }

  return null;
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
  const rawOverride = (slugOverride || '').toString().trim();
  const hymnId = (hymnNumber || '').toString().trim().toLowerCase();
  const t = (title || '').toString().trim();
  const collection = normalizeHymnCollection(extraInfo);
  const standardHymnDuplicateTitleUrlMap = {
    '173': 'https://www.churchofjesuschrist.org/media/music/songs/while-of-these-emblems-we-partake-saul?crumbs=hymns&lang=eng',
    '174': 'https://www.churchofjesuschrist.org/media/music/songs/while-of-these-emblems-we-partake-aeolian?crumbs=hymns&lang=eng',
    '176': 'https://www.churchofjesuschrist.org/media/music/songs/tis-sweet-to-sing-the-matchless-love-meredith?crumbs=hymns&lang=eng',
    '177': 'https://www.churchofjesuschrist.org/media/music/songs/tis-sweet-to-sing-the-matchless-love-hancock?crumbs=hymns&lang=eng'
  };
  const exceptionUrl = collection === 'hymns' ? standardHymnDuplicateTitleUrlMap[hymnId] : null;
  const safeSlug = rawOverride ? slugify(rawOverride) : slugify(title);
  const searchQuery = `${hymnId ? `${hymnId} ` : ''}${t}`.trim();
  const idMatch = hymnId.match(/^([0-9]{1,4})([a-z]?)$/i);
  const numericPart = idMatch ? parseInt(idMatch[1], 10) : null;
  const hasLetterSuffix = idMatch ? Boolean(idMatch[2]) : false;

  if(/^(https?:)?\/\//i.test(rawOverride)){
    const fullUrl = normalizeHref(rawOverride);
    console.log('[hymn-links] using full URL override', { fullUrl });
    return fullUrl;
  }

  if(rawOverride.startsWith('/')){
    const relativeUrl = `https://www.churchofjesuschrist.org${rawOverride}`;
    console.log('[hymn-links] using relative URL override', { relativeUrl });
    return relativeUrl;
  }

  if(collection === 'childrens_songbook'){
    if(safeSlug){
      const url = `https://www.churchofjesuschrist.org/study/manual/childrens-songbook/${safeSlug}?lang=eng`;
      console.log('[hymn-links] using direct slug-generated route', { collection, url });
      return url;
    }
    if(idMatch){
      const url = `https://www.churchofjesuschrist.org/study/manual/childrens-songbook/${hymnId}?lang=eng`;
      console.log('[hymn-links] using number-based fallback', { collection, hymnId, url });
      return url;
    }
  }

  if(collection === 'hymns_for_home_and_church'){
    if(safeSlug){
      const url = `https://www.churchofjesuschrist.org/study/music/hymns-for-home-and-church/${safeSlug}?lang=eng`;
      console.log('[hymn-links] using direct slug-generated route', { collection, url });
      return url;
    }
    if(typeof numericPart === 'number' && numericPart >= 1000){
      const url = `https://www.churchofjesuschrist.org/study/music/hymns-for-home-and-church?lang=eng#${numericPart}`;
      console.log('[hymn-links] using number-based fallback', { collection, hymnId, url });
      return url;
    }
  }

  if(collection === 'hymns' || !collection){
    if(exceptionUrl){
      console.log('[hymn-links] using duplicate-title exception URL', { collection: collection || 'hymns', hymnId, exceptionUrl });
      return exceptionUrl;
    }
    if(safeSlug){
      const url = `https://www.churchofjesuschrist.org/study/manual/hymns/${safeSlug}?lang=eng`;
      console.log('[hymn-links] using direct slug-generated route', { collection: collection || 'hymns', url });
      return url;
    }
    if(typeof numericPart === 'number' && numericPart <= 341 && !hasLetterSuffix){
      const url = `https://www.churchofjesuschrist.org/study/manual/hymns/${numericPart}?lang=eng`;
      console.log('[hymn-links] using number-based fallback', { collection: collection || 'hymns', hymnId, url });
      return url;
    }
  }

  if(searchQuery){
    const searchUrl = `https://www.churchofjesuschrist.org/search?q=${encodeURIComponent(searchQuery)}`;
    console.log('[hymn-links] using search fallback', { searchUrl });
    return searchUrl;
  }

  console.log('[hymn-links] unable to generate hymn URL', { title, hymnNumber, extraInfo });
  return null;
}
function parseHymnName(name){
  const raw = (name || '').toString().trim();
  const match = raw.match(/^([0-9]{1,4}[a-z]?)\s*\.\s*(.+)$/i);
  if(!match) return null;
  return {
    hymnNumber: (match[1] || '').toLowerCase(),
    hymnTitle: (match[2] || '').trim()
  };
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
  const n = $('#info-note');
  if(n) {
    n.hidden = false;
    n.textContent = msg;
    n.dataset.mode = 'error';
  } else {
    console.warn(msg);
  }
}
function clearError(){
  const n = $('#info-note');
  if(n && n.dataset.mode === 'error') {
    n.hidden = true;
    n.textContent = '';
    delete n.dataset.mode;
  }
}

function normalizeItemKey(s){ return (s||'').toString().trim().toLowerCase(); }

function getSpecialMeetingTitleFromAgenda(agRows){
  if(!Array.isArray(agRows)) return '';
  for(const r of agRows){
    if(!r || !r[0]) continue;
    const itemKey = normalizeItemKey(r[0]);
    if(itemKey !== 'special meeting') continue;
    return (r[1] || '').toString().trim();
  }
  return '';
}

function shouldSkipAgendaItemForSpecialMeeting(key){
  if(!key) return false;

  // Always keep core opening/sacrament/closing structure.
  if(
    key.includes('opening hymn') ||
    key.includes('invocation') ||
    key.includes('opening prayer') ||
    key.includes('sacrament hymn') ||
    key.includes('administration of the sacrament') ||
    key.includes('closing hymn') ||
    key.includes('closing prayer') ||
    key.includes('benediction')
  ){
    return false;
  }

  // Suppress middle-program/body items by normalized type, regardless of data in other columns.
  if(/^speaker(\s+\d+)?$/.test(key) || key.startsWith('speaker')) return true;
  if(key.includes('musical')) return true;
  if(key.includes('intermediate hymn')) return true;
  if(key === 'testimony' || key.includes('testimon')) return true;
  if(key.includes('ward business') || key.includes('business')) return true;
  if(key.includes('release') || key.includes('sustain')) return true;

  return false;
}

function isSpecialMeetingResumeItem(key){
  if(!key) return false;
  return (
    key.includes('closing hymn') ||
    key.includes('closing prayer') ||
    key.includes('benediction')
  );
}


function updateFooterWardWebsite(wardName, wardWebsiteRaw){
  const footerSiteEl = document.getElementById('footer-ward-site');
  const sideSiteEl = document.getElementById('side-ward-homepage');
  if(!footerSiteEl) return;

  const website = (wardWebsiteRaw || '').toString().trim();
  if(!website){
    footerSiteEl.hidden = true;
    footerSiteEl.innerHTML = '';
    if(sideSiteEl){
      sideSiteEl.hidden = true;
      sideSiteEl.removeAttribute('href');
    }
    return;
  }

  const safeWardName = escapeHtml((wardName || 'Ward').toString().trim() || 'Ward');
  const href = normalizeHref(website);
  footerSiteEl.hidden = false;
  footerSiteEl.innerHTML = `<a href="${escapeHtml(href)}" target="_blank" rel="noopener">${safeWardName} Homepage</a>`;
  if(sideSiteEl){
    sideSiteEl.hidden = false;
    sideSiteEl.href = href;
    sideSiteEl.textContent = `${(wardName || 'Ward').toString().trim() || 'Ward'} Homepage`;
  }
}


function parseWardLogoName(wardName){
  const normalized = (wardName || '').toString().replace(/\s+/g, ' ').trim();
  if(!normalized) return { mainLine: '', subLine: '' };

  const suffixPatterns = [
    /\b(\d{1,2}(?:st|nd|rd|th)\s+ward)$/i,
    /\b(YSA\s+Ward)$/i,
    /\b(YSA\s+Branch)$/i,
    /\b(Ward)$/i,
    /\b(Branch)$/i
  ];

  for(const pattern of suffixPatterns){
    const match = normalized.match(pattern);
    if(!match) continue;

    const subLine = match[1].replace(/\s+/g, ' ').trim();
    const mainLine = normalized.slice(0, match.index).trim();
    if(mainLine.length >= 2){
      return { mainLine, subLine };
    }
  }

  return { mainLine: normalized, subLine: '' };
}

function renderWardTextLogo(container, wardName){
  if(!container) return;

  const { mainLine, subLine } = parseWardLogoName(wardName);
  const fallbackMain = mainLine || (wardName || '').toString().trim() || 'Ward Program';

  container.innerHTML = '';

  const mainEl = document.createElement('span');
  mainEl.className = 'ward-text-logo__main';
  mainEl.textContent = fallbackMain;
  container.appendChild(mainEl);

  if(subLine){
    const subEl = document.createElement('span');
    subEl.className = 'ward-text-logo__sub';
    subEl.textContent = subLine;
    container.appendChild(subEl);
  }
}

function renderWardTextLogos(wardName){
  renderWardTextLogo(document.getElementById('church-logo'), wardName);
  renderWardTextLogo(document.getElementById('side-ward-logo'), wardName);
}
function shouldRenderAgendaItem(key, meetingType){

  const isTestimony = meetingType.includes('testimony');
  const isSpecialMeeting = meetingType === 'special meeting';
  const isSacrament = meetingType.includes('sacrament') || meetingType === '' || meetingType === 'sacrament meeting' || isSpecialMeeting;

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

function shouldInsertLineDivider(currentKey, previousKey){
  if(!currentKey) return false;

  const isSacramentBoundary = k => /administration of the sacrament/i.test(k || '');
  if(isSacramentBoundary(currentKey) || isSacramentBoundary(previousKey)) return false;

  const isInvocation = /invocation|opening prayer/i.test(previousKey || '');
  const isIntermediateHymn = /intermediate hymn/i.test(currentKey);
  const isClosingHymn = /closing hymn/i.test(currentKey);

  return isInvocation || isIntermediateHymn || isClosingHymn;
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
    qrImg.src = '';
    menu.style.display = 'none';
    status.style.display = 'none';
  };
  const showMenu = () => {
    menu.setAttribute('aria-hidden','false');
    shareBtn.setAttribute('aria-expanded','true');
    menu.setAttribute('aria-qr','false');
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
  if(type === "hymn") return `<img src="./icons/hymn.png" class="agenda-icon agenda-icon--image" alt="" aria-hidden="true">`;
  if(type === "speaker") return `<img src="./icons/speaker.png" class="agenda-icon agenda-icon--image" alt="" aria-hidden="true">`;
  if(type === "testimony") return `<img src="./icons/testimony.png" class="agenda-icon agenda-icon--image" alt="" aria-hidden="true">`;
  if(type === "prayer") return `<img src="./icons/prayer.png" class="agenda-icon agenda-icon--image" alt="" aria-hidden="true">`;
  if(type === "music") return `<img src="./icons/musicnumber.png" class="agenda-icon agenda-icon--image" alt="" aria-hidden="true">`;
  return "";
}

/* ---------- rendering helpers ---------- */
function createElemFromHTML(html){
  const div = document.createElement('div');
  div.innerHTML = html.trim();
  return div.firstChild;
}

function createAgendaText(typeLabel, value){
  return `
    <div class="agenda-textline">
      <span class="agenda-label">${typeLabel}</span>
      <span class="agenda-separator">—</span>
      <span class="agenda-value">${value || ''}</span>
    </div>
  `;
  }

function formatHymnDetails(title, hymnNumber, extraInfo){
  const parts = [];
  if(hymnNumber) parts.push(`<span>#${escapeHtml(hymnNumber)}</span>`);
  if(title) parts.push(`<span>${escapeHtml(title)}</span>`);

  if(extraInfo){
    const sourceText = extraInfo.toString().trim();
    const isMutedSource = /^(hymns|hymns for home and church|childrens songbook)$/i.test(sourceText);
    if(isMutedSource){
      parts.push(`<span class="hymn-source">${escapeHtml(sourceText)}</span>`);
    } else {
      parts.push(`<span>${escapeHtml(sourceText)}</span>`);
    }
  }

  return parts.length ? parts.join(' <span class="agenda-separator">·</span> ') : '';
}

function createHymnCard(title, hymnNumber, label='Opening Hymn', url=null, extraInfo=''){
  const el = document.createElement(url ? 'button' : 'div');
  el.className = `agenda-item hymn-card${url ? ' hymn-button' : ''}`;
  if(url){
    el.type = 'button';
    el.setAttribute('aria-label', `${label}: ${hymnNumber ? `hymn ${hymnNumber}` : ''}${title ? ` ${title}` : ''}`.trim());
    el.addEventListener('click', () => window.open(url, '_blank', 'noopener'));
  }

  const details = formatHymnDetails(title, hymnNumber, extraInfo);

  el.innerHTML = `
    <div class="icon">${getAgendaIcon("hymn")}</div>
    <div class="content hymn-content">
      ${createAgendaText(label, details)}
    </div>
    <div class="right">${url ? `<svg class="hymn-arrow" viewBox="0 0 24 24"><path d="M9 6l6 6-6 6"/></svg>` : ''}</div>
  `;
  return el;
}

function createRow(typeLabel, name, extra, iconType = 'default'){
  const el = document.createElement('div');
  el.className = 'agenda-item';
  const value = [name || '', extra || ''].filter(Boolean).join(' · ');
  el.innerHTML = `
    <div class="icon">${getAgendaIcon(iconType)}</div>
    <div class="content">
      ${createAgendaText(typeLabel, value)}
    </div>
    <div class="right"></div>
  `;
  return el;
}

function createSingleLineRow(text, iconType = 'default'){
  const el = document.createElement('div');
  el.className = 'agenda-item';
  el.innerHTML = `
    <div class="icon">${getAgendaIcon(iconType)}</div>
    <div class="content">
      <div class="agenda-textline"><span class="agenda-label">${text || ''}</span></div>
    </div>
    <div class="right"></div>
  `;
  return el;
}

function createDivider(label){
  const el = document.createElement('div');
  const isSacramentDivider = /administration of the sacrament/i.test(label || '');
  const sacramentDividerClasses = isSacramentDivider
    ? ' agenda-divider--sacrament agenda-divider--administration'
    : '';
  el.className = `agenda-divider${sacramentDividerClasses}`;
  el.innerHTML = `<div class="divider-text">${label}</div>`;
  return el;
}

function createLineDivider(){
  const el = document.createElement('div');
  el.className = 'agenda-divider agenda-divider--line';
  el.innerHTML = '<div class="divider-text" aria-hidden="true"></div>';
  return el;
}

function createSpecialMeetingTitleBlock(title){
  const safeTitle = (title || '').toString().trim();
  if(!safeTitle) return null;

  const el = document.createElement('div');
  el.className = 'agenda-divider agenda-divider--special-event special-meeting-block';
  el.innerHTML = `<div class="divider-text special-meeting-title">${escapeHtml(safeTitle)}</div>`;
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
      <path d="M7 10l5 5 5-5"/>
    </svg>
  `;
  section.appendChild(toggle);

  // panel
  const panel = document.createElement('div');
  panel.className = 'collapsible-panel';
  panel.id = 'announcements-panel';
  panel.style.display = 'none'; // keep closed by default (global wiring will toggle)

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
    p.style.color = '#374151';
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

/* ---------- Activities & Events (Calendar sheet) ---------- */
function asTruthy(value){
  const v = (value || '').toString().trim().toLowerCase();
  return v === 'true' || v === 'yes' || v === 'y' || v === '1';
}

function parseCalendarEvents(rows){
  if(!Array.isArray(rows) || !rows.length) return [];

  const normalizedRows = rows.map(r => (Array.isArray(r) ? r : []));
  const first = normalizedRows[0].map(c => (c || '').toString().trim().toLowerCase());
  const hasHeaderRow = first.includes('event id') && first.includes('show on site');

  let startIndex = 0;
  const index = {
    eventId: 0,
    showOnSite: 1,
    calendarType: -1,
    title: 2,
    start: 3,
    end: 4,
    allDay: 5,
    location: 6,
    description: 7,
    lastSynced: 8
  };

  if(hasHeaderRow){
    startIndex = 1;
    const findIndex = (...names) => first.findIndex(c => names.includes(c));
    const mapped = {
      eventId: findIndex('event id'),
      showOnSite: findIndex('show on site'),
      calendarType: findIndex('calendar'),
      title: findIndex('title'),
      start: findIndex('start'),
      end: findIndex('end'),
      allDay: findIndex('all day'),
      location: findIndex('location'),
      description: findIndex('description'),
      lastSynced: findIndex('last synced')
    };

    Object.keys(mapped).forEach((key) => {
      if(mapped[key] !== -1) index[key] = mapped[key];
    });
  }

  const out = [];
  for(let i = startIndex; i < normalizedRows.length; i++){
    const r = normalizedRows[i] || [];
    const showOnSite = asTruthy(r[index.showOnSite]);
    if(!showOnSite) continue;

    const eventCalendarType = (index.calendarType >= 0 ? (r[index.calendarType] || '') : '').toString().trim();

    const ev = {
      eventId: (r[index.eventId] || '').toString().trim(),
      calendarType: eventCalendarType,
      title: (r[index.title] || '').toString().trim(),
      start: (r[index.start] || '').toString().trim(),
      end: (r[index.end] || '').toString().trim(),
      allDay: asTruthy(r[index.allDay]),
      location: (r[index.location] || '').toString().trim(),
      description: (r[index.description] || '').toString().trim(),
      lastSynced: (r[index.lastSynced] || '').toString().trim()
    };

    if(ev.title || ev.start || ev.end || ev.location || ev.description) out.push(ev);
  }

  return out;
}

function formatEventDateRange(startRaw, endRaw, isAllDay){
  const start = startRaw ? new Date(startRaw) : null;
  const end = endRaw ? new Date(endRaw) : null;
  const startValid = start && !Number.isNaN(start.getTime());
  const endValid = end && !Number.isNaN(end.getTime());

  if(!startValid && !endValid){
    const fallback = [startRaw, endRaw].filter(Boolean).join(' - ').trim();
    return fallback || 'Date/time unavailable';
  }

  const dateFmt = new Intl.DateTimeFormat(undefined, { weekday: 'short', month: 'short', day: 'numeric' });
  const timeFmt = new Intl.DateTimeFormat(undefined, { hour: 'numeric', minute: '2-digit' });

  const startDate = startValid ? dateFmt.format(start) : '';
  const endDate = endValid ? dateFmt.format(end) : '';
  const sameDay = startValid && endValid && start.toDateString() === end.toDateString();

  if(isAllDay){
    if(sameDay || !endValid) return `${startDate} (All day)`;
    return `${startDate} - ${endDate} (All day)`;
  }

  const startTime = startValid ? timeFmt.format(start) : '';
  const endTime = endValid ? timeFmt.format(end) : '';
  if(sameDay){
    return `${startDate} · ${startTime}${endTime ? ` - ${endTime}` : ''}`;
  }

  const startText = startValid ? `${startDate}${startTime ? ` · ${startTime}` : ''}` : '';
  const endText = endValid ? `${endDate}${endTime ? ` · ${endTime}` : ''}` : '';
  return [startText, endText].filter(Boolean).join(' - ');
}

function extractUrls(text){
  const src = (text || '').toString();
  const urlRe = /((?:https?:\/\/|www\.)[^\s]+|(?:[a-z0-9-]+\.)+[a-z]{2,}(?:\/[^\s]*)?)/gi;
  const urls = [];
  const without = src.replace(urlRe, (m) => {
    const clean = m.replace(/[),.;!?]+$/, '');
    if(clean.includes('@')) return m;
    urls.push(clean);
    return '';
  }).replace(/\s{2,}/g, ' ').trim();
  return { text: without, urls };
}

function createDescriptionWithInlineLinks(rawText, options = {}){
  const source = (rawText || '').toString();
  if(!source.trim()) return '';

  const buttonLinks = options.buttonLinks !== false;
  const maxChars = Number.isFinite(options.maxChars) ? Math.max(0, options.maxChars) : Infinity;
  const urlRe = /((?:https?:\/\/|www\.)[^\s]+|(?:[a-z0-9-]+\.)+[a-z]{2,}(?:\/[^\s]*)?)/gi;

  let visibleCount = 0;
  let cursor = 0;
  let out = '';

  const appendText = (txt) => {
    if(!txt || visibleCount >= maxChars) return;
    const remaining = maxChars - visibleCount;
    const chunk = txt.slice(0, remaining);
    out += escapeHtml(chunk);
    visibleCount += chunk.length;
  };

  const appendEllipsis = () => {
    if(Number.isFinite(maxChars) && visibleCount >= maxChars) out += '…';
  };

  for(const m of source.matchAll(urlRe)){
    if(visibleCount >= maxChars) break;

    const full = m[0] || '';
    const matchIndex = m.index || 0;
    const cleaned = full.replace(/[),.;!?]+$/, '');
    const trailing = full.slice(cleaned.length);

    appendText(source.slice(cursor, matchIndex));
    if(visibleCount >= maxChars){
      appendEllipsis();
      break;
    }

    if(cleaned.includes('@')){
      appendText(full);
      if(visibleCount >= maxChars){
        appendEllipsis();
        break;
      }
    } else {
      const normalized = normalizeHref(cleaned);
      const label = getDisplayUrl(normalized, 34);
      if(buttonLinks){
        out += `<a class="activity-inline-link" href="${escapeHtml(normalized)}" target="_blank" rel="noopener noreferrer" title="${escapeHtml(normalized)}">${escapeHtml(label)}</a>`;
      } else {
        out += `<a class="activity-preview-link" href="${escapeHtml(normalized)}" target="_blank" rel="noopener noreferrer" title="${escapeHtml(normalized)}">${escapeHtml(label)}</a>`;
      }
      visibleCount += cleaned.length;
      appendText(trailing);
      if(visibleCount >= maxChars){
        appendEllipsis();
        break;
      }
    }

    cursor = matchIndex + full.length;
  }

  if(visibleCount < maxChars){
    appendText(source.slice(cursor));
    if(Number.isFinite(maxChars) && source.length > maxChars) appendEllipsis();
  }

  return out;
}

function createActivityCard(ev){
  const card = document.createElement('article');
  card.className = 'activity-card';

  const parsedDescription = extractUrls(ev.description || '');
  const descriptionText = parsedDescription.text;
  const descriptionHtml = createDescriptionWithInlineLinks(ev.description || '', { buttonLinks: true });
  const previewHtml = createDescriptionWithInlineLinks(ev.description || '', { buttonLinks: false, maxChars: 140 });
  const preview = descriptionText.length > 140 ? `${descriptionText.slice(0, 140).trim()}…` : descriptionText;
  const schedule = formatEventDateRange(ev.start, ev.end, ev.allDay);
  const mapHref = ev.location ? `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(ev.location)}` : '';
  const hasMore = (descriptionText && descriptionText.length > preview.length) || parsedDescription.urls.length > 0;

  card.setAttribute('aria-expanded', 'false');
  if(hasMore){
    card.setAttribute('role', 'button');
    card.setAttribute('tabindex', '0');
  }

  card.innerHTML = `
    <div class="activity-card-header">
      <div class="activity-card-title-row">
        <div class="activity-card-title">${escapeHtml(ev.title || 'Untitled event')}</div>
      </div>
      ${hasMore ? `<span class="activity-arrow-wrap" aria-hidden="true"><svg class="activity-arrow" viewBox="0 0 24 24"><path d="M9 7l5 5-5 5"/></svg></span>` : ''}
    </div>
    <div class="activity-card-body">
      <div class="activity-meta-line activity-card-time">
        <span class="activity-time-main"><img src="./icons/calendar.png" class="activity-meta-icon" alt="" aria-hidden="true"><span>${escapeHtml(schedule)}</span></span>
        ${ev.calendarType ? `<span class="activity-calendar-pill">${escapeHtml(ev.calendarType)}</span>` : ''}
      </div>
      ${ev.location ? `<a class="activity-meta-line activity-location-link" href="${mapHref}" target="_blank" rel="noopener"><img src="./icons/map.png" class="activity-meta-icon" alt="" aria-hidden="true"><span>${escapeHtml(ev.location)}</span></a>` : ''}
      ${preview ? `<div class="activity-card-preview">${previewHtml}</div>` : ''}
      <div class="activity-card-details" hidden>
        ${descriptionHtml ? `<p class="activity-detail-line">${descriptionHtml}</p>` : ''}
      </div>
      ${hasMore ? `<button class="activity-see-more" type="button">See more</button>` : ''}
    </div>
  `;

  const seeMore = card.querySelector('.activity-see-more');
  const details = card.querySelector('.activity-card-details');
  const header = card.querySelector('.activity-card-header');

  const setExpanded = (expanded) => {
    card.setAttribute('aria-expanded', expanded ? 'true' : 'false');
    if(details) details.hidden = !expanded;
    if(seeMore) seeMore.textContent = expanded ? 'See less' : 'See more';
  };

  const onToggle = () => {
    if(!details) return;
    const expanded = card.getAttribute('aria-expanded') === 'true';
    setExpanded(!expanded);
  };

  if(hasMore && header){
    header.addEventListener('click', (e) => {
      if(e.target.closest('a') || e.target.closest('.activity-see-more')) return;
      onToggle();
    });

    card.addEventListener('keydown', (e) => {
      if(e.key === 'Enter' || e.key === ' '){
        e.preventDefault();
        onToggle();
      }
    });
  }

  if(seeMore){
    seeMore.addEventListener('click', (e) => {
      e.stopPropagation();
      onToggle();
    });
  }

  return card;
}

function renderActivities(calendarRows){
  const list = document.getElementById('activities-list');
  const panel = document.getElementById('activities-panel');
  if(!list || !panel) return;

  list.innerHTML = '';
  panel.querySelectorAll('.activities-filter-wrap').forEach(el => el.remove());
  panel.querySelectorAll('.activities-placeholder, .muted.small').forEach(el => el.remove());

  const events = parseCalendarEvents(calendarRows);
  if(events.length){
    const calendarTypes = [...new Set(events.map(ev => (ev.calendarType || '').trim()).filter(Boolean))];
    const activeFilters = new Set();
    let showAll = true;

    const renderEventList = () => {
      list.innerHTML = '';
      const visibleEvents = showAll
        ? events
        : events.filter(ev => activeFilters.has((ev.calendarType || '').trim()));

      visibleEvents.forEach(ev => list.appendChild(createActivityCard(ev)));

      if(!visibleEvents.length){
        const empty = document.createElement('div');
        empty.className = 'small activities-placeholder';
        empty.textContent = 'No events found for selected calendar type.';
        list.appendChild(empty);
      }
    };

    if(calendarTypes.length){
      const filterWrap = document.createElement('div');
      filterWrap.className = 'activities-filter-wrap';

      const renderFilterButtons = () => {
        filterWrap.innerHTML = '';

        const allBtn = document.createElement('button');
        allBtn.type = 'button';
        allBtn.className = `activities-filter-pill${showAll ? ' is-active' : ''}`;
        allBtn.textContent = 'All';
        allBtn.addEventListener('click', () => {
          showAll = true;
          activeFilters.clear();
          renderFilterButtons();
          renderEventList();
        });
        filterWrap.appendChild(allBtn);

        calendarTypes.forEach((type) => {
          const isActive = activeFilters.has(type);
          const typeBtn = document.createElement('button');
          typeBtn.type = 'button';
          typeBtn.className = `activities-filter-pill${isActive ? ' is-active' : ''}`;
          typeBtn.textContent = type;
          typeBtn.addEventListener('click', () => {
            if(activeFilters.has(type)) activeFilters.delete(type);
            else activeFilters.add(type);

            if(activeFilters.size === 0){
              showAll = true;
            } else {
              showAll = false;
            }

            renderFilterButtons();
            renderEventList();
          });
          filterWrap.appendChild(typeBtn);
        });
      };

      renderFilterButtons();
      panel.insertBefore(filterWrap, list);
    }

    renderEventList();
    return;
  }

  const wardHomepageEl = document.getElementById('side-ward-homepage');
  const homepageHref = wardHomepageEl && !wardHomepageEl.hidden ? wardHomepageEl.getAttribute('href') : '';

  const placeholder = document.createElement('div');
  placeholder.className = 'small activities-placeholder';
  if(homepageHref){
    placeholder.innerHTML = `<a class="activities-home-link" href="${escapeHtml(homepageHref)}" target="_blank" rel="noopener">See ward homepage for calendar events</a>`;
  } else {
    placeholder.textContent = 'See ward homepage for calendar events';
  }
  panel.appendChild(placeholder);
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
  el.className = 'event-card conference-card hymn-button';
  const mapHref = ev.address ? `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(ev.address)}` : '';
  const locationInfo = [
    ev.location,
    ev.date || ev.time ? `${ev.date ? ev.date : ''}${ev.date && ev.time ? ' · ' : ''}${ev.time ? ev.time : ''}` : ''
  ].filter(Boolean).join(' <span class="agenda-separator">·</span> ');

  el.innerHTML = `
    <div class="gc-card-left">
      <div class="icon">
        <img src="./icons/map.png" class="event-icon" alt="" aria-hidden="true">
      </div>
      <div class="gc-card-text">
        <div class="event-title">${ev.event || ''}</div>
        ${locationInfo ? `<div class="agenda-textline conference-subtext">${locationInfo}</div>` : ''}
        ${ev.address ? `<div class="conference-subtext"><a href="${mapHref}" target="_blank" rel="noopener">${ev.address}</a></div>` : ''}
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
  const thead = document.createElement('thead');
  thead.innerHTML = `
    <tr>
      <th scope="col">Calling</th>
      <th scope="col">Name</th>
      <th scope="col">Contact</th>
    </tr>
  `;
  table.appendChild(thead);
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

function getConferenceWeekendDates(adminMap, conferenceType){
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const day = today.getDay();
  const keys = Object.keys(adminMap || {});
  const typeNeedle = (conferenceType || '').toLowerCase();

  function parseAdminDate(val){
    if(!val) return null;
    const d = new Date(val);
    if(!isNaN(d)) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    return null;
  }

  function findDateKey(dayName){
    const typed = keys.find(k => k.includes(typeNeedle) && k.includes(dayName) && k.includes('date'));
    if(typed) return typed;
    return keys.find(k => k.includes(dayName) && k.includes('date'));
  }

  let satDate = parseAdminDate(adminMap[findDateKey('saturday')]);
  let sunDate = parseAdminDate(adminMap[findDateKey('sunday')]);

  if(!satDate || !sunDate){
    if(day === 0){
      satDate = satDate || new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
      sunDate = sunDate || new Date(today);
    } else {
      const daysUntilSat = (6 - day + 7) % 7;
      const upcomingSat = new Date(today.getFullYear(), today.getMonth(), today.getDate() + daysUntilSat);
      satDate = satDate || upcomingSat;
      sunDate = sunDate || new Date(upcomingSat.getFullYear(), upcomingSat.getMonth(), upcomingSat.getDate() + 1);
    }
  }

  return { satDate, sunDate };
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
  
  const { satDate, sunDate } = getConferenceWeekendDates(adminMap, 'general');

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
  subtitle.className = 'gc-watch-title';
  subtitle.style.fontWeight = 700;
  subtitle.style.color = 'var(--muted)';
  subtitle.style.margin = '24px 4px 12px 4px';
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
  cardsContainer.style.marginTop = '0';

  platforms.forEach(p => {
    const el = document.createElement('div');
    el.className = 'event-card hymn-button';
    el.style.cursor = 'pointer';
    el.style.alignItems = 'center';

    const left = document.createElement('div');
    left.className = 'gc-card-left';

    const iconWrap = document.createElement('div');
    iconWrap.className = 'icon';

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
    left.appendChild(iconWrap);
    iconWrap.appendChild(img);

    const txt = document.createElement('div');
    txt.className = 'gc-card-text';

    const nameEl = document.createElement('div');
    nameEl.className = 'gc-card-name';
    nameEl.textContent = p.name;
    txt.appendChild(nameEl);

    const urlEl = document.createElement('div');
    urlEl.className = 'gc-url-wrap';
    const a = document.createElement('a');
    a.href = p.url;
    a.target = '_blank';
    a.rel = 'noopener';
    a.className = 'gc-url';
    a.textContent = getDisplayUrl(p.url, 38);  // shortened display
    a.title = p.url;                       // full URL on hover
    urlEl.appendChild(a);
    txt.appendChild(urlEl);

    left.appendChild(txt);

    const right = document.createElement('div');
    right.innerHTML = `<svg class="hymn-arrow" viewBox="0 0 24 24"><path d="M9 6l6 6-6 6"/></svg>`;
    right.style.marginLeft = '8px';

    el.appendChild(left);
    el.appendChild(right);

    el.addEventListener('click', () => {
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
  const wardDetails = [ward, stake].filter(Boolean).join(' · ');
  const wardWebsite = map['ward website (optional)'] || '';

  const headingMeetingType = normalizeItemKey(meetingTypeRaw) === 'special meeting'
    ? 'Sacrament Meeting'
    : (meetingType || 'Sacrament Meeting');
  $('#meeting-heading').textContent = headingMeetingType;
  $('#meeting-date').textContent = dateRaw ? new Date(dateRaw).toLocaleDateString(undefined, { weekday:'long', month:'long', day:'numeric', year:'numeric' }) : '';

  const wardDetailsEl = document.getElementById('ward-details');
  if (wardDetailsEl) {
    wardDetailsEl.textContent = wardDetails;
    wardDetailsEl.style.display = wardDetails ? '' : 'none';
  }

  updateFooterWardWebsite(ward, wardWebsite);
  renderWardTextLogos(ward);

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
  const mtLower = meetingType.toLowerCase();
  const isStakeConference = mtLower.includes('stake conference') || mtLower.includes('stake meeting') || mtLower === 'stake conference' || mtLower === 'stake';
  const isGeneralConference = mtLower.includes('general conference') || mtLower === 'general conference' || mtLower.includes('general');

  if (isStakeConference || isGeneralConference) {
    const { satDate, sunDate } = getConferenceWeekendDates(map, isGeneralConference ? 'general' : 'stake');
    const fmt = new Intl.DateTimeFormat(undefined, { weekday:'short', month:'short', day:'numeric', year:'numeric' });
    const satLabel = satDate ? `Saturday: ${fmt.format(satDate)}` : '';
    const sunLabel = sunDate ? `Sunday: ${fmt.format(sunDate)}` : '';
    $('#meeting-date').textContent = [satLabel, sunLabel].filter(Boolean).join(' • ');
    if (wardDetailsEl) wardDetailsEl.style.display = 'none';
    meetingTimeEl.style.display = 'none';
  } else if (meetingTime) {
    meetingTimeEl.textContent = meetingTime;
    meetingTimeEl.style.display = '';
    if (wardDetailsEl) wardDetailsEl.style.display = wardDetails ? '' : 'none';
  } else {
    meetingTimeEl.style.display = 'none';
    if (wardDetailsEl) wardDetailsEl.style.display = wardDetails ? '' : 'none';
  }

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
  let calendarCsvUrl = config.calendar_csv_url || (config.sheet_id && config.calendar_gid ? buildCsvUrl(config.sheet_id, config.calendar_gid) : null);
 
  if(!adminCsvUrl){
    showError('No admin CSV URL available. Set admin_gid or admin_csv_url in config.json');
    return;
  }
  if(!agendaCsvUrl){
    showError('No agenda CSV URL available. Set agenda_gid or agenda_csv_url in config.json');
    return;
  }

  try {
    // Required sheets: fail startup if either one cannot be loaded.
    const [admText, agText] = await Promise.all([
      fetchCsvText(adminCsvUrl, 'Admin'),
      fetchCsvText(agendaCsvUrl, 'Agenda')
    ]);

    // Optional sheets: startup continues even if one fails.
    const optionalJobs = [];
    if (leadershipCsvUrl) optionalJobs.push({ key: 'leadership', label: 'Leadership', url: leadershipCsvUrl });
    if (announcementsCsvUrl) optionalJobs.push({ key: 'announcements', label: 'Announcements', url: announcementsCsvUrl });
    if (calendarCsvUrl) optionalJobs.push({ key: 'calendar', label: 'Calendar', url: calendarCsvUrl });

    const optionalResults = await Promise.allSettled(
      optionalJobs.map(job => fetchCsvText(job.url, job.label))
    );

    let leadText = null;
    let annText = null;
    let calText = null;
    optionalResults.forEach((result, idx) => {
      const job = optionalJobs[idx];
      if(result.status === 'fulfilled'){
        if(job.key === 'leadership') leadText = result.value;
        if(job.key === 'announcements') annText = result.value;
        if(job.key === 'calendar') calText = result.value;
      } else {
        console.warn(`[app] Optional ${job.label} sheet failed to load:`, result.reason);
      }
    });

    const admRows = parseCSVtoRows(admText);
    const agRows = parseCSVtoRows(agText);
    const leadRows = leadText ? parseCSVtoRows(leadText) : null;
    const annRows  = annText  ? parseCSVtoRows(annText)  : null;
    const calRows  = calText  ? parseCSVtoRows(calText)  : null;

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

    const meetingType = normalizeItemKey(adminMap['meeting type']);
    const isSpecialMeeting = meetingType === 'special meeting';
    const specialMeetingTitle = getSpecialMeetingTitleFromAgenda(agRows);
    const isTestimony = meetingType.includes('testimony');
    const isSacrament = meetingType.includes('sacrament') || meetingType === '' || meetingType === 'sacrament meeting' || isSpecialMeeting;

    let previousRenderedKey = '';
    let suppressSpecialMeetingMiddleSection = false;
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
      const itemKey = normalizeItemKey(item);

      // Control/content row: never render as a standard agenda item.
      if(itemKey === 'special meeting'){
        continue;
      }

      // Administration divider special
      if(itemKey.includes('administration of the sacrament')){
        if(isSacrament || isTestimony){
          container.appendChild(createDivider(item));
          any = true;
          previousRenderedKey = itemKey;

          if(isSpecialMeeting){
            const specialMeetingBlock = createSpecialMeetingTitleBlock(specialMeetingTitle);
            if(specialMeetingBlock){
              container.appendChild(specialMeetingBlock);
              any = true;
            }
            suppressSpecialMeetingMiddleSection = true;
          }
        }
        if(isTestimony){
          container.appendChild(createSingleLineRow('Testimonies of the Congregation', 'testimony'));
          any = true;
          previousRenderedKey = 'testimonies of the congregation';
        }
        continue;
      }

      if(!name) continue;

      const key = itemKey;

      if(isSpecialMeeting && suppressSpecialMeetingMiddleSection){
        if(!isSpecialMeetingResumeItem(key)){
          continue;
        }
        suppressSpecialMeetingMiddleSection = false;
      }

      if(isSpecialMeeting && shouldSkipAgendaItemForSpecialMeeting(key)){
        continue;
      }

      if(!shouldRenderAgendaItem(key, meetingType)){
        continue;
      }

      const hymnWithoutCard = key.includes('hymn') && !parseHymnName(name);

      if(!hymnWithoutCard && shouldInsertLineDivider(key, previousRenderedKey)){
        container.appendChild(createLineDivider());
        any = true;
      }

      // hymn handling
      if(key.includes('hymn')){
        const parsedHymn = parseHymnName(name);
        if(!parsedHymn){
          container.appendChild(createRow(item, name, '', 'music'));
          any = true;
          previousRenderedKey = key;
          continue;
        }

        const hymnNumber = parsedHymn.hymnNumber;
        const hymnTitle = parsedHymn.hymnTitle;
        const hymnUrl = getHymnUrl(hymnTitle, hymnNumber, extra, slugOverride);
        container.appendChild(createHymnCard(hymnTitle, hymnNumber, item, hymnUrl, extra));
        any = true;
        previousRenderedKey = key;
        continue;
      }

      if(key.startsWith('speaker') || key === 'testimony' || key.includes('testimon')){
        container.appendChild(createRow('Speaker', name, '', 'speaker'));
        any = true;
        previousRenderedKey = key;
        continue;
      }

      if(key.includes('invocation') || key.includes('opening prayer') || key.includes('closing prayer') || key.includes('benediction') || key.includes('closing')){
        container.appendChild(createRow(item, name, '', 'prayer'));
        any = true;
        previousRenderedKey = key;
        continue;
      }

      if(key.includes('musical')){
        if(!isTestimony){
          container.appendChild(createRow('Musical Number', name, extra, 'music'));
          any = true;
          previousRenderedKey = key;
        }
        continue;
      }

      // fallback generic
      container.appendChild(createRow(item, name, extra));
      any = true;
      previousRenderedKey = key;
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
        } else if ((adminMap['meeting type'] || '').toString().toLowerCase().includes('general conference')) {
          // General Conference content is rendered by renderGeneralConference().
          // Suppress the fallback placeholder so "General Conference" does not appear twice.
          pc.innerHTML = '';
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

    // render Activities & Events from Calendar tab
    renderActivities(calRows || []);
    
    initShare();

    clearError();

    // wire collapsibles + menu section shortcuts
    initCollapsibles();
    initSideMenu();
    initPwaInstall();

  } catch(err){
    console.error('[app] error', err);
    showError('Failed to fetch sheets: ' + err.message);
  }
}



function setCollapsibleState(btn, expanded){
  if(!btn) return;
  const target = btn.getAttribute('data-target');
  const panel = target ? document.getElementById(target) : null;
  btn.setAttribute('aria-expanded', expanded ? 'true' : 'false');
  if(panel) panel.style.display = expanded ? 'block' : 'none';
}

function initCollapsibles(){
  document.querySelectorAll('.collapsible-toggle').forEach(btn => {
    if(btn.dataset.boundCollapsible === 'true') return;
    setCollapsibleState(btn, false);
    btn.addEventListener('click', () => {
      const expanded = btn.getAttribute('aria-expanded') === 'true';
      setCollapsibleState(btn, !expanded);
    });
    btn.dataset.boundCollapsible = 'true';
  });
}

function expandAndScrollToSection(sectionId){
  const section = document.getElementById(sectionId);
  if(!section) return;
  const toggle = section.querySelector('.collapsible-toggle');
  if(toggle) setCollapsibleState(toggle, true);
  section.scrollIntoView({ behavior: 'smooth', block: 'start' });
}


function detectInstallPlatform(){
  const ua = (navigator.userAgent || '').toLowerCase();
  const isIos = /iphone|ipad|ipod/.test(ua);
  const isAndroid = /android/.test(ua);
  return { isIos, isAndroid };
}

let deferredInstallPrompt = null;
let installPromptEventSeen = false;

function updateInstallStatus(msg){
  const statusEl = document.getElementById('side-install-status');
  if(!statusEl) return;
  if(msg){
    statusEl.hidden = false;
    statusEl.textContent = msg;
  } else {
    statusEl.hidden = true;
    statusEl.textContent = '';
  }
}

function isRunningStandalone(){
  return window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
}

async function handleInstallAppClick(){
  if(isRunningStandalone()){
    updateInstallStatus('This app is already installed on your device.');
    return;
  }

  if(deferredInstallPrompt){
    deferredInstallPrompt.prompt();
    const choice = await deferredInstallPrompt.userChoice;
    deferredInstallPrompt = null;
    if(choice && choice.outcome === 'accepted'){
      updateInstallStatus('Great! Follow your device prompts to finish installing.');
    } else {
      updateInstallStatus('Install cancelled. You can try again anytime.');
    }
    return;
  }

  const { isIos, isAndroid } = detectInstallPlatform();
  if(isIos){
    updateInstallStatus('On iPhone/iPad: tap Share, then choose “Add to Home Screen”.');
    return;
  }

  if(isAndroid && installPromptEventSeen){
    updateInstallStatus('Install is not available yet. Visit this page again after browsing for a bit, then tap Install.');
    return;
  }

  if(isAndroid){
    updateInstallStatus('If no install prompt appears, open your browser menu and tap “Install app” or “Add to Home screen”.');
  } else {
    updateInstallStatus('Use your browser menu to install this app to your device.');
  }
}

function initPwaInstall(){
  const installBtn = document.getElementById('side-install-app');
  if(!installBtn || installBtn.dataset.boundInstall === 'true') return;

  window.addEventListener('beforeinstallprompt', (event) => {
    installPromptEventSeen = true;
    event.preventDefault();
    deferredInstallPrompt = event;
    updateInstallStatus('Ready to install — tap the button above.');
  });

  window.addEventListener('appinstalled', () => {
    deferredInstallPrompt = null;
    updateInstallStatus('App installed successfully.');
  });

  const installSupportDetected = window.matchMedia('(display-mode: browser)').matches || detectInstallPlatform().isAndroid || detectInstallPlatform().isIos;
  if(!installSupportDetected){
    installBtn.disabled = true;
    installBtn.setAttribute('aria-disabled', 'true');
    updateInstallStatus('This browser does not support installing this app from the menu button.');
  }

  installBtn.addEventListener('click', () => {
    handleInstallAppClick().catch(() => {
      updateInstallStatus('Unable to start install right now. Please use your browser menu.');
    });
  });

  installBtn.dataset.boundInstall = 'true';
}

function initSideMenu(){
  const overlay = document.getElementById('menu-overlay');
  const openBtn = document.getElementById('menu-open');
  const closeBtn = document.getElementById('menu-close');
  const scrim = document.getElementById('menu-scrim');
  if(!overlay || !openBtn || !closeBtn || !scrim) return;
  if(overlay.dataset.boundMenu === 'true') return;

  const openMenu = () => {
    overlay.hidden = false;
    document.body.style.overflow = 'hidden';
  };
  const closeMenu = () => {
    overlay.hidden = true;
    document.body.style.overflow = '';
  };

  openBtn.addEventListener('click', openMenu);
  closeBtn.addEventListener('click', closeMenu);
  scrim.addEventListener('click', closeMenu);

  overlay.querySelectorAll('.side-menu-link').forEach(link => {
    link.addEventListener('click', () => {
      closeMenu();
      const id = link.getAttribute('data-section');
      if(id) expandAndScrollToSection(id);
    });
  });

  overlay.dataset.boundMenu = 'true';
}


function initMobileScrollStrip(){
  const threshold = 28;
  const getScrollY = () => Math.max(
    window.pageYOffset || 0,
    document.documentElement ? document.documentElement.scrollTop || 0 : 0,
    document.body ? document.body.scrollTop || 0 : 0
  );

  const update = () => {
    const isMobile = window.matchMedia('(max-width: 900px)').matches;
    const y = getScrollY();
    const atTop = y <= 2;
    document.body.classList.toggle('mobile-scrolled', isMobile && y > threshold && !atTop);
  };

  window.addEventListener('scroll', update, { passive: true });
  window.addEventListener('resize', update);
  window.addEventListener('orientationchange', update);
  window.addEventListener('touchend', update, { passive: true });
  window.addEventListener('pageshow', update);
  if (window.visualViewport) {
    window.visualViewport.addEventListener('resize', update);
    window.visualViewport.addEventListener('scroll', update);
  }

  update();
}


initMobileScrollStrip();
run();
