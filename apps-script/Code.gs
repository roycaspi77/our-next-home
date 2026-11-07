/**
 * Our Next Home: Google Apps Script backend, bound to a Google Sheet.
 *
 * A shared list of homes for sale (Facebook posts, Yad2 ads) with a status,
 * a call log and a chat per home. The screens are a static page on GitHub
 * Pages (APP_URL) that POSTs JSON to this web app. Each person opens a
 * personal link (APP_URL#u=<token>); the token identifies them.
 *
 * Analyze reads an ad (its link, pasted text and/or screenshots) with Google
 * Gemini and fills only the fields the ad states. Needs the script property
 * GEMINI_API_KEY (free key from aistudio.google.com). GEMINI_MODEL is optional.
 *
 * Sheets (created / upgraded by setup()):
 *   Homes — one row per home; columns from COLUMNS plus META
 *   Chats — chat messages per home
 *   Calls — call log per home
 *   Users — name, personal link token, what each person has read
 */

const APP_NAME = 'Our Next Home';
const APP_URL = 'https://roycaspi77.github.io/our-next-home/';
const TZ = Session.getScriptTimeZone();
const DEFAULT_MODEL = 'gemini-2.5-flash';

const STATUSES = ['Not seen', 'Seen', 'Not relevant'];

/**
 * The home's columns, in screen order. To add a column, add one line here,
 * save, run setup() once, and deploy a new version. The page needs no change.
 *   type: text | longtext | number | money | checkbox | select | date | phone | link
 *   icon: for a checkbox, shows a toggle with this emoji (e.g. ❤️) instead of a tick box
 *   ai:   a description for Analyze; columns without it are never analyzed
 *   min/max: numbers outside the range found by Analyze are dropped
 */
const COLUMNS = [
  { key: 'liked', label: 'Liked it', type: 'checkbox', icon: '❤️' },
  { key: 'link', label: 'Link', type: 'link' },
  { key: 'city', label: 'City', type: 'text', ai: 'City, town or moshav/kibbutz name, as written' },
  { key: 'street', label: 'Street', type: 'text', ai: 'Street name, with the house number only if written' },
  { key: 'plot_m2', label: 'Plot m²', type: 'number', min: 1, max: 100000, ai: 'Plot / land area in square meters (שטח מגרש)' },
  { key: 'built_m2', label: 'Built m²', type: 'number', min: 10, max: 5000, ai: 'Built area of the house in square meters (שטח בנוי)' },
  { key: 'rooms', label: 'Rooms', type: 'number', min: 0.5, max: 30, ai: 'Number of rooms (חדרים), may be a half like 5.5' },
  { key: 'basement', label: 'Basement', type: 'checkbox', ai: 'true only if the ad says there is a basement (מרתף); false only if it says there is none' },
  { key: 'description', label: 'Description', type: 'longtext', ai: 'A short summary of the ad (up to 3 sentences) in the language of the ad, only facts it states' },
  { key: 'price', label: 'Price ₪', type: 'money', min: 10000, max: 1000000000, ai: 'Asking price in shekels as a full number (2.5 מיליון = 2500000)' },
  { key: 'seller_name', label: 'Seller', type: 'text', ai: 'Name of the seller or contact person, as written' },
  { key: 'seller_phone', label: 'Phone', type: 'phone', ai: 'Phone number of the seller, as written' },
  { key: 'status', label: 'Status', type: 'select', options: STATUSES },
  { key: 'next_step', label: 'Next step', type: 'text' },
  { key: 'visit_date', label: 'Visit', type: 'date' },
];
const META = ['id', 'created_by', 'created_at', 'updated_by', 'updated_at'];

const HEADERS = {
  Homes: META.slice(0, 1).concat(COLUMNS.map(c => c.key), META.slice(1)),
  Chats: ['id', 'home_id', 'author', 'time', 'text'],
  Calls: ['id', 'home_id', 'author', 'date', 'note', 'time'],
  Users: ['name', 'token', 'last_read'],
};

// ============================================================================
// Setup (run from the Apps Script editor; safe to run again)
// ============================================================================

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEADERS).forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[name]);
      sh.setFrozenRows(1);
      return;
    }
    const existing = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const missing = HEADERS[name].filter(h => existing.indexOf(h) < 0);
    if (missing.length) sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  });
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);

  if (readAll_('Users').length === 0) {
    writeRow_('Users', { name: 'Roy', token: newToken_(), last_read: '{}' });
    writeRow_('Users', { name: 'Wife', token: newToken_(), last_read: '{}' });
  }
  if (!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')) {
    Logger.log('Analyze needs a Gemini key: Project Settings > Script Properties > GEMINI_API_KEY.');
  }
  Logger.log('Setup done. Rename the people in the Users sheet if you like, then run printLinks.');
}

/** Logs each person's personal link. */
function printLinks() {
  readAll_('Users').forEach(u => Logger.log(`${u.name}: ${APP_URL}#u=${u.token}`));
}

// ============================================================================
// Web app entry (JSON API)
// ============================================================================

// Only these can be called from the page; each takes the caller's token first.
const API = { getState, saveHomes, analyze, getChat, sendChat, addCall, deleteCall, markRead };

/** The page sends text/plain JSON {fn, args} so the browser skips a CORS preflight. */
function doPost(e) {
  let result;
  try {
    const request = JSON.parse(e.postData.contents);
    const fn = API.hasOwnProperty(request.fn) ? API[request.fn] : null;
    if (!fn) throw new Error('Unknown action.');
    result = { ok: true, data: fn.apply(null, request.args || []) };
  } catch (err) {
    result = { ok: false, error: err.message || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, app: APP_NAME }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================================
// Homes
// ============================================================================

function getState(token) {
  const me = auth_(token);
  const read = lastRead_(me);
  const unread = {};
  readAll_('Chats').forEach(m => {
    if (m.author !== me.name && ms_(m.time) > (read[m.home_id] || 0)) {
      unread[m.home_id] = (unread[m.home_id] || 0) + 1;
    }
  });
  const calls = {};
  readAll_('Calls').forEach(c => { calls[c.home_id] = (calls[c.home_id] || 0) + 1; });
  return {
    me: me.name,
    users: readAll_('Users').map(u => u.name),
    columns: COLUMNS.map(c => {
      const out = { key: c.key, label: c.label, type: c.type, analyzed: !!c.ai };
      if (c.options) out.options = c.options;
      if (c.icon) out.icon = c.icon;
      return out;
    }),
    homes: readAll_('Homes').map(homeView_),
    unread: unread,
    calls: calls,
    analyzeReady: !!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
  };
}

/**
 * changes = {
 *   adds:    [{field: value, ...}],
 *   edits:   [{id, fields: {key: {from, to}}}],   from = the value the page started with
 *   deletes: [{id, updated_at}],
 * }
 * A field is written if nobody changed it in the sheet since the page loaded it
 * (or it already holds the new value). Otherwise it's returned as a conflict and
 * left as it is. The page then lets the person choose and saves again.
 */
function saveHomes(token, changes) {
  const me = auth_(token);
  changes = changes || {};
  return withLock_(() => {
    const homes = readAll_('Homes');
    const byId = {};
    homes.forEach(h => { byId[h.id] = h; });
    const now = new Date();
    const conflicts = [];
    const added = {};

    (changes.adds || []).forEach(a => {
      const row = { id: uid_(), created_by: me.name, created_at: now, updated_by: me.name, updated_at: now };
      COLUMNS.forEach(c => { row[c.key] = toCell_(c, a[c.key]); });
      if (STATUSES.indexOf(row.status) < 0) row.status = STATUSES[0];
      writeRow_('Homes', row);
      if (a._tmp) added[a._tmp] = row.id;
    });

    (changes.edits || []).forEach(ed => {
      const row = byId[ed.id];
      if (!row) { conflicts.push({ id: ed.id, deleted: true }); return; }
      let changed = false;
      Object.keys(ed.fields || {}).forEach(key => {
        const col = COLUMNS.find(c => c.key === key);
        if (!col) return;
        const f = ed.fields[key];
        const current = norm_(cellOut_(col, row[key]));
        if (current === norm_(f.to)) return;
        if (current !== norm_(f.from)) {
          conflicts.push({ id: ed.id, key: key, mine: f.to, theirs: cellOut_(col, row[key]), by: row.updated_by });
          return;
        }
        row[key] = toCell_(col, f.to);
        changed = true;
      });
      if (changed) {
        row.updated_by = me.name;
        row.updated_at = now;
        writeRow_('Homes', row);
      }
    });

    const deleteIds = (changes.deletes || []).map(d => d.id).filter(id => byId[id]);
    if (deleteIds.length) {
      deleteRows_('Homes', h => deleteIds.indexOf(h.id) >= 0);
      deleteRows_('Chats', m => deleteIds.indexOf(m.home_id) >= 0);
      deleteRows_('Calls', c => deleteIds.indexOf(c.home_id) >= 0);
    }

    return { homes: readAll_('Homes').map(homeView_), conflicts: conflicts, added: added };
  });
}

function homeView_(h) {
  const out = { id: h.id, created_by: h.created_by, created_at: ms_(h.created_at),
                updated_by: h.updated_by, updated_at: ms_(h.updated_at) };
  COLUMNS.forEach(c => { out[c.key] = cellOut_(c, h[c.key]); });
  return out;
}

/** Page value -> sheet cell. */
function toCell_(col, v) {
  if (v === undefined || v === null || v === '') return col.type === 'checkbox' ? false : '';
  switch (col.type) {
    case 'checkbox': return v === true || String(v).toLowerCase() === 'true';
    case 'number': case 'money': {
      const n = Number(String(v).replace(/[,₪\s]/g, ''));
      return isNaN(n) ? '' : n;
    }
    case 'select': return col.options.indexOf(v) >= 0 ? v : col.options[0];
    case 'date': return /^\d{4}-\d{2}-\d{2}$/.test(String(v)) ? "'" + v : '';
    default: {
      const s = String(v).trim().slice(0, 5000);
      // Keep text text: no formulas, and "0521234567" must not become a number.
      return /^[=+\-@]|^[\d.\s-]+$/.test(s) ? "'" + s : s;
    }
  }
}

/** Sheet cell -> page value. */
function cellOut_(col, v) {
  if (col.type === 'checkbox') return v === true || String(v).toUpperCase() === 'TRUE';
  if (v === '' || v === null || v === undefined) return '';
  if (col.type === 'number' || col.type === 'money') return typeof v === 'number' ? v : (Number(v) || '');
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v);
}

function norm_(v) {
  if (v === null || v === undefined || v === false) return '';
  return String(v).trim();
}

// ============================================================================
// Analyze
// ============================================================================

/**
 * Reads an ad and returns only the fields it states.
 * images = [{mime, data}] with data in base64 (the page resizes them first).
 * Returns {fields, found: [keys], source, duplicate, needInput, note}.
 */
function analyze(token, link, text, images) {
  auth_(token);
  link = String(link || '').trim();
  text = String(text || '').trim().slice(0, 20000);
  images = (images || []).slice(0, 4).filter(im => im && /^image\//.test(im.mime) && im.data);

  const result = { fields: {}, found: [], source: [], duplicate: null, needInput: false, note: '' };
  if (link) {
    const key = normalizeLink(link);
    const dup = readAll_('Homes').find(h => h.link && normalizeLink(String(h.link)) === key);
    if (dup) result.duplicate = { id: dup.id, city: String(dup.city || ''), street: String(dup.street || '') };
  }

  let pageText = '';
  if (link) {
    const page = readLink_(link);
    pageText = page.text;
    if (page.text) result.source.push(page.kind);
    else result.note = page.note;
  }
  if (text) result.source.push('pasted text');
  if (images.length) result.source.push(images.length === 1 ? 'screenshot' : images.length + ' screenshots');

  if (!pageText && !text && !images.length) {
    result.needInput = true;
    if (!result.note) result.note = 'Paste the post text or a screenshot of it, then Analyze again.';
    return result;
  }

  const raw = askGemini_(pageText, text, images);
  const sourceText = images.length ? '' : pageText + '\n' + text;
  result.fields = cleanAnalysis(raw, sourceText);
  result.found = Object.keys(result.fields);
  if (!result.found.length) result.note = 'No details found in this ad. Fill them in by hand.';
  return result;
}

/** Fetches the ad. Returns {text, kind, note}; text is '' when the page can't be read. */
function readLink_(link) {
  if (!/^https?:\/\//i.test(link)) return { text: '', note: 'That doesn\'t look like a link.' };
  const isFacebook = /(^|\.)(facebook\.com|fb\.com|fb\.watch)$/i.test(hostOf(link));
  const notReadable = isFacebook
    ? 'Facebook doesn\'t let apps read posts. Paste the post text or a screenshot, then Analyze again.'
    : 'Couldn\'t read this link. Paste the ad text or a screenshot, then Analyze again.';
  let resp;
  try {
    resp = UrlFetchApp.fetch(link, {
      muteHttpExceptions: true, followRedirects: true,
      headers: {
        'User-Agent': isFacebook ? 'facebookexternalhit/1.1'
          : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36',
        'Accept-Language': 'he-IL,he;q=0.9,en;q=0.8',
      },
    });
  } catch (err) {
    return { text: '', note: notReadable };
  }
  if (resp.getResponseCode() !== 200) return { text: '', note: notReadable };
  const html = resp.getContentText();
  const text = isFacebook ? facebookText(html) : adText(html, link);
  return text ? { text: text, kind: isFacebook ? 'Facebook preview' : 'the ad page' } : { text: '', note: notReadable };
}

function askGemini_(pageText, text, images) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('Analyze isn\'t set up: add GEMINI_API_KEY in Apps Script > Project Settings > Script Properties.');
  const model = props.getProperty('GEMINI_MODEL') || DEFAULT_MODEL;

  const parts = [{ text: analyzePrompt() }];
  if (pageText) parts.push({ text: 'AD PAGE CONTENT:\n' + pageText });
  if (text) parts.push({ text: 'PASTED POST TEXT:\n' + text });
  images.forEach(im => parts.push({ inline_data: { mime_type: im.mime, data: im.data } }));

  const resp = UrlFetchApp.fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-goog-api-key': key },
      payload: JSON.stringify({
        contents: [{ role: 'user', parts: parts }],
        generationConfig: { temperature: 0, responseMimeType: 'application/json', responseSchema: analyzeSchema() },
      }),
    });
  const code = resp.getResponseCode();
  if (code === 404) throw new Error(`Gemini model "${model}" wasn't found. Set GEMINI_MODEL in Script Properties to a current model name.`);
  if (code === 429) throw new Error('Gemini\'s free quota is used up for now. Try again in a minute.');
  if (code !== 200) throw new Error('Gemini error ' + code + ': ' + resp.getContentText().slice(0, 200));
  const body = JSON.parse(resp.getContentText());
  const out = body.candidates && body.candidates[0] && body.candidates[0].content
    && body.candidates[0].content.parts && body.candidates[0].content.parts[0];
  if (!out || !out.text) return {};
  try { return JSON.parse(out.text); } catch (err) { return {}; }
}

// ---- Pure helpers (no Apps Script services; tested in Node by .dev/test_helpers.js) ----

function analyzeColumns() { return COLUMNS.filter(c => c.ai); }

function analyzePrompt() {
  return [
    'You read a real-estate ad for a home for sale in Israel (a Facebook post or a Yad2 ad, in Hebrew or English).',
    'It may come as page content, pasted text and/or screenshots. Fill the JSON fields from the ad only.',
    'Rules:',
    '- Copy only what the ad explicitly states. Never guess, estimate, infer, compute or look anything up.',
    '- If a field is not stated, return null. A blank field is always better than a wrong one.',
    '- Keep names, cities and streets exactly as written, in the ad\'s language.',
    '- Numbers are plain numbers without units or commas. Writing out a stated amount is fine (2.5 מיליון = 2500000).',
    '- If the content shows several ads (for example "similar ads"), use only the main ad.',
    '- Ignore the poster\'s profile name unless the ad says they are the contact person.',
    'Fields:',
  ].concat(analyzeColumns().map(c => `- ${c.key}: ${c.ai}`)).join('\n');
}

function analyzeSchema() {
  const props = {};
  analyzeColumns().forEach(c => {
    const type = c.type === 'checkbox' ? 'BOOLEAN' : (c.type === 'number' || c.type === 'money') ? 'NUMBER' : 'STRING';
    props[c.key] = { type: type, nullable: true };
  });
  return { type: 'OBJECT', properties: props };
}

/**
 * Keeps only plausible values. When the source is text (no screenshots),
 * names, places and the phone must also appear in that text.
 */
function cleanAnalysis(raw, sourceText) {
  const out = {};
  const src = squash(sourceText || '');
  const srcDigits = String(sourceText || '').replace(/\D/g, '');
  analyzeColumns().forEach(c => {
    let v = raw ? raw[c.key] : null;
    if (v === null || v === undefined || v === '') return;
    if (c.type === 'checkbox') {
      if (typeof v === 'boolean') out[c.key] = v;
      return;
    }
    if (c.type === 'number' || c.type === 'money') {
      const n = typeof v === 'number' ? v : Number(String(v).replace(/[,₪\s]/g, ''));
      if (!isFinite(n) || n < c.min || n > c.max) return;
      if (c.key === 'rooms' && Math.round(n * 2) !== n * 2) return;
      out[c.key] = n;
      return;
    }
    v = String(v).trim();
    if (!v) return;
    if (c.type === 'phone') {
      const phone = israeliPhone(v);
      if (!phone) return;
      if (src && srcDigits.indexOf(phone.replace(/\D/g, '').replace(/^972/, '0')) < 0
          && srcDigits.indexOf(phone.replace(/\D/g, '').slice(1)) < 0) return;
      out[c.key] = phone;
      return;
    }
    if (c.type === 'text' && src && src.indexOf(squash(v)) < 0) return;
    out[c.key] = v.slice(0, c.type === 'longtext' ? 1500 : 200);
  });
  return out;
}

/** "054 123-4567" / "+972541234567" -> "054-1234567"; null if it isn't an Israeli number. */
function israeliPhone(s) {
  let d = String(s).replace(/\D/g, '');
  if (d.indexOf('972') === 0) d = '0' + d.slice(3);
  if (/^05\d{8}$/.test(d)) return d.slice(0, 3) + '-' + d.slice(3);
  if (/^07\d{8}$/.test(d)) return d.slice(0, 3) + '-' + d.slice(3);
  if (/^0[2-489]\d{7}$/.test(d)) return d.slice(0, 2) + '-' + d.slice(2);
  return null;
}

function squash(s) {
  return String(s).toLowerCase().replace(/[֑-ׇ"'`׳״.,\-–()\s]/g, '');
}

function hostOf(link) {
  const m = String(link).match(/^https?:\/\/([^\/?#:]+)/i);
  return m ? m[1].toLowerCase().replace(/^www\.|^m\.|^web\./, '') : '';
}

/** The same ad gets the same key: no tracking parameters, no www/m., no trailing slash. */
function normalizeLink(link) {
  const s = String(link).trim();
  const m = s.match(/^https?:\/\/([^\/?#]+)([^?#]*)(\?[^#]*)?/i);
  if (!m) return s.toLowerCase();
  const host = m[1].toLowerCase().replace(/^(www|m|web|mbasic)\./, '');
  const path = m[2].replace(/\/+$/, '');
  const keep = (m[3] || '').replace(/^\?/, '').split('&').filter(p => {
    const k = p.split('=')[0].toLowerCase();
    return p && !/^(utm_|fbclid|mibextid|rdid|share_url|ref|refsrc|__cft__|__tn__|component|opened-from|spot|location)/.test(k);
  }).sort();
  return host + path + (keep.length ? '?' + keep.join('&') : '');
}

function decodeEntities(s) {
  return String(s)
    .replace(/&#(\d+);/g, (x, n) => String.fromCharCode(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (x, n) => String.fromCharCode(parseInt(n, 16)))
    .replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, '\'')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&');
}

function metaTag(html, name) {
  const re = new RegExp('<meta[^>]+(?:property|name)=["\']' + name.replace(/[:.]/g, '\\$&') + '["\'][^>]*>', 'i');
  const tag = (html.match(re) || [])[0];
  if (!tag) return '';
  const c = tag.match(/content=["']([^"']*)["']/i);
  return c ? decodeEntities(c[1]).trim() : '';
}

/** Text of a public Facebook post from its preview tags; '' if there is none. */
function facebookText(html) {
  const desc = metaTag(html, 'og:description');
  const title = metaTag(html, 'og:title');
  if (desc.length < 30 || /log in|התחבר|see posts, photos and more/i.test(desc)) return '';
  return (title ? title + '\n' : '') + desc;
}

function looksBlocked(html) {
  return /captcha|shieldsquare|perfdrive|are you for real|access denied|validate\.perfdrive/i.test(html.slice(0, 20000))
    && html.length < 200000;
}

/**
 * Text of an ad page. For Yad2 it uses the ad's own data from the page's
 * __NEXT_DATA__ (not the "similar ads" around it); otherwise the meta tags
 * and the visible text.
 */
function adText(html, link) {
  if (!html || looksBlocked(html)) return '';
  const parts = [];
  const nd = html.match(/<script[^>]+id=["']__NEXT_DATA__["'][^>]*>([\s\S]*?)<\/script>/i);
  if (nd) {
    try {
      const item = findAdObject(JSON.parse(nd[1]), adToken(link));
      if (item) parts.push(flatten(item).join('\n'));
    } catch (err) { /* fall through to the page text */ }
  }
  const title = metaTag(html, 'og:title') || ((html.match(/<title[^>]*>([\s\S]*?)<\/title>/i) || [])[1] || '');
  const desc = metaTag(html, 'og:description') || metaTag(html, 'description');
  if (title) parts.unshift('Title: ' + decodeEntities(title).trim());
  if (desc) parts.splice(1, 0, 'Description: ' + desc);
  if (!nd) {
    const body = decodeEntities(html
      .replace(/<(script|style|noscript|svg|head)[\s\S]*?<\/\1>/gi, ' ')
      .replace(/<[^>]+>/g, '\n'))
      .replace(/[ \t]+/g, ' ').replace(/\n\s*\n+/g, '\n').trim();
    if (body) parts.push(body.slice(0, 15000));
  }
  const text = parts.join('\n').trim();
  return text.length >= 40 ? text.slice(0, 30000) : '';
}

/** Yad2 links end in /item/<token>. */
function adToken(link) {
  const m = String(link).match(/\/item\/(?:[^\/?#]+\/)*([A-Za-z0-9]{5,})(?:[\/?#]|$)/);
  return m ? m[1] : '';
}

/** The object in the page data that holds this ad: its token, or the item query. */
function findAdObject(root, token) {
  let best = null;
  (function walk(node, depth) {
    if (best || !node || typeof node !== 'object' || depth > 30) return;
    if (token && !Array.isArray(node) && (node.token === token || node.adNumber === token || node.orderId === token)) {
      best = node;
      return;
    }
    Object.keys(node).forEach(k => walk(node[k], depth + 1));
  })(root, 0);
  if (best) return best;
  // Fallback: a react-query entry whose key starts with "item".
  const queries = root && root.props && root.props.pageProps && root.props.pageProps.dehydratedState
    && root.props.pageProps.dehydratedState.queries;
  const q = (queries || []).find(x => Array.isArray(x.queryKey) && /item/i.test(String(x.queryKey[0])));
  return q && q.state ? q.state.data : null;
}

/** {a:{b:1}} -> ["a.b: 1"], skipping images, links and ids. */
function flatten(obj) {
  const lines = [];
  (function walk(node, path, depth) {
    if (lines.length > 600 || depth > 12 || node === null || node === undefined) return;
    if (typeof node !== 'object') {
      const v = String(node);
      if (!v || v.length > 3000 || /^https?:|\.(jpe?g|png|webp)/i.test(v)) return;
      if (/(image|img|photo|url|src|icon|color|id$|^id|hash|token|analytics|tracking)/i.test(path.split('.').pop() || '')) return;
      lines.push(`${path}: ${v}`);
      return;
    }
    if (Array.isArray(node)) { node.slice(0, 30).forEach((x, i) => walk(x, path + '[' + i + ']', depth + 1)); return; }
    Object.keys(node).forEach(k => walk(node[k], path ? path + '.' + k : k, depth + 1));
  })(obj, '', 0);
  return lines;
}

// ============================================================================
// Chat and call log
// ============================================================================

function getChat(token, homeId) {
  const me = auth_(token);
  return chatView_(me, homeId);
}

function sendChat(token, homeId, text) {
  const me = auth_(token);
  text = String(text || '').trim().slice(0, 4000);
  if (!text) throw new Error('Write something first.');
  return withLock_(() => {
    requireHome_(homeId);
    writeRow_('Chats', { id: uid_(), home_id: homeId, author: me.name, time: new Date(), text: text });
    setRead_(me, homeId);
    return chatView_(me, homeId);
  });
}

function addCall(token, homeId, date, note) {
  const me = auth_(token);
  note = String(note || '').trim().slice(0, 2000);
  if (!note) throw new Error('Write what happened on the call.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date || ''))) date = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  return withLock_(() => {
    requireHome_(homeId);
    writeRow_('Calls', { id: uid_(), home_id: homeId, author: me.name, date: "'" + date, note: note, time: new Date() });
    return chatView_(me, homeId);
  });
}

function deleteCall(token, homeId, callId) {
  const me = auth_(token);
  return withLock_(() => {
    deleteRows_('Calls', c => c.id === callId && c.home_id === homeId);
    return chatView_(me, homeId);
  });
}

function markRead(token, homeId) {
  const me = auth_(token);
  withLock_(() => setRead_(me, homeId));
  return true;
}

function chatView_(me, homeId) {
  const messages = readAll_('Chats').filter(m => m.home_id === homeId)
    .map(m => ({ id: m.id, author: m.author, time: ms_(m.time), text: String(m.text) }))
    .sort((a, b) => a.time - b.time);
  const calls = readAll_('Calls').filter(c => c.home_id === homeId)
    .map(c => ({ id: c.id, author: c.author, date: c.date instanceof Date ? Utilities.formatDate(c.date, TZ, 'yyyy-MM-dd') : String(c.date),
                 note: String(c.note), time: ms_(c.time) }))
    .sort((a, b) => (a.date + a.time) < (b.date + b.time) ? -1 : 1);
  return { messages: messages, calls: calls, me: me.name };
}

function requireHome_(homeId) {
  if (!readAll_('Homes').some(h => h.id === homeId)) throw new Error('This home was deleted. Reload the page.');
}

function lastRead_(user) {
  try { return JSON.parse(user.last_read || '{}') || {}; } catch (err) { return {}; }
}

function setRead_(me, homeId) {
  const user = readAll_('Users').find(u => u.token === me.token);
  if (!user) return;
  const read = lastRead_(user);
  read[homeId] = Date.now();
  user.last_read = JSON.stringify(read);
  writeRow_('Users', user);
}

// ============================================================================
// Sheet + misc helpers
// ============================================================================

function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }

function readAll_(name) {
  const values = sheet_(name).getDataRange().getValues();
  const header = values.shift() || [];
  return values
    .map((row, i) => {
      const obj = { _row: i + 2 };
      header.forEach((key, j) => { obj[key] = row[j]; });
      return obj;
    })
    .filter(obj => obj[header[0]] !== '');
}

/** Writes by header name, so columns the code doesn't know are left as they are. */
function writeRow_(name, obj) {
  const sh = sheet_(name);
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (obj._row) {
    const current = sh.getRange(obj._row, 1, 1, header.length).getValues()[0];
    const row = header.map((key, j) => (key in obj ? cell_(obj[key]) : current[j]));
    sh.getRange(obj._row, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(header.map(key => cell_(obj[key])));
    obj._row = sh.getLastRow();
  }
}

/** Deletes matching rows, bottom up so row numbers stay valid. */
function deleteRows_(name, match) {
  const sh = sheet_(name);
  readAll_(name).filter(match).map(r => r._row).sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
}

/** Stops user text like "=IMPORTXML(...)" from becoming a formula. Values already quoted pass through. */
function cell_(value) {
  if (value === undefined || value === null) return '';
  if (typeof value === 'string' && value.charAt(0) !== '\'' && /^[=+\-@]/.test(value)) return "'" + value;
  return value;
}

function auth_(token) {
  const user = token ? readAll_('Users').find(u => u.token === token) : null;
  if (!user) throw new Error('Unknown link. Use your personal link.');
  return user;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function ms_(d) { return d instanceof Date ? d.getTime() : (d ? new Date(d).getTime() || null : null); }
function uid_() { return Utilities.getUuid().slice(0, 8); }
function newToken_() { return Utilities.getUuid().replace(/-/g, ''); }

// Lets .dev/test_helpers.js load the pure helpers in Node.
if (typeof module !== 'undefined') {
  module.exports = { COLUMNS, normalizeLink, israeliPhone, cleanAnalysis, adText, facebookText, adToken,
                     findAdObject, flatten, analyzeSchema, analyzePrompt, squash, hostOf };
}
