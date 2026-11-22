/**
 * Macro Tracker: photo or text -> Claude estimate -> row in the 'Log' sheet.
 * Container-bound script: create it from the sheet via Extensions > Apps Script.
 * The API key lives in Script Properties (Project Settings > Script properties > ANTHROPIC_API_KEY), never in this file.
 */
const CONFIG = {
  MODEL: 'claude-sonnet-5-5',   // change here to use a different Claude model
  MAX_TOKENS: 4000,   // generous: newer models may reason internally before answering
  USE_FOOD_LOOKUP: true,   // look up label data for branded products in Open Food Facts
  MAX_LOOKUP_ROUNDS: 3,    // how many lookup rounds Claude may take before it must answer
  LOOKUP_USER_AGENT: 'MacroTracker/1.0 (personal Google Apps Script)',
  LOG_SHEET: 'Log',
  TARGETS_SHEET: 'Targets'
};

/** The "form" Claude fills in, so the answer always arrives as structured data. */
const MEAL_TOOL = {
  name: 'log_meal',
  description: 'Record the nutrition estimate for the whole meal.',
  input_schema: {
    type: 'object',
    properties: {
      description: { type: 'string', description: 'Short meal description' },
      portion: { type: 'string', description: 'Portion assumptions in grams/ml' },
      kcal: { type: 'number' }, protein: { type: 'number' }, carbs: { type: 'number' },
      fat: { type: 'number' }, fiber: { type: 'number' },
      confidence: { type: 'string', enum: ['High', 'Medium', 'Low'] },
      notes: { type: 'string', description: 'Biggest uncertainty' },
      sources: { type: 'string', description: 'One short line: which values came from looked-up label data (and for which product) and which are your own estimates' }
    },
    required: ['description', 'portion', 'kcal', 'protein', 'carbs', 'fat', 'fiber', 'confidence', 'notes', 'sources']
  }
};

/** Lets Claude check official label values (per 100 g) for branded products. */
const LOOKUP_TOOL = {
  name: 'lookup_food',
  description: 'Look up nutrition label data (per 100 g) for ONE packaged or branded food product in the Open Food Facts database. ' +
    'Use it for named products (brand + product: crispbread, spreads, yoghurts, bars, sauces, drinks). ' +
    'Returns up to 5 candidate products; pick the best match, and say which one in your sources line. ' +
    'Do not use it for plain unbranded foods such as rice, an apple or a homemade dish.',
  input_schema: {
    type: 'object',
    properties: { query: { type: 'string', description: 'Brand and product name, e.g. "Finn Crisp Original Rye" or "Rügenwalder Mühle Teewurst"' } },
    required: ['query']
  }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Macros')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Memoised per request: each of these is a slow Google service call, and the old code repeated them for every row.
var MEMO_ = { ss: null, tz: null, sheets: {}, keys: {} };
function ss_() { return MEMO_.ss || (MEMO_.ss = SpreadsheetApp.getActive()); }
function tz_() { return MEMO_.tz || (MEMO_.tz = ss_().getSpreadsheetTimeZone()); }
function sheet_(name) {
  if (!(name in MEMO_.sheets)) MEMO_.sheets[name] = ss_().getSheetByName(name);
  return MEMO_.sheets[name];
}

/** Reads the Targets sheet by row label, so inserting rows there cannot break the app. */
function getTargets_() {
  const sh = sheet_(CONFIG.TARGETS_SHEET);
  if (!sh) throw new Error("Sheet 'Targets' not found.");
  const vals = sh.getRange(1, 1, Math.max(sh.getLastRow(), 2), 2).getValues();
  const byLabel = {};
  vals.forEach(function (r) { if (r[0] !== '') byLabel[String(r[0]).trim()] = r[1]; });
  const get = function (label) {
    if (!(label in byLabel)) throw new Error("Targets sheet: row '" + label + "' not found.");
    return Number(byLabel[label]) || 0;
  };
  return {
    kcal: get('Calorie target (kcal/day)'),
    maintenance: get('Maintenance calories (kcal/day)'),
    proteinMin: get('Protein min (g/day)'),
    proteinMax: get('Protein max (g/day)'),
    fat: get('Fat (g/day)'),
    carbs: get('Carbs (g/day)')
  };
}

function dateKey_(d) {
  if (Object.prototype.toString.call(d) !== '[object Date]') return String(d).slice(0, 10);
  const t = d.getTime();
  return MEMO_.keys[t] || (MEMO_.keys[t] = Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'));
}

/** Text that must never be interpreted as a formula when written to a cell. */
function safeText_(v, max) {
  let t = String(v === null || v === undefined ? '' : v).replace(/[\r\n]+/g, ' ').trim().slice(0, max || 200);
  if (/^[=+\-@]/.test(t)) t = "'" + t;
  return t;
}

const ACTIVITY_TYPES = ['Gym', 'Dance', 'Pilates', 'Self-defense', 'Long walk', 'Rest'];

/** "Dance, Gym" -> ['Dance', 'Gym'] (trimmed, de-duplicated, empty parts dropped). */
function parseActivity_(s) {
  return String(s === null || s === undefined ? '' : s).split(',')
    .map(function (x) { return x.trim(); })
    .filter(function (x, i, a) { return x && a.indexOf(x) === i; });
}

/** Reads Daily columns A (date), K (activity) and L (notes) into a map keyed by yyyy-MM-dd. */
function readActivityMap_() {
  const daily = sheet_('Daily');
  const map = {};
  if (!daily) return { map: map, exists: false };
  const last = daily.getLastRow();
  if (last >= 2) {
    const dates = daily.getRange(2, 1, last - 1, 1).getValues();
    const act = daily.getRange(2, 11, last - 1, 2).getValues();   // K (activity) and L (notes) only
    dates.forEach(function (r, i) {
      if (r[0] === '' || r[0] === null) return;
      const a = act[i];
      map[dateKey_(r[0])] = { row: i + 2, types: parseActivity_(a[0]), note: String(a[1] === null || a[1] === undefined ? '' : a[1]) };
    });
  }
  return { map: map, exists: true };
}

/** "9:05", "09:05", "9:05 PM" -> "09:05" (24 h) so meals can be sorted; '' stays ''. */
function normalizeTime_(t) {
  const m = /(\d{1,2}):(\d{2})\s*([AaPp][Mm])?/.exec(String(t || ''));
  if (!m) return '';
  let h = Number(m[1]);
  if (m[3]) { const pm = m[3].toLowerCase() === 'pm'; if (pm && h < 12) h += 12; if (!pm && h === 12) h = 0; }
  return (h < 10 ? '0' : '') + h + ':' + m[2];
}

/** Targets plus what has been logged on a given day (yyyy-MM-dd). Returns plain strings/numbers only (no Date objects). */
function getState(dateStr) {
  const log = sheet_(CONFIG.LOG_SHEET);
  if (!log) throw new Error("Sheet 'Log' not found.");
  const last = log.getLastRow();
  const totals = { kcal: 0, protein: 0, carbs: 0, fat: 0 };
  const meals = [];
  const recent = [];
  if (last >= 2) {
    const vals = log.getRange(2, 1, last - 1, 13).getValues();
    const times = log.getRange(2, 2, last - 1, 1).getDisplayValues();
    vals.forEach(function (r, i) {
      const d = r[0];
      if (d === '' || d === null) return;
      if (dateKey_(d) !== dateStr) return;
      totals.kcal += Number(r[5]) || 0;
      totals.protein += Number(r[6]) || 0;
      totals.carbs += Number(r[7]) || 0;
      totals.fat += Number(r[8]) || 0;
      meals.push({
        row: i + 2, meal: String(r[2]), time: String(times[i][0] || ''), desc: String(r[3]), portion: String(r[4]),
        kcal: Number(r[5]) || 0, protein: Number(r[6]) || 0, carbs: Number(r[7]) || 0, fat: Number(r[8]) || 0,
        fiber: Number(r[9]) || 0, confidence: String(r[10]), notes: String(r[12]), source: String(r[11])
      });
    });
    // Most recent distinct meals (any date), for one-tap "log again"
    const seen = {};
    for (let i = vals.length - 1; i >= 0 && recent.length < 6; i--) {
      const r = vals[i], desc = String(r[3] || '').trim();
      if (!desc || r[0] === '' || r[0] === null) continue;
      const k = desc.toLowerCase();
      if (seen[k]) continue;
      seen[k] = 1;
      recent.push({
        desc: desc, portion: String(r[4]), meal: String(r[2]), confidence: String(r[10]) || 'Medium',
        kcal: Number(r[5]) || 0, protein: Number(r[6]) || 0, carbs: Number(r[7]) || 0, fat: Number(r[8]) || 0, fiber: Number(r[9]) || 0
      });
    }
  }
  // Chronological order; entries without a time keep their entry order after the timed ones.
  meals.sort(function (a, b) {
    const ta = normalizeTime_(a.time) || '99:99', tb = normalizeTime_(b.time) || '99:99';
    return ta < tb ? -1 : ta > tb ? 1 : a.row - b.row;
  });
  const act = readActivityMap_();
  let activity = null;
  if (act.exists) {
    const e = act.map[dateStr];
    activity = e ? { types: e.types, note: e.note, inRange: true } : { types: [], note: '', inRange: false };
  }
  return { targets: getTargets_(), totals: totals, meals: meals, recent: recent, lastRow: last, activity: activity, activityTypes: ACTIVITY_TYPES };
}

/** Reads the key and strips whitespace, line breaks and stray quotes that often sneak in when pasting. */
function getApiKey_() {
  const raw = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  return raw ? String(raw).replace(/[\s"'`]/g, '') : '';
}

/**
 * Optional diagnostic: select "testApiKey" in the Apps Script editor and click Run, then open the Execution log.
 * It never prints the key itself, only its shape and what Anthropic answers.
 */
function testApiKey() {
  const key = getApiKey_();
  if (!key) { console.log('No ANTHROPIC_API_KEY found in Script properties.'); return; }
  console.log('Key length: ' + key.length + ' | starts with "sk-ant-": ' + (key.indexOf('sk-ant-') === 0));
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/models?limit=1', {
    method: 'get',
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    muteHttpExceptions: true
  });
  console.log('Anthropic answered with HTTP ' + res.getResponseCode() + (res.getResponseCode() === 200 ? ' - the key works.' : ' - ' + res.getContentText().slice(0, 200)));
}

/** Finds the first balanced {...} in free text that parses as JSON and looks like an estimate. */
function extractEstimate_(text) {
  for (let i = 0; i < text.length; i++) {
    if (text[i] !== '{') continue;
    let depth = 0, inStr = false, esc = false;
    for (let j = i; j < text.length; j++) {
      const ch = text[j];
      if (inStr) {
        if (esc) esc = false; else if (ch === '\\') esc = true; else if (ch === '"') inStr = false;
      } else if (ch === '"') inStr = true;
      else if (ch === '{') depth++;
      else if (ch === '}') {
        depth--;
        if (depth === 0) {
          try {
            const obj = JSON.parse(text.slice(i, j + 1));
            if (obj && typeof obj === 'object' && 'kcal' in obj) return obj;
          } catch (e) {}
          break;
        }
      }
    }
  }
  return null;
}

/** Looks up per-100 g label data in Open Food Facts. Never throws: failures come back as { error }. */
function lookupFood_(query) {
  const q = String(query || '').trim().slice(0, 100);
  if (!q) return { error: 'Empty query' };
  const cache = CacheService.getScriptCache();
  const ckey = 'off:' + q.toLowerCase().replace(/[^a-z0-9]+/g, '_').slice(0, 200);
  try { const hit = cache.get(ckey); if (hit) return JSON.parse(hit); } catch (e) {}
  let out;
  try {
    const url = 'https://world.openfoodfacts.org/cgi/search.pl?search_simple=1&action=process&json=1&page_size=8&sort_by=unique_scans_n' +
      '&fields=' + encodeURIComponent('code,product_name,brands,quantity,serving_size,nutriments') +
      '&search_terms=' + encodeURIComponent(q);
    const res = UrlFetchApp.fetch(url, { method: 'get', headers: { 'User-Agent': CONFIG.LOOKUP_USER_AGENT }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { error: 'Food database unavailable (HTTP ' + res.getResponseCode() + ')' };
    const products = (JSON.parse(res.getContentText()).products || []);
    const results = [];
    products.forEach(function (p) {
      const n = p.nutriments || {};
      const val = function (k) { const v = n[k]; return (v === undefined || v === null || v === '' || isNaN(Number(v))) ? null : Number(v); };
      let kcal = val('energy-kcal_100g');
      if (kcal === null && val('energy-kj_100g') !== null) kcal = Math.round(val('energy-kj_100g') / 4.184);
      if (kcal === null || !p.product_name || results.length >= 5) return;
      results.push({
        name: String(p.product_name).slice(0, 80), brand: String(p.brands || '').slice(0, 60),
        pack: String(p.quantity || ''), serving: String(p.serving_size || ''),
        per100g: { kcal: Math.round(kcal), protein: val('proteins_100g'), carbs: val('carbohydrates_100g'), fat: val('fat_100g'), fiber: val('fiber_100g') }
      });
    });
    out = results.length ? { source: 'Open Food Facts (crowd-sourced label data)', results: results }
                         : { results: [], note: 'No product with nutrition data found for this query.' };
  } catch (e) {
    return { error: 'Food database lookup failed' };
  }
  try { cache.put(ckey, JSON.stringify(out), 6 * 3600); } catch (e) {}
  return out;
}

/** One call to the Messages API. Retries transient failures and throws readable errors. */
function callClaude_(key, system, tools, messages) {
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: CONFIG.MODEL, max_tokens: CONFIG.MAX_TOKENS, system: system, tools: tools, messages: messages }),
    muteHttpExceptions: true
  };
  const TRANSIENT = [429, 500, 502, 503, 504, 529];
  let res, code, body;
  for (let attempt = 1; attempt <= 3; attempt++) {
    res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
    code = res.getResponseCode();
    body = res.getContentText();
    if (code === 200 || TRANSIENT.indexOf(code) < 0 || attempt === 3) break;
    Utilities.sleep(1500 * attempt);               // brief back-off, then try again
  }
  if (code !== 200) {
    let msg = body;
    try { msg = JSON.parse(body).error.message; } catch (e) {}
    if (code === 401) msg += ' (Check ANTHROPIC_API_KEY in Project Settings > Script properties: no spaces, quotes or line breaks, and it must start with sk-ant-.)';
    if (/credit balance/i.test(msg)) throw new Error('Your Anthropic credit balance is too low. Add credit at console.anthropic.com under Plans & Billing, then try again.');
    if (code === 429) throw new Error('Anthropic is rate-limiting requests right now. Wait a minute and try again.');
    if (code === 529 || code >= 500) throw new Error('Claude is temporarily overloaded. Please try again in a moment.');
    throw new Error('Claude API error (' + code + '): ' + msg);
  }
  const data = JSON.parse(body);
  console.log('Claude stop_reason=' + data.stop_reason + ' content=' + JSON.stringify(data.content).slice(0, 1500));
  return data;
}

/** Calls Claude (with optional food lookups). req = { text, image (base64, no prefix), mime, previous, correction } */
function estimateMeal(req) {
  const key = getApiKey_();
  if (!key) throw new Error('API key missing: add ANTHROPIC_API_KEY in Project Settings > Script properties.');
  if (!req.image && !(req.text && req.text.trim())) throw new Error('Add a photo or describe the meal.');

  const system =
    'You are a precise nutrition estimator for a macro-tracking app. The user lives in Berlin, so ' +
    'expect German/European foods, brands and restaurant portions; use metric units. ' +
    'From the photo and/or description, identify each item, state your portion assumption in grams or ml, ' +
    'then estimate the TOTAL for the whole meal. If a nutrition label is visible, use its values scaled to the portion. ' +
    (CONFIG.USE_FOOD_LOOKUP
      ? 'For NAMED packaged or branded products, first call lookup_food (one call per product, all in the same step) and use the label values per 100 g, ' +
        'scaled to the portion. Pick the closest matching variant; if no match is found or the match is doubtful, estimate from general knowledge and say so in sources. ' +
        'Looked-up data gives per-100 g values only: the portion in grams is still your estimate (use a serving or slice weight if the data or pack shows one). ' +
        'Never call lookup_food for plain unbranded foods. Only call log_meal after the lookups have returned. '
      : '') +
    'Account for cooking oil, dressings, sauces and drinks when plausible. ' +
    'Give realistic central estimates; do not under-estimate to be encouraging. ' +
    'Confidence: High = labelled or weighed items; Medium = typical plate; Low = hidden ingredients or unclear portion. ' +
    'Always finish by calling the log_meal tool. Keep description under 80 characters, portion under 120, notes under 120 ' +
    '(notes = the biggest uncertainty), sources under 160. All nutrient values are numbers: grams, except kcal.';

  const content = [];
  if (req.image) {
    content.push({ type: 'image', source: { type: 'base64', media_type: req.mime || 'image/jpeg', data: req.image } });
  }
  let userText = (req.text && req.text.trim())
    ? 'Meal description from user: ' + req.text.trim()
    : 'Estimate the meal in the photo.';
  if (req.previous && req.correction) {
    userText += '\n\nYour previous estimate (possibly edited by the user): ' + JSON.stringify(req.previous) +
      '\nThe user corrects or adds: ' + req.correction + '\nReturn the full updated estimate.';
  }
  content.push({ type: 'text', text: userText + '\n\nCall log_meal with your estimate' + (CONFIG.USE_FOOD_LOOKUP ? ' (after any lookups).' : '.') });

  const messages = [{ role: 'user', content: content }];
  let data = null, blocks = [], est = null;
  const maxRounds = CONFIG.USE_FOOD_LOOKUP ? CONFIG.MAX_LOOKUP_ROUNDS : 0;
  for (let round = 0; round <= maxRounds; round++) {
    const lastRound = round === maxRounds;
    const tools = (CONFIG.USE_FOOD_LOOKUP && !lastRound) ? [LOOKUP_TOOL, MEAL_TOOL] : [MEAL_TOOL];
    data = callClaude_(key, system, tools, messages);
    blocks = data.content || [];
    const meal = blocks.filter(function (x) { return x.type === 'tool_use' && x.name === 'log_meal'; })[0];
    if (meal && meal.input && typeof meal.input === 'object') { est = meal.input; break; }
    const lookups = blocks.filter(function (x) { return x.type === 'tool_use' && x.name === 'lookup_food'; });
    if (!lookups.length || lastRound) break;
    messages.push({ role: 'assistant', content: blocks });
    messages.push({ role: 'user', content: lookups.slice(0, 5).map(function (l) {
      return { type: 'tool_result', tool_use_id: l.id, content: JSON.stringify(lookupFood_(l.input && l.input.query)) };
    }) });
  }
  if (!est) {
    // Fallback: look for a JSON object inside any text the model returned
    const text = blocks.filter(function (x) { return x.type === 'text'; }).map(function (x) { return x.text; }).join('');
    est = extractEstimate_(text);
  }
  if (!est) {
    const kinds = blocks.map(function (x) { return x.type; }).join(',') || 'empty';
    if (data && data.stop_reason === 'max_tokens') throw new Error('The reply was cut off before the estimate was finished. Please try again.');
    throw new Error('Could not read the estimate (stop_reason: ' + (data && data.stop_reason) + ', content: ' + kinds + '). Please try again; if it repeats, open Executions in Apps Script and send me the log line "Claude stop_reason".');
  }
  const num = function (v) { return Math.max(0, Math.round(Number(v) || 0)); };
  return {
    description: String(est.description || '').slice(0, 120),
    portion: String(est.portion || '').slice(0, 200),
    kcal: num(est.kcal), protein: num(est.protein), carbs: num(est.carbs),
    fat: num(est.fat), fiber: num(est.fiber),
    confidence: ['High', 'Medium', 'Low'].indexOf(est.confidence) >= 0 ? est.confidence : 'Medium',
    notes: String(est.notes || '').slice(0, 200),
    sources: String(est.sources || '').slice(0, 220)
  };
}

/** Appends one meal to the Log sheet. m = { date, time, meal, description, portion, kcal, protein, carbs, fat, fiber, confidence, notes } */
function saveMeal(m) {
  if (!m || !/^\d{4}-\d{2}-\d{2}$/.test(String(m.date))) throw new Error('Pick a valid date.');
  const log = sheet_(CONFIG.LOG_SHEET);
  const date = Utilities.parseDate(m.date, tz_(), 'yyyy-MM-dd');
  const n = function (v) { return Math.max(0, Number(v) || 0); };
  const MEAL_TYPES = ['Breakfast', 'Lunch', 'Dinner', 'Snack', 'Drink'];
  if (MEAL_TYPES.indexOf(m.meal) < 0) throw new Error('Pick a meal type.');
  if (!String(m.description || '').trim()) throw new Error('Add a short description before saving.');
  const confidence = ['High', 'Medium', 'Low'].indexOf(m.confidence) >= 0 ? m.confidence : 'Medium';
  const row = log.getLastRow() + 1;
  // The sheet grid is finite (1,000 rows by default). Grow it before it fills up so saving never fails.
  if (typeof log.getMaxRows === 'function' && row > log.getMaxRows()) {
    log.insertRowsAfter(log.getMaxRows(), Math.max(500, row - log.getMaxRows()));
  }
  log.getRange(row, 1, 1, 13).setValues([[
    date, m.time || '', m.meal || '', safeText_(m.description, 120), safeText_(m.portion, 200),
    n(m.kcal), n(m.protein), n(m.carbs), n(m.fat), n(m.fiber),
    confidence, 'app', safeText_(m.notes, 400)
  ]]);
  log.getRange(row, 1).setNumberFormat('yyyy-mm-dd');
  log.getRange(row, 2).setNumberFormat('hh:mm');
  return getState(m.date);
}

/** Clears the last Log row, but only if this app wrote it. */
function undoLast(dateStr) {
  const log = sheet_(CONFIG.LOG_SHEET);
  const last = log.getLastRow();
  if (last >= 2 && log.getRange(last, 12).getValue() === 'app') {
    log.getRange(last, 1, 1, 13).clearContent();
  } else {
    throw new Error('The last row was not added by this app, so I left it alone.');
  }
  return getState(dateStr);
}

/** Per-day totals for a date range (yyyy-MM-dd, inclusive, max 62 days) for the Averages view. */
function getPeriod(startStr, endStr) {
  const re = /^\d{4}-\d{2}-\d{2}$/;
  if (!re.test(String(startStr)) || !re.test(String(endStr))) throw new Error('Invalid date range.');
  const utc = function (k) { const a = k.split('-'); return Date.UTC(Number(a[0]), Number(a[1]) - 1, Number(a[2])); };
  const s0 = utc(startStr), n = Math.round((utc(endStr) - s0) / 86400000) + 1;
  if (n < 1 || n > 62) throw new Error('Date range must be between 1 and 62 days.');
  const byDay = {};
  const log = sheet_(CONFIG.LOG_SHEET);
  const last = log.getLastRow();
  if (last >= 2) {
    log.getRange(2, 1, last - 1, 9).getValues().forEach(function (r) {
      if (r[0] === '' || r[0] === null) return;
      const k = dateKey_(r[0]);
      const v = byDay[k] || (byDay[k] = { kcal: 0, protein: 0, carbs: 0, fat: 0, n: 0 });
      v.kcal += Number(r[5]) || 0; v.protein += Number(r[6]) || 0; v.carbs += Number(r[7]) || 0; v.fat += Number(r[8]) || 0; v.n++;
    });
  }
  const actMap = readActivityMap_().map;
  const days = [];
  for (let i = 0; i < n; i++) {
    const k = new Date(s0 + i * 86400000).toISOString().slice(0, 10);
    const v = byDay[k];
    days.push({
      date: k, logged: !!v,
      kcal: v ? Math.round(v.kcal) : 0, protein: v ? Math.round(v.protein) : 0,
      carbs: v ? Math.round(v.carbs) : 0, fat: v ? Math.round(v.fat) : 0,
      activity: actMap[k] ? actMap[k].types : [], note: actMap[k] ? actMap[k].note.slice(0, 80) : ''
    });
  }
  const t = getTargets_();
  return { days: days, targets: { kcal: t.kcal, proteinMin: t.proteinMin, proteinMax: t.proteinMax, carbs: t.carbs, fat: t.fat } };
}

/**
 * Sets the activity for one date in the Daily sheet (column K, comma-separated) and optionally its note (column L).
 * types: array such as ['Gym'] or ['Dance', 'Gym'] or ['Rest']; note: string, or null to leave the note unchanged.
 */
function setActivity(dateStr, types, note) {
  if (!Array.isArray(types)) throw new Error('Activity must be a list.');
  const act = readActivityMap_();
  if (!act.exists) throw new Error("Sheet 'Daily' not found.");
  const e = act.map[dateStr];
  if (!e) {
    const keys = Object.keys(act.map).sort();
    throw new Error('That date is outside the Daily table' + (keys.length ? ' (it covers ' + keys[0] + ' to ' + keys[keys.length - 1] + ')' : '') +
      '. To extend it, copy the last row of the Daily tab down.');
  }
  types.forEach(function (t) {
    const x = String(t).trim();
    if (x.length > 30 || /[\r\n]/.test(x)) throw new Error('Activity names must be short single lines.');
  });
  let list = parseActivity_(types.join(','));
  if (list.length > 1) list = list.filter(function (x) { return x !== 'Rest'; });   // Rest excludes everything else
  const daily = sheet_('Daily');
  daily.getRange(e.row, 11).setValue(list.join(', '));
  if (note !== null && note !== undefined) {
    let txt = String(note).replace(/[\r\n]+/g, ' ').trim().slice(0, 80);
    if (/^[=+\-@]/.test(txt)) txt = "'" + txt;   // never let a note become a formula
    daily.getRange(e.row, 12).setValue(txt);
  }
  SpreadsheetApp.flush();
  return getState(dateStr);
}

/** Confirms that a Log row is still the entry the app thinks it is (guards against stale rows). */
function checkMealRef_(log, ref) {
  const row = Number(ref && ref.row), last = log.getLastRow();
  if (!(row >= 2 && row <= last)) throw new Error('That entry no longer exists. Reload and try again.');
  const r = log.getRange(row, 1, 1, 13).getValues()[0];
  if (r[0] === '' || r[0] === null || dateKey_(r[0]) !== ref.date || String(r[3]) !== String(ref.desc)) {
    throw new Error('That entry changed in the sheet. Reload and try again.');
  }
  return row;
}

/** Edits the description and numbers of one logged meal. ref = { row, date, desc } (as last seen), m = { description, kcal, protein, carbs, fat, fiber }. */
function updateMeal(ref, m) {
  const log = sheet_(CONFIG.LOG_SHEET);
  const row = checkMealRef_(log, ref);
  const n = function (v) { return Math.max(0, Number(v) || 0); };
  log.getRange(row, 4).setValue(safeText_(m.description, 120));
  log.getRange(row, 6, 1, 5).setValues([[n(m.kcal), n(m.protein), n(m.carbs), n(m.fat), n(m.fiber)]]);
  SpreadsheetApp.flush();
  return getState(ref.date);
}

/** Removes one logged meal (clears its row). ref = { row, date, desc }. */
function deleteMeal(ref) {
  const log = sheet_(CONFIG.LOG_SHEET);
  const row = checkMealRef_(log, ref);
  log.getRange(row, 1, 1, 13).clearContent();
  SpreadsheetApp.flush();
  return getState(ref.date);
}
