const SHEET_ID = '1ZNN1B8_56isckH5LeWbuBl6YJ4Ng9FVZeeofGdtb780';
const EDITOR_TOKEN_DAYS = 30;
const MAX_AUTH_FAILURES = 10;
const AUTH_LOCK_SECONDS = 15 * 60;

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  try {
    if (action === 'getInventory') return jsonResponse_(getInventory());
    return jsonResponse_({ error: 'Unknown action' });
  } catch (err) {
    return jsonResponse_({ error: err.message || String(err) });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = payload.action;

    if (action === 'authorizeEditor') {
      return jsonResponse_(authorizeEditor_(payload));
    }

    requireEditor_(payload.editorToken);

    if (action === 'updateCoating') {
      return jsonResponse_(withWriteLock_(function () {
        return updateCoating(payload.name, payload.qty);
      }));
    }
    if (action === 'updateChip') {
      return jsonResponse_(withWriteLock_(function () {
        return updateChip(payload.name, payload.qty);
      }));
    }
    if (action === 'saveJobs') {
      return jsonResponse_(withWriteLock_(function () {
        return saveJobs(payload.jobs, payload.allowEmpty === true);
      }));
    }
    return jsonResponse_({ error: 'Unknown action' });
  } catch (err) {
    return jsonResponse_({ error: err.message || String(err) });
  }
}

function jsonResponse_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function authorizeEditor_(payload) {
  if (payload.editorToken && verifyEditorToken_(payload.editorToken)) {
    return { authorized: true };
  }

  const cache = CacheService.getScriptCache();
  const failures = Number(cache.get('editor_auth_failures') || 0);
  if (failures >= MAX_AUTH_FAILURES) {
    throw new Error('Too many incorrect PIN attempts. Try again in 15 minutes.');
  }

  const configuredPin = PropertiesService.getScriptProperties().getProperty('EDITOR_PIN');
  if (!configuredPin) throw new Error('Editor PIN is not configured.');

  if (!safeEqual_(String(payload.pin || ''), configuredPin)) {
    cache.put('editor_auth_failures', String(failures + 1), AUTH_LOCK_SECONDS);
    throw new Error('Invalid PIN');
  }

  cache.remove('editor_auth_failures');
  return { authorized: true, editorToken: issueEditorToken_() };
}

function requireEditor_(token) {
  if (!verifyEditorToken_(token)) throw new Error('Editor authorization required');
}

function issueEditorToken_() {
  const expires = Date.now() + EDITOR_TOKEN_DAYS * 24 * 60 * 60 * 1000;
  const payload = Utilities.base64EncodeWebSafe(
    JSON.stringify({ exp: expires, nonce: Utilities.getUuid() })
  ).replace(/=+$/g, '');
  return payload + '.' + sign_(payload);
}

function verifyEditorToken_(token) {
  if (typeof token !== 'string') return false;
  const parts = token.split('.');
  if (parts.length !== 2 || !safeEqual_(parts[1], sign_(parts[0]))) return false;
  try {
    const decoded = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
    const payload = JSON.parse(decoded);
    return Number.isFinite(payload.exp) && payload.exp > Date.now();
  } catch (_) {
    return false;
  }
}

function sign_(value) {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('EDITOR_TOKEN_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('EDITOR_TOKEN_SECRET', secret);
  }
  return Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(value, secret)
  ).replace(/=+$/g, '');
}

function safeEqual_(a, b) {
  a = String(a);
  b = String(b);
  if (a.length !== b.length) return false;
  let difference = 0;
  for (let i = 0; i < a.length; i++) difference |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return difference === 0;
}

function withWriteLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function getInventory() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const coatings = requiredSheet_(ss, 'Coatings').getDataRange().getValues().slice(1)
    .filter(function (r) { return r[0]; })
    .map(function (r) { return { name: String(r[0]), qty: finiteNumber_(r[1], 0) }; });

  const chips = requiredSheet_(ss, 'Chips').getDataRange().getValues().slice(1)
    .filter(function (r) { return r[0]; })
    .map(function (r) { return { name: String(r[0]), qty: finiteNumber_(r[1], 0) }; });

  const jobRows = requiredSheet_(ss, 'Jobs').getDataRange().getValues().slice(1)
    .filter(function (r) { return r[0]; });

  const jobs = jobRows.map(function (r) {
    return {
      name: String(r[0]),
      sqft: finiteNumber_(r[1], 0),
      jobType: r[2] || 'standard',
      chipColor: r[3] || '',
      anchorBSku: r[4] || null,
      tint: r[5] || null,
      chipOverride: r[6] ? parseInt(r[6], 10) : null,
      installDate: normalizeDate_(r[7])
    };
  });

  return { coatings: coatings, chips: chips, jobs: jobs };
}

function updateCoating(name, qty) {
  return updateQuantity_('Coatings', name, qty);
}

function updateChip(name, qty) {
  return updateQuantity_('Chips', name, qty);
}

function updateQuantity_(sheetName, name, qty) {
  name = cleanText_(name, 100, 'Product name');
  qty = boundedNumber_(qty, 0, 10000, 'Quantity');

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = requiredSheet_(ss, sheetName);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === name) {
      sheet.getRange(i + 1, 2).setValue(qty);
      return { status: 'ok' };
    }
  }
  throw new Error('Product not found');
}

function saveJobs(jobs, allowEmpty) {
  if (!Array.isArray(jobs)) throw new Error('Jobs must be an array');
  if (jobs.length > 250) throw new Error('Too many jobs');
  if (jobs.length === 0 && !allowEmpty) throw new Error('Refusing to clear jobs without confirmation');

  const rows = jobs.map(validateJob_);
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = requiredSheet_(ss, 'Jobs');
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 8).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, 8).setValues(rows);
  return { status: 'ok' };
}

function validateJob_(job) {
  if (!job || typeof job !== 'object') throw new Error('Invalid job');
  const jobType = String(job.jobType || 'standard');
  if (['standard', 'outdoor', 'double'].indexOf(jobType) === -1) {
    throw new Error('Invalid job type');
  }

  const installDate = normalizeDate_(job.installDate);
  if (!installDate) throw new Error('Invalid install date');

  let chipOverride = '';
  if (job.chipOverride !== null && job.chipOverride !== undefined && job.chipOverride !== '') {
    chipOverride = Math.round(boundedNumber_(job.chipOverride, 1, 100, 'Chip override'));
  }

  return [
    sheetSafeText_(cleanText_(job.name, 120, 'Job name')),
    boundedNumber_(job.sqft, 1, 100000, 'Square footage'),
    jobType,
    sheetSafeText_(cleanText_(job.chipColor, 100, 'Chip color')),
    optionalSheetText_(job.anchorBSku, 100),
    optionalSheetText_(job.tint, 100),
    chipOverride,
    installDate
  ];
}

function requiredSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name);
  return sheet;
}

function normalizeDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const text = String(value || '');
  const match = text.match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : '';
}

function finiteNumber_(value, fallback) {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function boundedNumber_(value, min, max, label) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < min || number > max) {
    throw new Error(label + ' is invalid');
  }
  return number;
}

function cleanText_(value, maxLength, label) {
  const text = String(value == null ? '' : value).trim();
  if (!text || text.length > maxLength) throw new Error(label + ' is invalid');
  return text;
}

function optionalSheetText_(value, maxLength) {
  const text = String(value == null ? '' : value).trim();
  if (text.length > maxLength) throw new Error('Text value is too long');
  return sheetSafeText_(text);
}

function sheetSafeText_(text) {
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}
