/**
 * Expense Tracker – Google Apps Script backend
 * Receipt scan: Google Drive OCR (image -> text) + DeepSeek (text -> structured JSON)
 */

const CONFIG = {
  SHEET_EXPENSES: 'Expenses',
  SHEET_USERS: 'Users',
  RECEIPT_FOLDER: 'Expense Receipts',
  CATEGORIES: [
    'Food & Beverage', 'Groceries', 'Transport', 'Fuel', 'Utilities',
    'Office Supplies', 'Travel', 'Entertainment', 'Healthcare', 'Shopping', 'Others'
  ],
  DEEPSEEK_URL: 'https://api.deepseek.com/chat/completions',
  DEEPSEEK_MODEL: 'deepseek-chat',
  SESSION_SECONDS: 6 * 60 * 60,
  MAX_LOGIN_FAILS: 5,
};

const EXPENSE_HEADERS = ['Date', 'Pay To', 'Item', 'Amount', 'Category', 'User', 'Submitted At', 'Receipt'];
const USER_HEADERS = ['Username', 'Salt', 'PIN Hash', 'Role', 'Active'];

/* ------------------------------------------------------------------ */
/* Web app entry                                                       */
/* ------------------------------------------------------------------ */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Expense Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

/* ------------------------------------------------------------------ */
/* Sheet menu + setup (run from the spreadsheet)                       */
/* ------------------------------------------------------------------ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Expense App')
    .addItem('1. Setup sheets', 'setup')
    .addItem('2. Set DeepSeek API key', 'promptApiKey')
    .addItem('Add / reset user', 'promptAddUser')
    .addToUi();
}

function setup() {
  const ss = getSpreadsheet_();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  const exp = ss.getSheetByName(CONFIG.SHEET_EXPENSES) || ss.insertSheet(CONFIG.SHEET_EXPENSES);
  if (exp.getLastRow() === 0) {
    exp.appendRow(EXPENSE_HEADERS);
    exp.getRange(1, 1, 1, EXPENSE_HEADERS.length).setFontWeight('bold').setBackground('#e8f0fe');
    exp.setFrozenRows(1);
    exp.getRange('A:A').setNumberFormat('yyyy-mm-dd');
    exp.getRange('D:D').setNumberFormat('#,##0.00');
    exp.getRange('G:G').setNumberFormat('yyyy-mm-dd hh:mm');
  }

  const users = ss.getSheetByName(CONFIG.SHEET_USERS) || ss.insertSheet(CONFIG.SHEET_USERS);
  if (users.getLastRow() === 0) {
    users.appendRow(USER_HEADERS);
    users.getRange(1, 1, 1, USER_HEADERS.length).setFontWeight('bold').setBackground('#e8f0fe');
    users.setFrozenRows(1);
  }

  getReceiptFolder_();
  SpreadsheetApp.getUi().alert('Setup done. Next: set the DeepSeek API key, then add users.');
}

function promptApiKey() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('DeepSeek API key', 'Paste your DeepSeek API key (stored in Script Properties, not in the sheet):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = res.getResponseText().trim();
  if (!key) return;
  PropertiesService.getScriptProperties().setProperty('DEEPSEEK_API_KEY', key);
  ui.alert('API key saved.');
}

function promptAddUser() {
  const ui = SpreadsheetApp.getUi();
  const u = ui.prompt('Add / reset user', 'Username:', ui.ButtonSet.OK_CANCEL);
  if (u.getSelectedButton() !== ui.Button.OK) return;
  const p = ui.prompt('Add / reset user', 'PIN / password (min 4 chars):', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  const r = ui.alert('Make this user an admin? (admins see everyone\'s expenses)', ui.ButtonSet.YES_NO);
  addOrResetUser_(u.getResponseText(), p.getResponseText(), r === ui.Button.YES ? 'admin' : 'user');
  ui.alert('User saved.');
}

function addOrResetUser_(username, pin, role) {
  username = String(username || '').trim().toLowerCase();
  pin = String(pin || '');
  if (!/^[a-z0-9._-]{2,40}$/.test(username)) throw new Error('Username: 2-40 chars, letters/numbers/._- only');
  if (pin.length < 4) throw new Error('PIN must be at least 4 characters');

  const sh = getSpreadsheet_().getSheetByName(CONFIG.SHEET_USERS);
  const salt = Utilities.getUuid();
  const row = [username, salt, hash_(salt, pin), role, true];
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === username) {
      sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return;
    }
  }
  sh.appendRow(row);
}

/* ------------------------------------------------------------------ */
/* Auth                                                                */
/* ------------------------------------------------------------------ */

function login(username, pin) {
  username = String(username || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + username;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= CONFIG.MAX_LOGIN_FAILS) throw new Error('Too many attempts. Try again in 15 minutes.');

  const data = getSpreadsheet_().getSheetByName(CONFIG.SHEET_USERS).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const [u, salt, pinHash, role, active] = data[i];
    if (String(u).toLowerCase() === username && active !== false && hash_(salt, String(pin)) === pinHash) {
      const token = Utilities.getUuid();
      cache.put('sess_' + token, JSON.stringify({ username: username, role: role || 'user' }), CONFIG.SESSION_SECONDS);
      cache.remove(failKey);
      return { token: token, username: username, role: role || 'user', categories: CONFIG.CATEGORIES };
    }
  }
  cache.put(failKey, String(fails + 1), 15 * 60);
  throw new Error('Invalid username or PIN');
}

function logout(token) {
  CacheService.getScriptCache().remove('sess_' + token);
  return true;
}

function resumeSession(token) {
  const s = requireUser_(token);
  return { token: token, username: s.username, role: s.role, categories: CONFIG.CATEGORIES };
}

function requireUser_(token) {
  const raw = token && CacheService.getScriptCache().get('sess_' + token);
  if (!raw) throw new Error('SESSION_EXPIRED');
  return JSON.parse(raw);
}

function hash_(salt, pin) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

/* ------------------------------------------------------------------ */
/* Receipt scanning                                                    */
/* ------------------------------------------------------------------ */

/**
 * @param {string} token   session token
 * @param {string} base64  image data (no data: prefix)
 * @param {string} mimeType e.g. image/jpeg
 */
function scanReceipt(token, base64, mimeType) {
  const user = requireUser_(token);
  if (!base64) throw new Error('No image received');

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType || 'image/jpeg', user.username + '-' + stamp + '.jpg');

  // Keep a copy of the receipt image in Drive
  const file = getReceiptFolder_().createFile(blob);

  const text = ocrImage_(blob);
  if (!text.trim()) {
    return { fileId: file.getId(), fileUrl: file.getUrl(), rawText: '', data: emptyExpense_(), warning: 'No text detected – please fill in manually.' };
  }

  const data = extractWithDeepSeek_(text);
  return { fileId: file.getId(), fileUrl: file.getUrl(), rawText: text, data: data };
}

/** Image -> text using Google Drive's built-in OCR (Drive advanced service v3). */
function ocrImage_(blob) {
  const doc = Drive.Files.create(
    { name: 'tmp-ocr-' + Utilities.getUuid(), mimeType: MimeType.GOOGLE_DOCS },
    blob,
    { ocrLanguage: 'en', fields: 'id' }
  );
  try {
    return DocumentApp.openById(doc.id).getBody().getText();
  } finally {
    DriveApp.getFileById(doc.id).setTrashed(true);
  }
}

/** OCR text -> {date, payTo, item, amount, category} using DeepSeek. */
function extractWithDeepSeek_(receiptText) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('DEEPSEEK_API_KEY');
  if (!apiKey) throw new Error('DeepSeek API key not set. Use the "Expense App" menu in the sheet.');

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const system =
    'You extract expense data from OCR text of a shop receipt. Reply with JSON only, using exactly these keys:\n' +
    '{"date":"YYYY-MM-DD","payTo":"merchant name","item":"short description of items bought","amount":0.00,"category":"..."}\n' +
    'Rules:\n' +
    '- date: the transaction date on the receipt, ISO format. If missing, use ' + today + '.\n' +
    '- payTo: the merchant / shop / company name (not the address).\n' +
    '- item: list the main items, comma separated, max 120 characters. If only one item, just its name.\n' +
    '- amount: the final TOTAL paid (after tax/discount/rounding) as a number, no currency symbol.\n' +
    '- category: choose exactly one of: ' + CONFIG.CATEGORIES.join(', ') + '.\n' +
    '- OCR text may be messy; use your best judgement. Use "" for unknown text and 0 for unknown amount.';

  const res = UrlFetchApp.fetch(CONFIG.DEEPSEEK_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      model: CONFIG.DEEPSEEK_MODEL,
      temperature: 0,
      response_format: { type: 'json_object' },
      messages: [
        { role: 'system', content: system },
        { role: 'user', content: 'Receipt OCR text:\n"""\n' + receiptText.slice(0, 8000) + '\n"""' }
      ]
    })
  });

  const code = res.getResponseCode();
  if (code !== 200) throw new Error('DeepSeek API error ' + code + ': ' + res.getContentText().slice(0, 300));

  const content = JSON.parse(res.getContentText()).choices[0].message.content;
  let parsed;
  try {
    parsed = JSON.parse(content);
  } catch (e) {
    const m = content.match(/\{[\s\S]*\}/);
    parsed = m ? JSON.parse(m[0]) : {};
  }
  return normaliseExpense_(parsed);
}

function emptyExpense_() {
  return normaliseExpense_({});
}

function normaliseExpense_(o) {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const date = /^\d{4}-\d{2}-\d{2}$/.test(String(o.date || '')) ? String(o.date) : today;
  const amount = Number(String(o.amount == null ? '' : o.amount).replace(/[^0-9.\-]/g, '')) || 0;
  const category = CONFIG.CATEGORIES.indexOf(o.category) >= 0 ? o.category : 'Others';
  return {
    date: date,
    payTo: String(o.payTo || '').trim(),
    item: String(o.item || '').trim(),
    amount: Math.round(amount * 100) / 100,
    category: category
  };
}

/* ------------------------------------------------------------------ */
/* Save + list                                                         */
/* ------------------------------------------------------------------ */

function saveExpense(token, exp, fileId) {
  const user = requireUser_(token);
  const e = normaliseExpense_(exp || {});
  if (!e.payTo) throw new Error('Pay To is required');
  if (!(e.amount > 0)) throw new Error('Amount must be greater than 0');

  const parts = e.date.split('-').map(Number);
  const receiptUrl = fileId ? 'https://drive.google.com/file/d/' + fileId + '/view' : '';

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSpreadsheet_().getSheetByName(CONFIG.SHEET_EXPENSES);
    sh.appendRow([
      new Date(parts[0], parts[1] - 1, parts[2]),
      safeText_(e.payTo),
      safeText_(e.item),
      e.amount,
      e.category,
      user.username,
      new Date(),
      receiptUrl
    ]);
  } finally {
    lock.releaseLock();
  }
  return true;
}

/** Recent expenses for the current user (admins see everyone) + this month's total. */
function getRecentExpenses(token) {
  const user = requireUser_(token);
  const sh = getSpreadsheet_().getSheetByName(CONFIG.SHEET_EXPENSES);
  const last = sh.getLastRow();
  if (last < 2) return { items: [], monthTotal: 0 };

  const rows = sh.getRange(2, 1, last - 1, EXPENSE_HEADERS.length).getValues();
  const tz = Session.getScriptTimeZone();
  const thisMonth = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  const isAdmin = user.role === 'admin';

  let monthTotal = 0;
  const items = [];
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    if (!isAdmin && r[5] !== user.username) continue;
    const d = r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : String(r[0]);
    if (d.slice(0, 7) === thisMonth) monthTotal += Number(r[3]) || 0;
    if (items.length < 20) {
      items.push({ date: d, payTo: String(r[1]), item: String(r[2]), amount: Number(r[3]) || 0, category: String(r[4]), user: String(r[5]), receipt: String(r[7]) });
    }
  }
  return { items: items, monthTotal: Math.round(monthTotal * 100) / 100 };
}

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function getReceiptFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('RECEIPT_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* recreate below */ }
  }
  const folder = DriveApp.createFolder(CONFIG.RECEIPT_FOLDER);
  props.setProperty('RECEIPT_FOLDER_ID', folder.getId());
  return folder;
}

/** Stop text like "=IMPORTXML(...)" from being treated as a formula. */
function safeText_(s) {
  s = String(s || '').slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
