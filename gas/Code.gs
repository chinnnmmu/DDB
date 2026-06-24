// ══════════════════════════════════════════════
//  班表管理系統 — Google Apps Script Backend
//  Sheets: 當日排班 / 週報 / 自由編輯區 / 母檔 / 現走表 / 設定
// ══════════════════════════════════════════════

const SHEETS = {
  DAY:      '當日排班',
  WEEK:     '週報',
  NOTE:     '自由編輯區',
  DB:       '母檔',
  WALK:     '現走表',
  SETTINGS: '設定'
};

// 母檔可能超過 50k，需分 chunk 存多個 row
const CHUNK_SIZE = 45000;

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('班表管理系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Sheet helpers ──────────────────────────────

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/** 讀單一 cell A1（小資料用） */
function readCell_(name) {
  try {
    const v = getSheet_(name).getRange('A1').getValue();
    return v ? String(v) : '';
  } catch(e) { return ''; }
}

/** 寫單一 cell A1 */
function writeCell_(name, val) {
  try {
    getSheet_(name).getRange('A1').setValue(String(val || ''));
    return true;
  } catch(e) { return false; }
}

/** 讀 chunked（母檔用，A1 A2 A3… 拼接） */
function readChunked_(name) {
  try {
    const sheet = getSheet_(name);
    const last = sheet.getLastRow();
    if (!last) return '';
    return sheet.getRange(1, 1, last, 1).getValues()
      .map(r => String(r[0] || '')).join('');
  } catch(e) { return ''; }
}

/** 寫 chunked（自動分割超長字串） */
function writeChunked_(name, val) {
  try {
    const sheet = getSheet_(name);
    sheet.clearContents();
    if (!val) return true;
    const s = String(val);
    const chunks = [];
    for (let i = 0; i < s.length; i += CHUNK_SIZE)
      chunks.push([s.slice(i, i + CHUNK_SIZE)]);
    if (chunks.length) sheet.getRange(1, 1, chunks.length, 1).setValues(chunks);
    return true;
  } catch(e) { return false; }
}

// ── Public API（client 端呼叫）─────────────────

/** 一次載入所有 Sheet 資料 */
function loadAll() {
  return {
    day:      readCell_(SHEETS.DAY),
    week:     readCell_(SHEETS.WEEK),
    note:     readCell_(SHEETS.NOTE),
    db:       readChunked_(SHEETS.DB),
    walk:     readCell_(SHEETS.WALK),
    settings: readCell_(SHEETS.SETTINGS)
  };
}

/**
 * 儲存單一欄位
 * @param {string} key  day | week | note | db | walk | settings
 * @param {string} val  要存的字串
 */
function save(key, val) {
  switch (key) {
    case 'day':      return writeCell_(SHEETS.DAY, val);
    case 'week':     return writeCell_(SHEETS.WEEK, val);
    case 'note':     return writeCell_(SHEETS.NOTE, val);
    case 'db':       return writeChunked_(SHEETS.DB, val);
    case 'walk':     return writeCell_(SHEETS.WALK, val);
    case 'settings': return writeCell_(SHEETS.SETTINGS, val);
    default:         return false;
  }
}
