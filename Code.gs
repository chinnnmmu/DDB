// ════════════════════════════════════════════════════════════════
// DDB 工作台 - Google Apps Script 後端
// 部署：「網頁應用程式」，執行身份「我」，存取權「任何人」
// ════════════════════════════════════════════════════════════════

var SHEET_ID = '1bm_b-0CUiUGhqcNzPfQEXgI7V3fwdy1_9MT6yUQn2mw';

var SHEET_NAMES = {
  kv:      'KV_Store',     // 所有 key-value 資料（工單、便簽、對應表等）
  notepad: 'Notepad',      // 記事本內容
  rawLog:  'RawOrdersLog', // 轉檔器歷史備份
};

// ════════════════════════════════════════════════════════════════
// doGet — 回傳 HTML 頁面
// ════════════════════════════════════════════════════════════════
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('DDB 工作台')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ════════════════════════════════════════════════════════════════
// KV 工具
// ════════════════════════════════════════════════════════════════
function getOrCreateSheet(name) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function kvGet(key) {
  var sh = getOrCreateSheet(SHEET_NAMES.kv);
  var data = sh.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

function kvSet(key, value) {
  var sh = getOrCreateSheet(SHEET_NAMES.kv);
  var data = sh.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}

// ════════════════════════════════════════════════════════════════
// saveAllData — 儲存每日工單 + 便簽
// payload: { workOrders, multiNotes }
// ════════════════════════════════════════════════════════════════
function saveAllData(payload) {
  try {
    if (payload.workOrders !== undefined) {
      kvSet('workOrders', JSON.stringify(payload.workOrders));
    }
    if (payload.multiNotes !== undefined) {
      kvSet('multiNotes', JSON.stringify(payload.multiNotes));
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ════════════════════════════════════════════════════════════════
// loadAllData — 載入每日工單 + 便簽
// ════════════════════════════════════════════════════════════════
function loadAllData() {
  try {
    var workOrdersRaw = kvGet('workOrders');
    var multiNotesRaw = kvGet('multiNotes');
    return {
      workOrders: workOrdersRaw ? JSON.parse(workOrdersRaw) : {},
      multiNotes: multiNotesRaw ? JSON.parse(multiNotesRaw) : [],
      kvData: {}
    };
  } catch (e) {
    return { workOrders: {}, multiNotes: [], kvData: {} };
  }
}

// ════════════════════════════════════════════════════════════════
// 記事本 — 單一純文字，存在 Notepad sheet A1
// ════════════════════════════════════════════════════════════════
function saveMemo(text) {
  try {
    var sh = getOrCreateSheet(SHEET_NAMES.notepad);
    sh.getRange('A1').setValue(text || '');
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getMemo() {
  try {
    var sh = getOrCreateSheet(SHEET_NAMES.notepad);
    return sh.getRange('A1').getValue() || '';
  } catch (e) {
    return '';
  }
}

// ════════════════════════════════════════════════════════════════
// 轉檔器後台 — 妹妹→經紀對應 + 阿姨表
// ════════════════════════════════════════════════════════════════
function getInitData() {
  try {
    return {
      mmgg:   JSON.parse(kvGet('mmgg')   || '{}'),
      ggList: JSON.parse(kvGet('ggList') || '[]')
    };
  } catch (e) {
    return { mmgg: {}, ggList: [] };
  }
}

function updateMapping(mmgg, ggList) {
  try {
    kvSet('mmgg',   JSON.stringify(mmgg));
    kvSet('ggList', JSON.stringify(ggList));
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getAuntieList() {
  try {
    return JSON.parse(kvGet('auntieList') || '[]');
  } catch (e) {
    return [];
  }
}

function saveAuntieList(list) {
  try {
    kvSet('auntieList', JSON.stringify(list));
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ════════════════════════════════════════════════════════════════
// saveToSheet — 轉檔器：把當週工單寫入獨立分頁
// ════════════════════════════════════════════════════════════════
function saveToSheet(rows, raw, dateInput) {
  try {
    var ss  = SpreadsheetApp.openById(SHEET_ID);
    var now = new Date();
    var ts  = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');

    var weekTabName = Utilities.formatDate(getMonday(now), Session.getScriptTimeZone(), 'M/d') + '週';
    var sh = ss.getSheetByName(weekTabName);
    if (!sh) {
      sh = ss.insertSheet(weekTabName);
      sh.appendRow(['日期','阿姨','妹妹','經紀','時間','地點','PS','寫','牌價','實收']);
      sh.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#FCF9FA');
    }

    rows.forEach(function(r) {
      sh.appendRow([
        r.date  || '', r.ee    || '', r.mm      || '', r.gg  || '',
        r.time  || '', r.place || '', r.psCol   || '', r.writeCol || '',
        r.price || '', r.shou  || ''
      ]);
    });

    var logSh = getOrCreateSheet(SHEET_NAMES.rawLog);
    if (logSh.getLastRow() === 0) {
      logSh.appendRow(['Timestamp','Date','WeekTab','Count','Raw']);
    }
    logSh.appendRow([ts, dateInput, weekTabName, rows.length, raw]);

    return { success: true };
  } catch (e) {
    throw new Error('寫入失敗：' + e.message);
  }
}

function getRawOrdersHistory() {
  try {
    var sh   = getOrCreateSheet(SHEET_NAMES.rawLog);
    var data = sh.getDataRange().getValues();
    if (data.length <= 1) return [];
    var result = [];
    for (var i = data.length - 1; i >= 1; i--) {
      result.push({
        timestamp: data[i][0] || '',
        date:      data[i][1] || '',
        weekTab:   data[i][2] || '',
        count:     data[i][3] || 0,
        raw:       data[i][4] || ''
      });
      if (result.length >= 20) break;
    }
    return result;
  } catch (e) {
    return [];
  }
}

// ════════════════════════════════════════════════════════════════
// 工具
// ════════════════════════════════════════════════════════════════
function getMonday(d) {
  var day  = d.getDay();
  var diff = d.getDate() - day + (day === 0 ? -6 : 1);
  var mon  = new Date(d);
  mon.setDate(diff);
  mon.setHours(0, 0, 0, 0);
  return mon;
}
