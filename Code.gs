// ════════════════════════════════════════════════════════════════
// DDB 工作台 - Google Apps Script 後端
// 部署方式：在 Google Apps Script 中新增此檔案，以「網頁應用程式」
//           身份部署，存取權設為「任何人」
// ════════════════════════════════════════════════════════════════

// ── 設定：請填入你的 Google Spreadsheet ID ──────────────────────
var SHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';  // ← 請貼上你的試算表 ID

// Sheet 名稱
var SHEET_NAMES = {
  kv:      'KV_Store',      // key-value 儲存（便簽、班表設定等）
  orders:  'WorkOrders',    // 工單資料（JSON 格式）
  notes:   'MultiNotes',    // 多色便簽
  mapping: 'MmGgMapping',   // 妹妹→經紀 對應
  aunties: 'AuntieList',    // 阿姨列表
  rawLog:  'RawOrdersLog',  // 轉檔器原始備份
};

// ════════════════════════════════════════════════════════════════
// doGet — 回傳 HTML 頁面（如果從 GAS 部署）
// ════════════════════════════════════════════════════════════════
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('DDB 工作台')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ════════════════════════════════════════════════════════════════
// 工具函數
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
// saveAllData — 統一儲存（工單 + 便簽 + KV 設定）
// payload: { multiNotes, workOrders, kvData }
// ════════════════════════════════════════════════════════════════
function saveAllData(payload) {
  try {
    // 工單
    if (payload.workOrders !== undefined) {
      kvSet('workOrders', JSON.stringify(payload.workOrders));
    }
    // 多色便簽
    if (payload.multiNotes !== undefined) {
      kvSet('multiNotes', JSON.stringify(payload.multiNotes));
    }
    // KV 資料（班表、現走表等）
    if (payload.kvData) {
      var kv = payload.kvData;
      if (kv.notepad !== undefined)    kvSet('notepad', kv.notepad);
      if (kv.weekInput !== undefined)  kvSet('weekInput', kv.weekInput);
      if (kv.db !== undefined)         kvSet('db', kv.db);
      if (kv.walk !== undefined)       kvSet('walk', kv.walk);
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ════════════════════════════════════════════════════════════════
// loadAllData — 統一載入
// ════════════════════════════════════════════════════════════════
function loadAllData() {
  try {
    var workOrdersRaw = kvGet('workOrders');
    var multiNotesRaw = kvGet('multiNotes');
    var notepad       = kvGet('notepad')   || '';
    var weekInput     = kvGet('weekInput') || '';
    var db            = kvGet('db')        || '';
    var walkRaw       = kvGet('walk');

    var workOrders = workOrdersRaw ? JSON.parse(workOrdersRaw) : {};
    var multiNotes = multiNotesRaw ? JSON.parse(multiNotesRaw) : [];
    var walk = walkRaw ? JSON.parse(walkRaw) : null;

    return {
      workOrders: workOrders,
      multiNotes: multiNotes,
      kvData: {
        notepad:   notepad,
        weekInput: weekInput,
        db:        db,
        walk:      walk
      }
    };
  } catch (e) {
    return { workOrders: {}, multiNotes: [], kvData: {} };
  }
}

// ════════════════════════════════════════════════════════════════
// 轉檔器相關
// ════════════════════════════════════════════════════════════════

// 取得妹妹→經紀對應 + 經紀列表
function getInitData() {
  try {
    var mmggRaw  = kvGet('mmgg')   || '{}';
    var ggListRaw = kvGet('ggList') || '[]';
    return {
      mmgg:   JSON.parse(mmggRaw),
      ggList: JSON.parse(ggListRaw)
    };
  } catch (e) {
    return { mmgg: {}, ggList: [] };
  }
}

// 更新妹妹→經紀對應
function updateMapping(mmgg, ggList) {
  try {
    kvSet('mmgg',   JSON.stringify(mmgg));
    kvSet('ggList', JSON.stringify(ggList));
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// 取得阿姨列表
function getAuntieList() {
  try {
    var raw = kvGet('auntieList') || '[]';
    return JSON.parse(raw);
  } catch (e) {
    return [];
  }
}

// 儲存阿姨列表
function saveAuntieList(list) {
  try {
    kvSet('auntieList', JSON.stringify(list));
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ════════════════════════════════════════════════════════════════
// saveToSheet — 將轉換結果寫入 Google Sheets
// rows: 轉換後的工單陣列
// raw:  原始文字（備份用）
// dateInput: 日期字串
// ════════════════════════════════════════════════════════════════
function saveToSheet(rows, raw, dateInput) {
  try {
    var ss   = SpreadsheetApp.openById(SHEET_ID);
    var now  = new Date();
    var ts   = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');

    // 週次 Tab 名稱（以當週一日期命名）
    var weekMonday = getMonday(now);
    var weekTabName = Utilities.formatDate(weekMonday, Session.getScriptTimeZone(), 'M/d') + '週';

    // 找或建立週次 Sheet
    var sh = ss.getSheetByName(weekTabName);
    if (!sh) {
      sh = ss.insertSheet(weekTabName);
      sh.appendRow(['日期','阿姨','妹妹','經紀','時間','地點','PS','寫','牌價','實收','備用1','備用2','備用3','備用4','備用5','備用6','備用7']);
      sh.getRange(1, 1, 1, 17).setFontWeight('bold').setBackground('#FCF9FA');
    }

    // 寫入轉換後資料
    var headerRow = sh.getLastRow() === 0 ? 0 : sh.getLastRow();
    rows.forEach(function(r) {
      sh.appendRow([
        r.date   || '',
        r.ee     || '',
        r.mm     || '',
        r.gg     || '',
        r.time   || '',
        r.place  || '',
        r.psCol  || '',
        r.writeCol || '',
        r.price  || '',
        r.shou   || '',
        '', '', '', '', '', '', ''
      ]);
    });

    // 備份原始文字
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

// 取得歷史備份列表
function getRawOrdersHistory() {
  try {
    var sh = getOrCreateSheet(SHEET_NAMES.rawLog);
    var data = sh.getDataRange().getValues();
    if (data.length <= 1) return [];
    // 跳過標題行，最新的在前
    var result = [];
    for (var i = data.length - 1; i >= 1; i--) {
      result.push({
        timestamp: data[i][0] || '',
        date:      data[i][1] || '',
        weekTab:   data[i][2] || '',
        count:     data[i][3] || 0,
        raw:       data[i][4] || ''
      });
      if (result.length >= 20) break; // 最多顯示 20 筆
    }
    return result;
  } catch (e) {
    return [];
  }
}

// ════════════════════════════════════════════════════════════════
// 工具：取得當週星期一
// ════════════════════════════════════════════════════════════════
function getMonday(d) {
  var day = d.getDay();
  var diff = d.getDate() - day + (day === 0 ? -6 : 1);
  var mon = new Date(d);
  mon.setDate(diff);
  mon.setHours(0, 0, 0, 0);
  return mon;
}
