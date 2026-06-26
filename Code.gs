// ==================== DDB GAS Web App 後端 ====================

var SHEET_NAME_NOTES  = '便簽';
var SHEET_NAME_ORDERS = '工單';
var SHEET_NAME_META   = '同步紀錄';
var SHEET_NAME_KV     = '班表資料';

// ---- 入口點 ----
function doGet(e) {
  var page = e && e.parameter && e.parameter.page === 'admin' ? 'Admin' : 'Index';
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle(page === 'Admin' ? '班表管理系統' : 'DDB 工作台')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---- 工具：取得或建立 Sheet ----
function getOrCreateSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

// ==================== 主存檔：saveAllData ====================
// 一次儲存所有資料：便簽、工單、班表資料(KV)
function saveAllData(payload) {
  try {
    if (typeof payload === 'string') payload = JSON.parse(payload);

    // 1. 儲存便簽
    if (payload.multiNotes && Array.isArray(payload.multiNotes)) {
      var sh = getOrCreateSheet(SHEET_NAME_NOTES);
      sh.clearContents();
      sh.appendRow(['編號', '時間', '內容']);
      payload.multiNotes.forEach(function(n) {
        sh.appendRow([n.id || '', n.time || '', n.text || '']);
      });
    }

    // 2. 儲存工單
    if (payload.workOrders && typeof payload.workOrders === 'object') {
      var shO = getOrCreateSheet(SHEET_NAME_ORDERS);
      shO.clearContents();
      shO.appendRow(['日期','牌','人名','時間地點','阿姨','實收','EE','退','備記','已結','結帳日','結帳金額','編號']);
      Object.keys(payload.workOrders).forEach(function(dateKey) {
        var rows = payload.workOrders[dateKey] || [];
        rows.forEach(function(r) {
          shO.appendRow([
            dateKey,
            r.pai || '', r.name || '', r.place || '',
            r.auntie || '', r.shi || '', r.ee || '', r.rebate || '',
            r.note || '', r.settled ? 'Y' : 'N',
            r.settledDate || '', r.settledAmount || '', r.id || ''
          ]);
        });
      });
    }

    // 3. 儲存班表資料（KV：day/week/note/db/walk 等）
    if (payload.kvData && typeof payload.kvData === 'object') {
      var shK = getOrCreateSheet(SHEET_NAME_KV);
      var existing = shK.getDataRange().getValues();
      var kvMap = {};
      // 建立現有資料的 map
      existing.forEach(function(row, i) { if (row[0]) kvMap[String(row[0])] = i + 1; });
      // 逐一更新或新增
      Object.keys(payload.kvData).forEach(function(key) {
        var val = payload.kvData[key] || '';
        if (kvMap[key]) {
          shK.getRange(kvMap[key], 2).setValue(val);
        } else {
          shK.appendRow([key, val]);
          kvMap[key] = shK.getLastRow();
        }
      });
    }

    // 4. 記錄同步時間
    var shM = getOrCreateSheet(SHEET_NAME_META);
    shM.clearContents();
    shM.appendRow(['項目', '時間']);
    shM.appendRow(['最後同步', new Date().toLocaleString('zh-TW')]);

    return { success: true, message: '同步成功' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ==================== 主讀取：loadAllData ====================
// 一次讀取所有資料：便簽、工單、班表資料(KV)
function loadAllData() {
  try {
    var result = { multiNotes: [], workOrders: {}, kvData: { day:'', week:'', note:'', db:'', walk:'' } };

    // 1. 讀取便簽
    var sh = getOrCreateSheet(SHEET_NAME_NOTES);
    var data = sh.getDataRange().getValues();
    if (data.length > 1) {
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        if (row[0] || row[2]) {
          result.multiNotes.push({ id: row[0], time: row[1], text: row[2] });
        }
      }
    }

    // 2. 讀取工單
    var shO = getOrCreateSheet(SHEET_NAME_ORDERS);
    var oData = shO.getDataRange().getValues();
    if (oData.length > 1) {
      for (var j = 1; j < oData.length; j++) {
        var r = oData[j];
        var dateKey = String(r[0]);
        if (!dateKey) continue;
        if (!result.workOrders[dateKey]) result.workOrders[dateKey] = [];
        result.workOrders[dateKey].push({
          pai: String(r[1]), name: String(r[2]), place: String(r[3]),
          auntie: String(r[4]), shi: String(r[5]), ee: String(r[6]),
          rebate: String(r[7]), note: String(r[8]),
          settled: r[9] === 'Y',
          settledDate: String(r[10]), settledAmount: String(r[11]),
          id: String(r[12])
        });
      }
    }

    // 3. 讀取班表資料（KV）
    var shK = getOrCreateSheet(SHEET_NAME_KV);
    var kData = shK.getDataRange().getValues();
    kData.forEach(function(row) {
      var k = String(row[0]);
      if (k && k in result.kvData) {
        result.kvData[k] = String(row[1] || '');
      }
    });

    return result;
  } catch (err) {
    return { multiNotes: [], workOrders: {}, kvData: { day:'', week:'', note:'', db:'', walk:'' }, error: err.message };
  }
}

// ==================== 轉檔器 Sheets ====================
var SHEET_NAME_MMGG       = '妹妹配對';
var SHEET_NAME_AUNTIE     = '阿姨表';
var SHEET_NAME_TF_RECORDS = '轉換紀錄';
var SHEET_NAME_TF_BACKUP  = '工單備份';

// 讀取妹妹配對 + 經濟人清單（供轉檔器初始化）
function getInitData() {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_MMGG);
    var data = sh.getDataRange().getValues();
    var mmgg = {};
    data.forEach(function(row) {
      if (row[0] && String(row[0]) !== '妹妹名稱') mmgg[String(row[0])] = String(row[1] || '');
    });
    // 從 KV 讀取 ggList
    var shK = getOrCreateSheet(SHEET_NAME_KV);
    var kData = shK.getDataRange().getValues();
    var ggList = [];
    kData.forEach(function(row) {
      if (String(row[0]) === 'ggList') { try { ggList = JSON.parse(String(row[1])); } catch(e) {} }
    });
    return { mmgg: mmgg, ggList: ggList };
  } catch(e) { return { mmgg: {}, ggList: [] }; }
}

// 儲存妹妹配對 + 經濟人清單
function updateMapping(mmgg, ggList) {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_MMGG);
    sh.clearContents();
    sh.appendRow(['妹妹名稱', '經濟人']);
    Object.keys(mmgg || {}).forEach(function(mm) {
      if (mm) sh.appendRow([mm, mmgg[mm] || '']);
    });
    // ggList 存到 KV
    var shK = getOrCreateSheet(SHEET_NAME_KV);
    var kData = shK.getDataRange().getValues();
    var ggVal = JSON.stringify(ggList || []);
    var found = false;
    for (var i = 0; i < kData.length; i++) {
      if (String(kData[i][0]) === 'ggList') { shK.getRange(i+1,2).setValue(ggVal); found=true; break; }
    }
    if (!found) shK.appendRow(['ggList', ggVal]);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// 轉換結果 + 原始備份 寫入 Sheets
function saveToSheet(records, rawText, dateInput) {
  try {
    var shR = getOrCreateSheet(SHEET_NAME_TF_RECORDS);
    if (shR.getLastRow() === 0) shR.appendRow(['日期','阿姨','妹妹','GG','時間','地點','PS','寫','牌價','M收','建立時間']);
    var now = new Date().toLocaleString('zh-TW');
    (records || []).forEach(function(r) {
      shR.appendRow([r.date||'',r.ee||'',r.mm||'',r.gg||'',r.time||'',r.place||'',r.psCol||'',r.writeCol||'',r.price||'',r.shou||'',now]);
    });
    var shB = getOrCreateSheet(SHEET_NAME_TF_BACKUP);
    if (shB.getLastRow() === 0) shB.appendRow(['日期','備份時間','筆數','原始工單']);
    shB.appendRow([dateInput||'', now, (records||[]).length, rawText||'']);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// 讀取歷史備份（最新 20 筆）
function getRawOrdersHistory() {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_TF_BACKUP);
    var data = sh.getDataRange().getValues();
    if (data.length <= 1) return [];
    var result = [];
    for (var i = data.length-1; i >= 1 && result.length < 20; i--) {
      result.push({ date: String(data[i][0]), timestamp: String(data[i][1]), count: data[i][2], raw: String(data[i][3]), weekTab: '備份 #'+(data.length-i) });
    }
    return result;
  } catch(e) { return []; }
}

// 讀取阿姨表
function getAuntieList() {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_AUNTIE);
    var data = sh.getDataRange().getValues();
    var result = [];
    data.forEach(function(row) {
      if (row[0] && String(row[0]) !== '阿姨名稱') result.push({ name: String(row[0]), type: String(row[1]||''), bonus: Number(row[2])||0 });
    });
    return result;
  } catch(e) { return []; }
}

// 儲存阿姨表
function saveAuntieList(list) {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_AUNTIE);
    sh.clearContents();
    sh.appendRow(['阿姨名稱', '類型(JP/TW)', '退補']);
    (list || []).forEach(function(a) { if(a.name) sh.appendRow([a.name, a.type||'', a.bonus||0]); });
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// ==================== Admin.html 相容函式 ====================
// Admin.html 使用 save(key, val) 單筆寫入班表資料
function save(key, val) {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_KV);
    var data = sh.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(key)) {
        sh.getRange(i + 1, 2).setValue(val);
        return { success: true };
      }
    }
    sh.appendRow([key, val]);
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// Admin.html 使用 loadAll() 讀取所有班表 KV 資料
function loadAll() {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_KV);
    var data = sh.getDataRange().getValues();
    var result = { day: '', week: '', note: '', db: '', walk: '' };
    data.forEach(function(row) {
      var k = String(row[0]);
      if (k in result) result[k] = String(row[1] || '');
    });
    return result;
  } catch (err) {
    return { day: '', week: '', note: '', db: '', walk: '', error: err.message };
  }
}
