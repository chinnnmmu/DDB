// ==================== DDB GAS Web App 後端 ====================

var SHEET_NAME_NOTES    = '便簽';
var SHEET_NAME_ORDERS   = '工單';
var SHEET_NAME_META     = '同步紀錄';
var SHEET_NAME_ADMIN    = '格式設定';
var SHEET_NAME_KV       = '班表資料';

// ---- 入口點 ----
function doGet(e) {
  var page = e && e.parameter && e.parameter.page === 'admin' ? 'Admin' : 'Index';
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle(page === 'Admin' ? '班表管理系統' : 'DDB 工作台')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== System 2: save(key, val) ====================
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

// ==================== System 2: loadAll() ====================
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

// ---- 工具：取得或建立 Sheet ----
function getOrCreateSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

// ==================== System 1: saveAllData ====================
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

    // 3. 記錄同步時間
    var shM = getOrCreateSheet(SHEET_NAME_META);
    shM.clearContents();
    shM.appendRow(['最後同步', new Date().toLocaleString('zh-TW')]);

    return { success: true, message: '同步成功' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ==================== System 1: loadAllData ====================
function loadAllData() {
  try {
    var result = { multiNotes: [], workOrders: {} };

    // 1. 讀取便簽（依欄位順序讀，不依標題）
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

    // 2. 讀取工單（依欄位順序讀，不依標題）
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

    return result;
  } catch (err) {
    return { multiNotes: [], workOrders: {}, error: err.message };
  }
}

// ==================== System 2: saveAdminData ====================
function saveAdminData(payload) {
  try {
    if (typeof payload === 'string') payload = JSON.parse(payload);

    var shA = getOrCreateSheet(SHEET_NAME_ADMIN);
    shA.clearContents();
    shA.appendRow(['設定名稱', '設定值']);

    if (payload.schDB && typeof payload.schDB === 'string') {
      shA.appendRow(['schDB', payload.schDB]);
    }
    if (payload.formatPresets && Array.isArray(payload.formatPresets)) {
      shA.appendRow(['formatPresets', JSON.stringify(payload.formatPresets)]);
    }

    var shM = getOrCreateSheet(SHEET_NAME_META);
    var metaData = shM.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < metaData.length; i++) {
      if (metaData[i][0] === 'adminSync') {
        shM.getRange(i + 1, 2).setValue(new Date().toLocaleString('zh-TW'));
        found = true; break;
      }
    }
    if (!found) shM.appendRow(['adminSync', new Date().toLocaleString('zh-TW')]);

    return { success: true, message: '管理資料同步成功' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ==================== System 2: loadAdminData ====================
function loadAdminData() {
  try {
    var result = { schDB: '', formatPresets: [] };

    var shA = getOrCreateSheet(SHEET_NAME_ADMIN);
    var aData = shA.getDataRange().getValues();
    for (var k = 1; k < aData.length; k++) {
      var key = String(aData[k][0]);
      if (key === 'schDB') {
        result.schDB = String(aData[k][1]);
      }
      if (key === 'formatPresets') {
        try { result.formatPresets = JSON.parse(aData[k][1]); } catch(e) {}
      }
    }

    return result;
  } catch (err) {
    return { schDB: '', formatPresets: [], error: err.message };
  }
}
