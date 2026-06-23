// ==================== DDB GAS Web App 後端 ====================
// 對應 BB4.HTML 前端使用的 API

var SHEET_NAME_NOTES   = 'MultiNotes';
var SHEET_NAME_ORDERS  = 'WorkOrders';
var SHEET_NAME_META    = 'Meta';

// ---- 入口點 ----
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('DDB 工作台')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---- 工具：取得或建立 Sheet ----
function getOrCreateSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

// ==================== saveAllData ====================
// payload = { action, multiNotes, workOrders }
// 前端呼叫：google.script.run.saveAllData(payload)
function saveAllData(payload) {
  try {
    if (typeof payload === 'string') payload = JSON.parse(payload);

    // 1. 儲存便簽 (MultiNotes)
    if (payload.multiNotes && Array.isArray(payload.multiNotes)) {
      var sh = getOrCreateSheet(SHEET_NAME_NOTES);
      sh.clearContents();
      sh.appendRow(['id', 'time', 'text']);
      payload.multiNotes.forEach(function(n) {
        sh.appendRow([n.id || '', n.time || '', n.text || '']);
      });
    }

    // 2. 儲存工單 (WorkOrders) — 扁平化所有日期下的工單
    if (payload.workOrders && typeof payload.workOrders === 'object') {
      var shO = getOrCreateSheet(SHEET_NAME_ORDERS);
      shO.clearContents();
      shO.appendRow(['date','pai','name','place','auntie','shi','ee','rebate','note','settled','settledDate','settledAmount','id']);
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
    shM.appendRow(['lastSync', new Date().toLocaleString('zh-TW')]);

    return { success: true, message: '同步成功' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ==================== loadAllData ====================
// 前端呼叫：google.script.run.loadAllData()
function loadAllData() {
  try {
    var result = { multiNotes: [], workOrders: {} };

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

    return result;
  } catch (err) {
    return { multiNotes: [], workOrders: {}, error: err.message };
  }
}
