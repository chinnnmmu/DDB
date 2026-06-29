// ==================== DDB GAS 完整版本 v4 ====================

// ==================== Sheet 名稱定義（中文） ====================
var SHEET_NAME_NOTES  = '便簽';
var SHEET_NAME_ORDERS = '工單';
var SHEET_NAME_META   = '同步紀錄';
var SHEET_NAME_KV     = '班表資料';
var SHEET_NAME_MMGG   = '妹妹配對';
var SHEET_NAME_AUNTIE = '阿姨表';
var SHEET_NAME_TF_RECORDS = '轉換紀錄';
var SHEET_NAME_TF_BACKUP  = '工單備份';
var SHEET_NAME_BOSS   = '老闆帳目';
var SHEET_NAME_REPORTS = '報表紀錄';

// ==================== 入口點 ====================
function doGet(e) {
  var page = e && e.parameter && e.parameter.page === 'admin' ? 'Admin' : 'Index';
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle(page === 'Admin' ? '班表管理系統' : 'DDB 工作台')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== 工具：日期格式化 YYYY/MM/DD ====================
// Google Sheets getValues() 回傳日期欄位時是 Date 物件，需轉換
function formatDateKey(val) {
  if (!val) return '';
  var d = (val instanceof Date) ? val : new Date(val);
  if (isNaN(d.getTime())) return String(val);
  var y = d.getFullYear();
  var m = d.getMonth() + 1;
  var day = d.getDate();
  return y + '/' + (m < 10 ? '0' : '') + m + '/' + (day < 10 ? '0' : '') + day;
}

// ==================== 獲取或建立 Sheet ====================
function getOrCreateSheet(name) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    return sh;
  } catch (e) {
    throw new Error('無法獲取或建立 Sheet: ' + name + ' (' + e.message + ')');
  }
}

// ==================== 初始化所有 Sheet ====================
function initializeAllSheets() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetNames = [
      SHEET_NAME_NOTES, SHEET_NAME_ORDERS, SHEET_NAME_META, SHEET_NAME_KV,
      SHEET_NAME_MMGG, SHEET_NAME_AUNTIE, SHEET_NAME_TF_RECORDS, SHEET_NAME_TF_BACKUP, SHEET_NAME_BOSS
    ];
    sheetNames.forEach(function(name) {
      if (!ss.getSheetByName(name)) ss.insertSheet(name);
    });
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==================== 自動清理和初始化 ====================
function cleanupAndInitialize() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var keepSheets = ['便簽', '工單', '班表資料', '同步紀錄', '妹妹配對', '阿姨表', '轉換紀錄', '工單備份', '老闆帳目'];
    ss.getSheets().forEach(function(sheet) {
      if (keepSheets.indexOf(sheet.getName()) === -1) {
        try { ss.deleteSheet(sheet); } catch (e) {}
      }
    });
    keepSheets.forEach(function(name) {
      if (!ss.getSheetByName(name)) ss.insertSheet(name);
    });
    return { success: true, message: '✅ 清理完成', sheets: keepSheets };
  } catch (err) {
    return { success: false, message: '❌ 清理失敗: ' + err.message };
  }
}

// ==================== 核心：保存所有資料 ====================
function saveAllData(payload) {
  try {
    if (typeof payload === 'string') {
      try { payload = JSON.parse(payload); }
      catch (parseErr) { return { success: false, message: 'JSON 解析失敗: ' + parseErr.message }; }
    }
    if (!payload) return { success: false, message: 'payload 為空' };

    initializeAllSheets();
    var timestamp = new Date().toLocaleString('zh-TW');

    // 1. 保存便簽
    if (payload.multiNotes && Array.isArray(payload.multiNotes)) {
      try {
        var sh = getOrCreateSheet(SHEET_NAME_NOTES);
        sh.clearContents();
        sh.appendRow(['編號', '時間', '內容']);
        payload.multiNotes.forEach(function(n) {
          if (n) sh.appendRow([n.id || '', n.time || '', n.text || '']);
        });
      } catch (e) { console.error('保存便簽失敗:', e); }
    }

    // 2. 保存工單（僅在有資料時才清除覆蓋）
    if (payload.workOrders && typeof payload.workOrders === 'object') {
      var woKeys = Object.keys(payload.workOrders || {});
      if (woKeys.length > 0) {
        try {
          var shO = getOrCreateSheet(SHEET_NAME_ORDERS);
          shO.clearContents();
          shO.appendRow(['日期','牌','人名','時間地點','阿姨','實收','EE','退','備記','已結','結帳日','結帳金額','編號']);
          woKeys.forEach(function(dateKey) {
            var rows = payload.workOrders[dateKey];
            if (Array.isArray(rows)) {
              rows.forEach(function(r) {
                if (r) {
                  shO.appendRow([
                    dateKey,
                    r.pai || '', r.name || '', r.place || '',
                    r.auntie || '', r.shi || '', r.ee || '', r.rebate || '',
                    r.note || '', r.settled ? 'Y' : 'N',
                    r.settledDate || '', r.settledAmount || '', r.id || ''
                  ]);
                }
              });
            }
          });
        } catch (e) { console.error('保存工單失敗:', e); }
      }
    }

    // 3. 保存班表 KV 資料
    if (payload.kvData && typeof payload.kvData === 'object') {
      try {
        var shK = getOrCreateSheet(SHEET_NAME_KV);
        var data = shK.getDataRange().getValues() || [];
        var kvMap = {};
        for (var i = 1; i < data.length; i++) {
          if (data[i] && data[i][0]) kvMap[String(data[i][0])] = i + 1;
        }
        Object.keys(payload.kvData).forEach(function(key) {
          if (!key) return;
          var val = String(payload.kvData[key] || '');
          if (kvMap[key]) shK.getRange(kvMap[key], 2).setValue(val);
          else shK.appendRow([key, val]);
        });
      } catch (e) { console.error('保存 KV 失敗:', e); }
    }

    // 4. 記錄同步時間
    try {
      var shM = getOrCreateSheet(SHEET_NAME_META);
      shM.clearContents();
      shM.appendRow(['項目', '值']);
      shM.appendRow(['最後同步時間', timestamp]);
      shM.appendRow(['狀態', '成功']);
    } catch (e) {}

    return { success: true, message: '資料已保存', timestamp: timestamp };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ==================== 核心：讀取所有資料 ====================
function loadAllData() {
  try {
    initializeAllSheets();
    var result = { multiNotes: [], workOrders: {}, kvData: {} };

    // 1. 讀取便簽
    try {
      var sh = getOrCreateSheet(SHEET_NAME_NOTES);
      var data = sh.getDataRange().getValues() || [];
      if (data.length > 1) {
        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          if (row && (row[0] || row[2])) {
            result.multiNotes.push({
              id: String(row[0] || ''), time: String(row[1] || ''), text: String(row[2] || '')
            });
          }
        }
      }
    } catch (e) { console.error('讀取便簽失敗:', e); }

    // 2. 讀取工單（若工單空則從轉換紀錄自動重建）
    try {
      var shO = getOrCreateSheet(SHEET_NAME_ORDERS);
      var oData = shO.getDataRange().getValues() || [];
      if (oData.length <= 1) {
        // 工單空 → 嘗試從轉換紀錄重建
        var shR2 = getOrCreateSheet(SHEET_NAME_TF_RECORDS);
        var rData2 = shR2.getDataRange().getValues() || [];
        if (rData2.length > 1) {
          var tfRecs = [];
          for (var ti = 1; ti < rData2.length; ti++) {
            var tr = rData2[ti];
            if (tr) tfRecs.push({
              date: formatDateKey(tr[0]), ee: String(tr[1]||''), mm: String(tr[2]||''),
              gg: String(tr[3]||''), time: String(tr[4]||''), place: String(tr[5]||''),
              psCol: String(tr[6]||''), writeCol: String(tr[7]||''),
              price: String(tr[8]||''), shou: String(tr[9]||'')
            });
          }
          _mergeTfRecordsToOrders(tfRecs);
          oData = shO.getDataRange().getValues() || [];
        }
      }
      if (oData.length > 1) {
        for (var j = 1; j < oData.length; j++) {
          var r = oData[j];
          if (!r) continue;
          // 日期欄可能是 Date 物件，用 formatDateKey 轉換
          var dateKey = formatDateKey(r[0]);
          if (!dateKey) continue;
          if (!result.workOrders[dateKey]) result.workOrders[dateKey] = [];
          result.workOrders[dateKey].push({
            pai: String(r[1]||''), name: String(r[2]||''), place: String(r[3]||''),
            auntie: String(r[4]||''), shi: String(r[5]||''), ee: String(r[6]||''),
            rebate: String(r[7]||''), note: String(r[8]||''),
            settled: r[9] === 'Y',
            settledDate: String(r[10]||''), settledAmount: String(r[11]||''),
            id: String(r[12]||'')
          });
        }
      }
    } catch (e) { console.error('讀取工單失敗:', e); }

    // 3. 讀取 KV 資料
    try {
      var shK = getOrCreateSheet(SHEET_NAME_KV);
      var kData = shK.getDataRange().getValues() || [];
      for (var k = 1; k < kData.length; k++) {
        if (kData[k] && kData[k][0]) {
          result.kvData[String(kData[k][0])] = String(kData[k][1] || '');
        }
      }
    } catch (e) { console.error('讀取 KV 失敗:', e); }

    return result;
  } catch (err) {
    return { multiNotes: [], workOrders: {}, kvData: {}, error: err.message };
  }
}

// ==================== 兼容舊版 API ====================
function save(key, val) {
  try {
    var shK = getOrCreateSheet(SHEET_NAME_KV);
    var data = shK.getDataRange().getValues() || [];
    for (var i = 1; i < data.length; i++) {
      if (data[i] && String(data[i][0]) === String(key)) {
        shK.getRange(i + 1, 2).setValue(val);
        return { success: true };
      }
    }
    shK.appendRow([key, val]);
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function loadAll() {
  try {
    var shK = getOrCreateSheet(SHEET_NAME_KV);
    var data = shK.getDataRange().getValues() || [];
    var result = {};
    for (var i = 1; i < data.length; i++) {
      if (data[i] && data[i][0]) result[String(data[i][0])] = String(data[i][1] || '');
    }
    return result;
  } catch (err) { return { error: err.message }; }
}

// ==================== 妹妹配對管理 ====================
function getInitData() {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_MMGG);
    var data = sh.getDataRange().getValues() || [];
    var mmgg = {};
    data.forEach(function(row) {
      if (row && row[0] && String(row[0]) !== '妹妹名稱') mmgg[String(row[0])] = String(row[1] || '');
    });
    var ggList = [];
    try {
      var kvData = loadAllData().kvData;
      if (kvData && kvData.ggList) ggList = JSON.parse(kvData.ggList);
    } catch (e) {}
    return { mmgg: mmgg, ggList: ggList };
  } catch (e) { return { mmgg: {}, ggList: [] }; }
}

function updateMapping(mmgg, ggList) {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_MMGG);
    sh.clearContents();
    sh.appendRow(['妹妹名稱', '經濟人']);
    if (mmgg && typeof mmgg === 'object') {
      Object.keys(mmgg).forEach(function(mm) { if (mm) sh.appendRow([mm, mmgg[mm] || '']); });
    }
    save('ggList', JSON.stringify(ggList || []));
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

// ==================== 阿姨表管理 ====================
function getAuntieList() {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_AUNTIE);
    var data = sh.getDataRange().getValues() || [];
    var result = [];
    data.forEach(function(row) {
      if (row && row[0] && String(row[0]) !== '阿姨名稱') {
        result.push({ name: String(row[0]), type: String(row[1]||''), bonus: Number(row[2])||0 });
      }
    });
    return result;
  } catch (e) { return []; }
}

function saveAuntieList(list) {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_AUNTIE);
    sh.clearContents();
    sh.appendRow(['阿姨名稱', '類型(JP/TW)', '退補']);
    if (Array.isArray(list)) {
      list.forEach(function(a) { if (a && a.name) sh.appendRow([a.name, a.type||'', a.bonus||0]); });
    }
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

// ==================== 轉檔器相關 ====================
function saveToSheet(records, rawText, dateInput) {
  try {
    var now = new Date().toLocaleString('zh-TW');

    // 1. 保存轉換紀錄
    var shR = getOrCreateSheet(SHEET_NAME_TF_RECORDS);
    if (shR.getLastRow() === 0) shR.appendRow(['日期','阿姨','妹妹','GG','時間','地點','PS','寫','牌價','M收','建立時間']);
    if (Array.isArray(records)) {
      records.forEach(function(r) {
        if (r) shR.appendRow([r.date||'', r.ee||'', r.mm||'', r.gg||'', r.time||'', r.place||'', r.psCol||'', r.writeCol||'', r.price||'', r.shou||'', now]);
      });
    }

    // 2. 保存備份
    var shB = getOrCreateSheet(SHEET_NAME_TF_BACKUP);
    if (shB.getLastRow() === 0) shB.appendRow(['日期','備份時間','筆數','原始工單']);
    shB.appendRow([dateInput||'', now, Array.isArray(records) ? records.length : 0, rawText||'']);

    // 3. 同步到工單
    if (Array.isArray(records)) _mergeTfRecordsToOrders(records);

    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function _mergeTfRecordsToOrders(records) {
  try {
    if (!Array.isArray(records) || records.length === 0) return;
    var shO = getOrCreateSheet(SHEET_NAME_ORDERS);
    if (shO.getLastRow() === 0) shO.appendRow(['日期','牌','人名','時間地點','阿姨','實收','EE','退','備記','已結','結帳日','結帳金額','編號']);

    var existing = shO.getDataRange().getValues() || [];
    var keySet = {};
    for (var i = 1; i < existing.length; i++) {
      if (existing[i] && existing[i].length >= 4) {
        var k = formatDateKey(existing[i][0]) + '|' + String(existing[i][2]||'') + '|' + String(existing[i][3]||'');
        keySet[k] = true;
      }
    }

    records.forEach(function(r) {
      if (!r) return;
      var placeStr = (r.time && r.place) ? (r.time + ' ' + r.place) : (r.time || r.place || '');
      var k = (r.date||'') + '|' + (r.mm||'') + '|' + placeStr;
      if (keySet[k]) return;
      keySet[k] = true;
      var pai = r.price ? String(parseInt(r.price) / 10) : '';
      var ee = (r.price && r.shou) ? String(parseInt(r.price) - parseInt(r.shou||0)) : '';
      var note = [r.psCol||'', r.gg ? 'GG:' + r.gg : ''].filter(Boolean).join(' ');
      shO.appendRow([
        r.date||'', pai, r.mm||'', placeStr,
        r.ee||'', r.shou||'', ee, r.writeCol||'',
        note, 'N', '', '', 'tf_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6)
      ]);
    });
  } catch (e) { console.error('_mergeTfRecordsToOrders 失敗:', e); }
}

function syncTfRecordsToOrders() {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_TF_RECORDS);
    var data = sh.getDataRange().getValues() || [];
    if (data.length <= 1) return { success: true, count: 0 };
    var records = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row) records.push({
        date: formatDateKey(row[0]), ee: String(row[1]||''), mm: String(row[2]||''),
        gg: String(row[3]||''), time: String(row[4]||''), place: String(row[5]||''),
        psCol: String(row[6]||''), writeCol: String(row[7]||''),
        price: String(row[8]||''), shou: String(row[9]||'')
      });
    }
    _mergeTfRecordsToOrders(records);
    return { success: true, count: records.length };
  } catch (e) { return { success: false, message: e.message }; }
}

// 讀取轉換紀錄（供前端可編輯表格用）
function getTfRecords() {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_TF_RECORDS);
    var data = sh.getDataRange().getValues() || [];
    if (data.length <= 1) return { success: true, records: [] };
    var records = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row) records.push({
        rowIndex: i + 1,
        date: formatDateKey(row[0]), ee: String(row[1]||''), mm: String(row[2]||''),
        gg: String(row[3]||''), time: String(row[4]||''), place: String(row[5]||''),
        psCol: String(row[6]||''), writeCol: String(row[7]||''),
        price: String(row[8]||''), shou: String(row[9]||''),
        created: String(row[10]||'')
      });
    }
    return { success: true, records: records };
  } catch (e) { return { success: false, message: e.message, records: [] }; }
}

// 更新單筆轉換紀錄
function updateTfRecord(rowIndex, record) {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_TF_RECORDS);
    var range = sh.getRange(rowIndex, 1, 1, 10);
    range.setValues([[
      record.date||'', record.ee||'', record.mm||'', record.gg||'',
      record.time||'', record.place||'', record.psCol||'', record.writeCol||'',
      record.price||'', record.shou||''
    ]]);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

// 刪除轉換紀錄行
function deleteTfRecord(rowIndex) {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_TF_RECORDS);
    sh.deleteRow(rowIndex);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function getRawOrdersHistory() {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_TF_BACKUP);
    var data = sh.getDataRange().getValues() || [];
    if (data.length <= 1) return [];
    var result = [];
    for (var i = Math.max(1, data.length - 20); i < data.length; i++) {
      if (data[i]) result.push({
        date: String(data[i][0]||''), timestamp: String(data[i][1]||''),
        count: data[i][2]||0, raw: String(data[i][3]||''),
        weekTab: '備份 #' + (data.length - i)
      });
    }
    return result;
  } catch (e) { return []; }
}

// ==================== 老闆帳目 ====================
function saveBossData(jsonStr) {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_BOSS);
    sh.clearContents();
    sh.appendRow(['key', 'value']);
    sh.appendRow(['data', jsonStr]);
    sh.appendRow(['updated', new Date().toLocaleString('zh-TW')]);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function loadBossData() {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_BOSS);
    var rows = sh.getDataRange().getValues() || [];
    for (var i = 0; i < rows.length; i++) {
      if (rows[i] && String(rows[i][0]) === 'data') return String(rows[i][1] || '');
    }
    return '';
  } catch (e) { return ''; }
}

// ==================== 診斷工具 ====================
function healthCheck() {
  try {
    var result = { sheets: {}, timestamp: new Date().toLocaleString('zh-TW'), errors: [] };
    [SHEET_NAME_NOTES, SHEET_NAME_ORDERS, SHEET_NAME_META, SHEET_NAME_KV,
     SHEET_NAME_MMGG, SHEET_NAME_AUNTIE, SHEET_NAME_TF_RECORDS, SHEET_NAME_TF_BACKUP, SHEET_NAME_BOSS
    ].forEach(function(name) {
      try {
        var sh = getOrCreateSheet(name);
        result.sheets[name] = { status: 'OK', rows: sh.getLastRow() };
      } catch (e) {
        result.sheets[name] = { status: 'ERROR', error: e.message };
        result.errors.push(name + ': ' + e.message);
      }
    });
    return result;
  } catch (err) { return { success: false, error: err.message }; }
}

function testWrite() {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_NOTES);
    sh.appendRow(['test-id', new Date().toLocaleString('zh-TW'), '測試寫入成功']);
    return { success: true, message: '寫入測試成功', lastRow: sh.getLastRow() };
  } catch (err) { return { success: false, message: err.message }; }
}

// ==================== 報表存檔 ====================
function saveReport(reportData) {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_REPORTS);
    if (sh.getLastRow() === 0) {
      sh.appendRow(['報表類型','起始日','結束日','GG總','EE總','S總','牌價總','M收總','筆數','存檔時間','明細JSON']);
    }
    var now = new Date().toLocaleString('zh-TW');
    sh.appendRow([
      reportData.type || '',
      reportData.dateFrom || '',
      reportData.dateTo || '',
      reportData.totalGG || 0,
      reportData.totalEE || 0,
      reportData.totalS || 0,
      reportData.totalPrice || 0,
      reportData.totalMRec || 0,
      reportData.count || 0,
      now,
      JSON.stringify(reportData.rows || [])
    ]);
    return { success: true, timestamp: now };
  } catch (e) { return { success: false, message: e.message }; }
}

function loadReports(limit) {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_REPORTS);
    var data = sh.getDataRange().getValues() || [];
    if (data.length <= 1) return { success: true, reports: [] };
    var n = limit || 30;
    var start = Math.max(1, data.length - n);
    var reports = [];
    for (var i = data.length - 1; i >= start; i--) {
      var r = data[i];
      reports.push({
        type: String(r[0]||''), dateFrom: String(r[1]||''), dateTo: String(r[2]||''),
        totalGG: r[3]||0, totalEE: r[4]||0, totalS: r[5]||0,
        totalPrice: r[6]||0, totalMRec: r[7]||0, count: r[8]||0,
        savedAt: String(r[9]||''), rowIndex: i+1
      });
    }
    return { success: true, reports: reports };
  } catch (e) { return { success: false, message: e.message, reports: [] }; }
}

function deleteReport(rowIndex) {
  try {
    var sh = getOrCreateSheet(SHEET_NAME_REPORTS);
    sh.deleteRow(rowIndex);
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}
