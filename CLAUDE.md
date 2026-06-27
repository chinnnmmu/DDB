# DDB 工作台 — 系統說明

## 專案概述
Google Apps Script (GAS) Web App，單一檔案 `Index.html`，用 `HtmlService` 提供服務。
無框架，純原生 JS + HTML + CSS。部署後是手機優先的 PWA 介面。

## 檔案結構
- `Index.html` — 唯一前端檔案（4000+ 行），包含所有 CSS / HTML / JS
- `Code.gs` — GAS 後端（讀寫 Google Sheets、儲存資料）
- `CLAUDE.md` — 本說明文件

---

## 主要分頁（底部導覽列）

| tab id | 名稱 | 說明 |
|--------|------|------|
| `tab-cal` | 工單 | 每日工單主介面（行事曆 + 快速輸入） |
| `tab-all` | 所有工單 | 全部工單列表，可篩選排序 |
| `tab-bb` | 班表 | 班表管理，解析班表文字 |
| `tab-stat` | 統計 | 阿姨應收、結清功能 |
| `tab-tf` | 轉檔 | 工單轉換輸出（舊功能，部分已整合） |
| `tab-boss` | 老闆帳目 | 密碼鎖，v4 帳務系統（老闆專用） |
| `tab-walk` | 班表管理 | 班表自由編輯區 |

切換用 `ST(tabId)` 函數。

---

## 核心資料結構

### 工單 `notes`
```js
notes = {
  "2025/06/24": [
    {
      id: "r...",
      pai: "16",        // 牌價（×10 = 實際價，如 16 = 160）
      name: "莫凡",     // 妹妹人名
      place: "2200香格里拉",  // 時間地點
      auntie: "加藤",   // 阿姨名
      shi: "320",       // 實收
      rebate: "",       // 退款
      ee: "128",        // EE（員工費，系統計算）
      note: "V",        // 備記
      settled: false,
      settledDate: "",
      settledAmount: ""
    }
  ]
}
```
- 日期 key 格式：`YYYY/MM/DD`（由 `DK(y,m,d)` 產生）
- 儲存到 localStorage `ddb-work-orders` 及 GAS Sheets
- `save()` 儲存，`RR()` 重繪工單列表，`RC()` 重繪日曆

### EE 計算公式（員工費）
```js
// calculateEE(pai, shi, rebate)
EE = shi - (pai × 6) + rebate
// pai=16, shi=320, rebate=0 → EE = 320 - 96 = 224（不對，要看實際）
// pai 直接輸入數字（不×10），內部 ×6 計算基礎費
```

---

## 老闆帳目 v4 系統（tab-boss）

密碼儲存在 localStorage `boss-pw`，預設 `0000`。
解鎖後顯示 `#boss-content`，隱藏 `#boss-lock`。

### v4 資料狀態
```js
var v4_rawData = [];   // 從 notes 計算出的帳務資料
var v4_payments = [];  // 付款紀錄
var v4_db = {
  agents: {},          // 經紀（GG）統計
  aunties: {},         // 阿姨統計
  mmgg: {},            // 妹妹人名 → GG 名（持久）
  auntieBonus: {},     // 阿姨名 → 額外退款數（+3/+5等）（持久）
  wallets: ['我','男友','玉山銀行'],
  walletCash: {},
  ledgers: { '司機':[], '家人':[], '現金':[], '其他':[] },
  activeLedger: '司機',
  // ...其他統計欄位
};
```

### v4 帳務公式
```js
// actualP = pai × 10（牌價，如 pai=16 → actualP=160）
// ggRefund = V4_REFUND_TABLE 查表 + auntieBonus[阿姨名]
// 若 actualP === 0（無牌）：
//   G回(back) = shi, EE$(ee_pay) = shi, S = 0
// 若 actualP > 0：
//   G回(back)    = shi - (actualP×0.5 + ggRefund)
//   EE$(ee_pay)  = shi - actualP×0.6
//   S(s_profit)  = actualP×0.1 - ggRefund
```

### V4_REFUND_TABLE（GG 退款對照表）
```js
// [牌價下限, 牌價上限, GG退款]
[50,100,3],[110,140,4],[150,190,5],[200,240,6],[250,290,7],
[300,340,10],[350,390,12],[400,450,15],[460,500,18],[510,590,20],
[600,690,20],[700,790,23],[800,890,26],[900,990,30],[1000,1100,35],
[1110,1200,40],[1210,1300,45]
```

### v4 重要函數
- `v4_generateRawData()` — 從 `notes` 重新計算所有帳務資料
- `v4_analyzeData()` — 統計分析
- `v4_renderPage(page)` — 渲染 v4 頁面（'company'/'agent'/'aunty'/'payment'/'data'/'settings'）
- `v4_saveToLocal()` — 儲存到 localStorage `boss_v4_data`（或 GAS `saveBossData`）
- `v4_loadData(cb)` — 載入（或 GAS `loadBossData`）

### v4 員工不可見欄位
員工看得到：牌、人名、地點、阿姨、實收
員工**看不到**（老闆密碼後才顯示）：G回、S、EE$、公司利潤

---

## 智能解析工單

### 格式一（空格格式）
```
16莫凡 2200香格里拉2+1 加藤(V)320 128
```
解析：牌16、人名莫凡、地點2200香格里拉2+1、阿姨加藤、實收320、EE128、備記V

### 格式二（竹葉格式）
```
38咪露｜1330｜三重江月3｜小陳｜1140/456
```
解析：牌38、人名咪露、時間1330、地點三重江月3、阿姨小陳、實收1140、EE456

### 相關函數
- `parseImpLine(line)` — 單行智能解析，回傳 row 物件
- `submitQaSmart()` — 智能輸入框 Enter 新增（在 `#qa-smart` input）
- `doImport()` — 批次多行匯入（`#imp-raw` textarea）

---

## 主要 JS 函數索引

| 函數 | 說明 |
|------|------|
| `ST(tabId)` | 切換主分頁 |
| `SD(y,m,d)` | 選擇日期 |
| `selKey()` | 取得目前日期的 notes key |
| `save()` | 儲存 notes 到 localStorage + 雲端 |
| `RR()` | 重繪工單列表 |
| `RC()` | 重繪月曆 |
| `RW()` | 重繪週統計 |
| `addRow()` | 新增一筆空白工單 |
| `submitQa(cont)` | 分欄快速輸入新增 |
| `submitQaSmart()` | 智能單行輸入新增 |
| `parseImpLine(line)` | 智能解析一行文字 |
| `calculateEE(pai,shi,rebate)` | 計算 EE 員工費 |
| `bossUnlock()` | 老闆帳目密碼解鎖 |
| `v4_generateRawData()` | v4 從工單重算帳務 |
| `v4_renderPage(p)` | v4 渲染指定頁面 |

---

## CSS 架構

- CSS 變數定義在 `:root`（亮色預設）
- 深色主題：`body.dark-theme`
- 淺色主題：`body.light-theme`
- 暖色主題：`body.warm-theme`
- v4 老闆帳目 CSS 全部用 `--v4-` 前綴變數，scoped 在 `#tab-boss` 下，避免衝突

---

## 開發注意事項

1. **不要直接輸出整個 Index.html** — 檔案太大，用 Edit 工具修改特定區塊
2. **v4 函數全部加 `v4_` 前綴** — 避免與主系統衝突
3. **v4 CSS class 全部加 `v4-` 前綴**
4. **GAS 環境偵測**：`var IS_GAS = typeof google !== 'undefined' && google.script`
5. **branch**：`claude/amazing-bell-el40qn`
6. **push 後給使用者下載 Index.html** 讓她貼到 GAS

---

## 常見需求對應

| 需求 | 修改位置 |
|------|---------|
| 改 EE 計算公式 | `calculateEE()` 約第 1950 行 |
| 改 GG 退款表 | `V4_REFUND_TABLE` 約第 4270 行 |
| 改老闆帳目介面 | `#tab-boss` HTML + `v4_render*` JS 函數 |
| 改工單快速輸入 | `#quick-add-wrap` HTML + `submitQa()` |
| 改智能解析規則 | `parseImpLine()` 約第 2943 行 |
| 新增工單欄位 | notes 物件 + `RR()` 渲染 + `submitQa()` 收集 |
