# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 專案概述

這是一個台灣美容業的**班表管理系統**，以 Google Apps Script (GAS) Web App 形式部署，搭配 Google Sheets 作為後端資料庫。

## 檔案結構

```
gas/
  Code.gs       # GAS 後端：讀寫 Google Sheets 的所有 API
  index.html    # 前端 UI（單一 HTML 檔含所有 CSS + JS）
班表3N.html     # 原始獨立版本（localStorage 儲存，無 GAS）
```

## 部署方式

- **`gas/` 目錄**是要上傳到 Google Apps Script 的程式碼
- 在 [script.google.com](https://script.google.com) 建立專案，把 `Code.gs` 和 `index.html` 的內容貼進去
- 部署為 Web App（執行身份：擁有者，存取：所有人）
- **不能使用 `getActiveSpreadsheet()`**，Web App 環境必須用 `SpreadsheetApp.openById(SPREADSHEET_ID)`

## Google Sheets 結構

Spreadsheet ID：`1DUx0mfmoSuW10UnydNJl_gNMoWZXQzmTER0KHu5NPvU`

| Sheet 名稱 | 用途 | 存取方式 |
|-----------|------|---------|
| 當日排班   | 當日班表文字 | `readCell_` / `writeCell_`（A1 單格） |
| 週報      | 週排班原文 | `readCell_` / `writeCell_` |
| 自由編輯區 | 筆記本 | `readCell_` / `writeCell_` |
| 母檔      | 員工資料庫（可超 50k 字） | `readChunked_` / `writeChunked_`（分塊） |
| 現走表    | 四款預設設定 JSON | `readCell_` / `writeCell_` |
| 設定      | 系統設定 | `readCell_` / `writeCell_` |

## 資料存取關鍵邏輯

**母檔分塊存取（Chunked Storage）**：Google Sheets 單一 cell 上限 50,000 字，母檔資料庫可能超過此限制。`writeChunked_` 以 45,000 字為單位分割存入 A1、A2、A3…；`readChunked_` 讀取全欄後 join 還原。

**自動存檔策略**：
- `autoSave(key, val)`：1.5 秒 debounce，使用者停止輸入後才送出
- `immediateSave(key, val)`：立即送出，用於 `onblur` 事件（防止關頁面前未存）
- 母檔 textarea 同時掛 `oninput="autoSave(...)"` 和 `onblur="immediateSave(...)"`

**Client → Server 呼叫**：
```js
google.script.run
  .withSuccessHandler(fn)
  .withFailureHandler(fn)
  .save(key, val)   // 或 .loadAll()
```

## 前端架構（index.html）

兩個 Tab：
- **班表管理**（`#tab-schedule`）：當日排班轉換、週報解析、7 日看板、母檔資料庫
- **現走表**（`#tab-walk`）：廣告文案四款同時轉檔（A/B/C/D 並排四欄輸出）

**現走表輸出 HTML 結構**：
```html
<div class="walk-4col-grid">  <!-- 四欄等寬 grid -->
  <div class="walk-out-card">  <!-- A款 -->
  <div class="walk-out-card">  <!-- B款 -->
  <div class="walk-out-card">  <!-- C款 -->
  <div class="walk-out-card">  <!-- D款 -->
</div>
```

**Walk 四款預設**：`WALK_DEFAULTS` 陣列含 4 個 preset 物件，欄位有 `symTitle / symDot / symMid / symJp / symTw / timePrefix / timeSuffix / symGroupSep / symStatusLine / symComing`。存入 Sheets 格式為 JSON：`{ presets: [...], quickRest: "...", quickFull: "..." }`。

**母檔解析**：`parseToMap(raw)` 以雙換行分割 block，用 `🤍(\d+)` 取得價格、`『([^』]+)』` 取得姓名。

## 視覺規範

- **配色**：Morandi 莫蘭迪風格，CSS variables 定義在 `:root` 和 `html.dark`
- **主色調**：奶油色 `#FCFAF8`（淺）/ `#2A2420`（深）
- **深淺切換**：`toggleDark()` 切換 `html.dark` class，狀態存 `localStorage`
- **字型**：Cormorant Garamond（標題）、Noto Sans TC / Noto Serif TC（中文）、JetBrains Mono（輸出框）
