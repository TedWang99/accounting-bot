# Accounting Bot 專案記憶

## 開發規範（使用者要求）

### 版號管理
1. 有程式碼變更時，適時調整版號。
2. 調整版號時，不直接修改原檔案：先複製成新版本（如 7.5.js → 7.6.js），再對新版檔案修改。
3. 舊版清理：每次修改後，超過 3 個版號以前的舊檔一律移至 `old/` 子資料夾保存（不刪除）。

### 部署方式
- 最新版位於 `8.0/` 資料夾（11 個 .gs 模組）
- 每次部署須在 `8.0/` 資料夾下執行 `clasp push`，否則會推舊版
- Git 已於 2026-03-03 初始化，之後每次修改皆應 commit

---

## 專案架構
- **平台**：Google Apps Script（GAS），部署為 LINE Webhook
- **資料庫**：Google Sheets（第一個工作表為主帳本）
- **AI**：Gemini 2.5 Flash（自然語言解析、分類、財務建議）
- **前端 Web App**：`index.html`（PWA，用 `doGet` 提供）

---

## 帳本欄位結構（Col 1–10）

| 欄 | 內容 |
|---|---|
| 1 | 日期 `yyyy/MM/dd` |
| 2 | 品名 |
| 3 | 金額 |
| 4 | 記錄人 |
| 5 | 支付方式 |
| 6 | 分類（`✈️日本-飲食` 格式，旅遊模式下帶前綴） |
| 7 | 專案（如 `🏠 一般開銷`、`✈️日本`、`👶 生寶寶／備孕`） |
| 8 | 發票號碼 |
| 9 | 來源（`LINE`、`Web`、`Web逐項`、`OCR圖片辨識` …） |
| 10 | 卡別（v7.7 新增） |

> **WebApp.js `writeToSheet()` 欄位名稱**：`date`, `item`, `amount`, `userName`, `payment`, `category`, `project`, `invoiceNum`, `source`, `cardId`

---

## 重要函數一覽

| 函數 | 檔案 | 說明 |
|---|---|---|
| `writeToSheet(record)` | Expense.js | 寫入一筆帳 |
| `canDeleteRecord(recordUser, currentUser)` | Expense.js | 刪除權限判斷 |
| `aggregateMonthlyData(data, month, year)` | Reports.js | 月份資料彙整 |
| `aggregateDateRangeData(data, startDate, endDate)` | Reports.js | 日期區間彙整 |
| `replyLine(replyToken, message, ...)` | Utils.js | LINE 回覆 |
| `pushLine(userId, message, ...)` | Utils.js | LINE 主動推播 |
| `fetchWithRetry(url, options, maxRetries)` | Utils.js | 帶重試的 HTTP |
| `checkDuplicate(recordObj)` | Expense.js | LINE Bot 重複偵測 |
| `webCheckDuplicate(token, date, amount)` | WebApp.js | Web App 重複偵測 |
| `buildSearchResultFlexMessage(keyword, result)` | FlexUI.js | 查詢結果卡片 |
| `handleEditLastEntry()` | Expense.js | 修改上一筆 |
| `_getActiveTravelProject()` | WebApp.js | 讀取旅遊模式專案（共用 AppProps） |
| `parseReceiptItemsWithGeminiVision(base64, mime)` | AI_Parse.js | 逐項掃描 OCR |
| `webSaveItemizedReceipt(token, items, commonData, force)` | WebApp.js | 批次儲存逐項明細 |

---

## CONFIG 重要設定

| Key | 儲存位置 | 說明 |
|---|---|---|
| `SHEET_ID` | ScriptProperties | 試算表 ID |
| `LINE_ACCESS_TOKEN` | ScriptProperties | LINE Token |
| `GEMINI_API_KEY` | ScriptProperties | Gemini API |
| `LINE_CHANNEL_SECRET` | ScriptProperties | LINE Secret |
| `USERS_MAP` | ScriptProperties | JSON `{"userId": "名稱"}` |
| `CARDS_MAP` | ScriptProperties | 信用卡設定 |
| `CATEGORY_BUDGETS` | ScriptProperties | 分類預算 |
| `travel_project` | ScriptProperties | 旅遊模式專案名稱 |
| `travel_end` | ScriptProperties | 旅遊結束日期 |
| `MONTHLY_BUDGET` | Config.js 常數 | 60000 |
| `DEFAULT_CATEGORY` | Config.js 常數 | `'其他'` |
| `DEFAULT_PROJECT` | Config.js 常數 | `'🏠 一般開銷'` |
