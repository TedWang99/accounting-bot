// ============================================
// WebApp - 家庭記帳本 PWA 後端
// ============================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('家庭記帳本')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── 登入（第一步：驗證帳密，推播 OTP）─────────
function webLogin(username, password) {
  try {
    const usersRaw = AppProps.getProperty('WEB_USERS');
    const users = usersRaw ? JSON.parse(usersRaw) : {};
    if (!users[username] || users[username] !== password) {
      return { success: false, error: '帳號或密碼錯誤' };
    }

    const lineIdsRaw = AppProps.getProperty('WEB_LINE_IDS');
    const lineIds = lineIdsRaw ? JSON.parse(lineIdsRaw) : {};
    const lineUserId = lineIds[username];
    if (!lineUserId) {
      return { success: false, error: '尚未設定 LINE 綁定，請聯絡管理員' };
    }

    const otp = String(Math.floor(100000 + Math.random() * 900000));
    CacheService.getScriptCache().put('otp_' + username, otp, 300);
    pushLine(lineUserId, `🔐 記帳本登入驗證碼\n\n${otp}\n\n5 分鐘內有效，請勿洩漏給他人。`);
    return { success: true, requireOtp: true };
  } catch (e) {
    return { success: false, error: '登入失敗：' + e.message };
  }
}

// ── 登入（第二步：驗證 OTP，發 token）──────────
function webVerifyOtp(username, otp) {
  try {
    const cached = CacheService.getScriptCache().get('otp_' + username);
    if (!cached) return { success: false, error: '驗證碼已過期，請重新登入' };
    if (cached !== otp.trim()) return { success: false, error: '驗證碼錯誤' };
    CacheService.getScriptCache().remove('otp_' + username);
    const token = Utilities.getUuid();
    CacheService.getScriptCache().put('web_' + token, username, 21600);
    return { success: true, token, name: username };
  } catch (e) {
    return { success: false, error: '驗證失敗：' + e.message };
  }
}

function webCheckToken(token) {
  const name = CacheService.getScriptCache().get('web_' + token);
  return name ? { valid: true, name } : { valid: false };
}

function webLogout(token) {
  CacheService.getScriptCache().remove('web_' + token);
  return { success: true };
}

// ── 取得月份資料（含 rowIndex 供刪除/修改用）──
function webGetMonthData(token, year, month) {
  const username = CacheService.getScriptCache().get('web_' + token);
  if (!username) return { success: false, error: '請重新登入' };

  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheets()[0];
    const data = sheet.getDataRange().getValues();

    const records = [];
    let totalExpense = 0;
    let totalIncome = 0;
    const categoryMap = {};
    const targetYear = parseInt(year);
    const targetMonth = parseInt(month);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      const dateVal = row[0];
      let rowDate;
      if (dateVal instanceof Date) {
        rowDate = dateVal;
      } else {
        const parts = String(dateVal).split('/');
        if (parts.length !== 3) continue;
        rowDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
      }

      if (rowDate.getFullYear() !== targetYear || (rowDate.getMonth() + 1) !== targetMonth) continue;

      const dateStr = Utilities.formatDate(rowDate, CONFIG.TIMEZONE, 'MM/dd');
      const item = String(row[1] || '');
      const amount = parseFloat(row[2]) || 0;
      const person = String(row[3] || '');
      const payment = String(row[4] || '');
      const category = String(row[5] || '');
      const project = String(row[6] || '');  // G欄
      const card = String(row[9] || '');

      records.push({ date: dateStr, item, amount, person, payment, category, project, card, rowIndex: i + 1 });

      if (amount >= 0) {
        totalExpense += amount;
        categoryMap[category] = (categoryMap[category] || 0) + amount;
      } else {
        totalIncome += Math.abs(amount);
      }
    }

    records.sort((a, b) => b.date.localeCompare(a.date));

    const categoryBreakdown = Object.entries(categoryMap)
      .map(([cat, amt]) => ({ category: cat, amount: amt }))
      .sort((a, b) => b.amount - a.amount)
      .slice(0, 8);

    return {
      success: true,
      records,
      summary: {
        totalExpense, totalIncome,
        count: records.length,
        budget: CONFIG.MONTHLY_BUDGET,
        remaining: CONFIG.MONTHLY_BUDGET - totalExpense
      },
      categoryBreakdown
    };
  } catch (e) {
    return { success: false, error: '取得資料失敗：' + e.message };
  }
}

// ── 取得年度每月合計（趨勢圖用）───────────────
function webGetYearSummary(token, year) {
  const username = CacheService.getScriptCache().get('web_' + token);
  if (!username) return { success: false, error: '請重新登入' };
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheets()[0];
    const data = sheet.getDataRange().getValues();
    const monthlyTotals = Array(12).fill(0);
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const dateVal = row[0];
      let rowDate;
      if (dateVal instanceof Date) { rowDate = dateVal; }
      else {
        const parts = String(dateVal).split('/');
        if (parts.length !== 3) continue;
        rowDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
      }
      if (rowDate.getFullYear() !== parseInt(year)) continue;
      const amount = parseFloat(row[2]) || 0;
      if (amount > 0) monthlyTotals[rowDate.getMonth()] += amount;
    }
    return { success: true, monthlyTotals };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ── [v8.0] 取得目前有效的旅遊模式專案（與 LINE Bot 共用 AppProps）
function _getActiveTravelProject() {
  try {
    const project = AppProps.getProperty('travel_project');
    if (!project) return null;
    const endStr = AppProps.getProperty('travel_end');
    if (endStr && new Date(endStr) < new Date()) {
      // 旅遊已結束，清除狀態
      AppProps.deleteProperty('travel_project');
      AppProps.deleteProperty('travel_end');
      AppProps.deleteProperty('travel_currency');
      AppProps.deleteProperty('travel_rate');
      return null;
    }
    return project; // e.g. "✈️ 日本"
  } catch (e) { return null; }
}

// ── 新增記帳（含卡別與重複偵測）──────────────────
// skipDupCheck=true 時跳過重複偵測（使用者確認後強制新增）
function webAddRecord(token, recordData, skipDupCheck) {
  const username = CacheService.getScriptCache().get('web_' + token);
  if (!username) return { success: false, error: '請重新登入' };

  try {
    const { item, amount, payment, category, date, card } = recordData;
    if (!item || !amount) return { success: false, error: '品名與金額為必填' };

    // ── 重複偵測（與新增在同一次 GAS 執行，使用字串日期比對）──
    if (!skipDupCheck && date) {
      const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheets()[0];
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        // 讀取全部資料列，但只取前 3 欄（日期、品名、金額），降低記憶體與 quota 消耗
        const sheetData = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
        const amt = parseFloat(amount);
        const targetStr = String(date).replace(/-/g, '/');
        for (let i = sheetData.length - 1; i >= 0; i--) {
          const row = sheetData[i];
          if (!row[0]) continue;
          const rowDateStr = _parseDateToStr(row[0]);
          if (!rowDateStr || rowDateStr.length < 10) continue;
          if (_daysBetweenStr(targetStr, rowDateStr) > 30) continue;
          const rowAmt = parseFloat(row[2]);
          if (Math.abs(rowAmt - amt) <= 1) {
            const displayDate = row[0] instanceof Date
              ? Utilities.formatDate(row[0], CONFIG.TIMEZONE, 'MM/dd')
              : rowDateStr.substring(5).replace('/', '/');
            return {
              success: false,
              duplicateFound: true,
              dup: { date: displayDate, item: String(row[1] || ''), amount: rowAmt, person: '' }
            };
          }
        }
      }
    }

    const expenseDate = date ? new Date(date + 'T12:00:00+08:00') : new Date();
    const dateStr = Utilities.formatDate(expenseDate, CONFIG.TIMEZONE, 'yyyy/MM/dd');

    const record = {
      date: dateStr,
      item: item.trim(),
      amount: parseFloat(amount),
      userName: username,
      payment: payment || CONFIG.DEFAULT_PAYMENT,
      category: CONFIG.CATEGORIES.includes(category) ? category : CONFIG.DEFAULT_CATEGORY,
      project: _getActiveTravelProject() || CONFIG.DEFAULT_PROJECT,  // 自動帶入旅遊模式專案
      invoiceNum: '',
      source: 'Web',
      cardId: card || ''
    };

    writeToSheet(record);
    return { success: true };
  } catch (e) {
    return { success: false, error: '新增失敗：' + e.message };
  }
}

// ── 掃描帳單圖片（Gemini Vision）────────────────
function webScanReceipt(token, base64, mimeType) {
  const username = CacheService.getScriptCache().get('web_' + token);
  if (!username) return { success: false, error: '請重新登入' };
  try {
    const parsed = parseReceiptWithGeminiVision(base64, mimeType || 'image/jpeg');
    if (!parsed || !parsed.amount) return { success: false, error: '無法識別收據，請手動輸入' };
    return { success: true, data: parsed };
  } catch (e) {
    return { success: false, error: '掃描失敗：' + e.message };
  }
}

// ── 掃描多頁帳單（Gemini Vision 多圖）───────────
function webScanReceiptMulti(token, imagesData) {
  const username = CacheService.getScriptCache().get('web_' + token);
  if (!username) return { success: false, error: '請重新登入' };
  try {
    const parsed = parseMultiPageReceiptWithGeminiVision(imagesData);
    if (!parsed || !parsed.amount) return { success: false, error: '無法識別收據，請手動輸入' };
    return { success: true, data: parsed };
  } catch (e) {
    return { success: false, error: '掃描失敗：' + e.message };
  }
}

// ── [v8.0] 逐項明細掃描（Gemini Vision 逐行解析）──────
function webScanReceiptItems(token, base64, mimeType) {
  const username = CacheService.getScriptCache().get('web_' + token);
  if (!username) return { success: false, error: '請重新登入' };
  try {
    const parsed = parseReceiptItemsWithGeminiVision(base64, mimeType || 'image/jpeg');
    if (!parsed || !parsed.items || parsed.items.length === 0) {
      return { success: false, error: '無法識別明細，請確認圖片清晰並重試' };
    }
    return { success: true, data: parsed };
  } catch (e) {
    return { success: false, error: '逐項掃描失敗：' + e.message };
  }
}

// ── [v8.0] 批次寫入逐項明細（確認後一次存 N 筆）──────
function webSaveItemizedReceipt(token, items, commonData, force) {
  const username = CacheService.getScriptCache().get('web_' + token);
  if (!username) return { success: false, error: '請重新登入' };
  if (!items || items.length === 0) return { success: false, error: '沒有要儲存的項目' };

  try {
    const { date, payment, card, store, currency } = commonData || {};
    const expenseDate = date ? new Date(date + 'T12:00:00+08:00') : new Date();
    const dateStr = Utilities.formatDate(expenseDate, CONFIG.TIMEZONE, 'yyyy/MM/dd');
    const finalPayment = payment || CONFIG.DEFAULT_PAYMENT;
    // 自動套用旅遊模式專案（與 LINE Bot 共用）
    const finalProject = _getActiveTravelProject() || CONFIG.DEFAULT_PROJECT;
    const isForex = currency && currency !== 'TWD';

    // ── 重複偵測：同日同店舖已有 3 筆以上就警告 ──
    if (!force && store) {
      const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheets()[0];
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const sheetData = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // [日期, 品名]
        const targetDate = dateStr.replace(/-/g, '/');
        let matchCount = 0;
        for (let i = 0; i < sheetData.length; i++) {
          const rowDate = _parseDateToStr(sheetData[i][0]);
          const rowItem = String(sheetData[i][1] || '');
          if (rowDate !== targetDate) continue;
          // 同店名前綴（含旅遊前綴的情境）
          if (rowItem.includes(store)) matchCount++;
        }
        if (matchCount >= 3) {
          return {
            success: false,
            duplicateWarning: true,
            existingCount: matchCount,
            store: store,
            date: dateStr
          };
        }
      }
    }

    let savedCount = 0;
    for (const item of items) {
      let name = String(item.name || '').trim();
      const amount = parseFloat(item.price_twd);
      const category = CONFIG.CATEGORIES.includes(item.category) ? item.category : CONFIG.DEFAULT_CATEGORY;
      if (!name || isNaN(amount)) continue;

      // 加上店名前綴（若有店名且品名未包含店名）
      if (store && !name.startsWith(store)) {
        name = `${store} - ${name}`;
      }
      // 外幣：在品名後附上原幣金額
      if (isForex && item.price_original != null) {
        const currSymbol = currency === 'JPY' ? '¥' : currency;
        name = `${name} (${currSymbol}${Math.round(item.price_original)})`;
      }
      // 旅遊模式：分類欄加上 ✈️ 目的地前綴（F 欄），與 LINE Bot 格式一致
      const travelPrefix = (finalProject && finalProject !== CONFIG.DEFAULT_PROJECT)
        ? `${finalProject}-` : '';
      const finalCategory = `${travelPrefix}${category}`;

      writeToSheet({
        date: dateStr,
        item: name,
        amount: amount,
        userName: username,
        payment: finalPayment,
        category: finalCategory,
        project: finalProject,
        invoiceNum: '',
        source: 'Web逐項',
        cardId: card || ''
      });
      savedCount++;
    }
    return { success: true, count: savedCount };
  } catch (e) {
    return { success: false, error: '儲存失敗：' + e.message };
  }
}


// ── 重複記帳偵測（字串日期比對，避免時區時間戳問題）──────
function _parseDateToStr(rawDate) {
  // 統一回傳 'yyyy/MM/dd' 格式字串
  if (rawDate instanceof Date) {
    return Utilities.formatDate(rawDate, CONFIG.TIMEZONE, 'yyyy/MM/dd');
  }
  return String(rawDate).substring(0, 10).replace(/-/g, '/');
}
function _daysBetweenStr(s1, s2) {
  // s1, s2 格式 'yyyy/MM/dd'；回傳相差天數（Julian Day Number 公式）
  const [y1, m1, d1] = s1.split('/').map(Number);
  const [y2, m2, d2] = s2.split('/').map(Number);
  const jd = (y, m, d) => 367 * y - Math.floor(7 * (y + Math.floor((m + 9) / 12)) / 4) + Math.floor(275 * m / 9) + d;
  return Math.abs(jd(y1, m1, d1) - jd(y2, m2, d2));
}

function webCheckDuplicate(token, date, amount) {
  const username = CacheService.getScriptCache().get('web_' + token);
  if (!username) return { success: true, duplicates: [] };
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheets()[0];
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, duplicates: [] };
    // 讀取全部資料列，只取前 3 欄（日期、品名、金額）
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const amt = parseFloat(amount);
    const targetStr = String(date).replace(/-/g, '/'); // 'yyyy/MM/dd'
    const duplicates = [];

    // ─── 路徑①：單筆金額匹配（原有）±30 天內，差距 ≤ 1 ─────────────────
    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      if (!row[0]) continue;
      const rowDateStr = _parseDateToStr(row[0]);
      if (!rowDateStr || rowDateStr.length < 10) continue;
      const dayDiff = _daysBetweenStr(targetStr, rowDateStr);
      if (dayDiff > 30) continue; // ±30 天內才比對
      const rowAmt = parseFloat(row[2]);
      if (Math.abs(rowAmt - amt) <= 1) {
        const displayDate = row[0] instanceof Date
          ? Utilities.formatDate(row[0], CONFIG.TIMEZONE, 'MM/dd')
          : rowDateStr.substring(5).replace('/', '/');
        duplicates.push({ date: displayDate, item: String(row[1] || ''), amount: rowAmt, person: '' });
        if (duplicates.length >= 3) break;
      }
    }

    // ─── 路徑②：同日記帳分析 ─────────────────────────────────────────────
    // 路徑①未找到時，額外檢查同日加總是否疑似重複：
    //   A. 同日筆數 >= 5（大量明細，如超市逐項記帳）
    //   B. 同日加總在 OCR 金額的 ±5% 以內（且筆數 > 1）
    if (duplicates.length === 0) {
      let sameDaySum = 0;
      let sameDayCount = 0;
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row[0]) continue;
        const rowDateStr = _parseDateToStr(row[0]);
        if (!rowDateStr || rowDateStr.length < 10) continue;
        if (rowDateStr !== targetStr) continue;
        const rowAmt = parseFloat(row[2]);
        if (!isNaN(rowAmt)) { sameDaySum += rowAmt; sameDayCount++; }
      }
      const pctDiff = amt > 0 ? Math.abs(sameDaySum - amt) / amt : 1;
      const triggerA = sameDayCount >= 5;
      const triggerB = sameDayCount > 1 && pctDiff <= 0.05;
      if (triggerA || triggerB) {
        const displayDate = targetStr.substring(5).replace('/', '/');
        duplicates.push({
          date: displayDate,
          item: `共 ${sameDayCount} 筆明細，合計 $${Math.round(sameDaySum).toLocaleString('zh-TW')}，請確認是否已逐筆記帳`,
          amount: sameDaySum,
          person: ''
        });
      }
    }

    return { success: true, duplicates };
  } catch (e) {
    return { success: false, error: e.message, duplicates: [] };
  }
}


// ── 刪除記帳 ──────────────────────────────────
function webDeleteRecord(token, rowIndex) {
  const username = CacheService.getScriptCache().get('web_' + token);
  if (!username) return { success: false, error: '請重新登入' };
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheets()[0];
    const lastRow = sheet.getLastRow();
    if (rowIndex < 2 || rowIndex > lastRow) return { success: false, error: '記錄不存在' };
    sheet.deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    return { success: false, error: '刪除失敗：' + e.message };
  }
}

// ── 修改記帳 ──────────────────────────────────
function webEditRecord(token, rowIndex, data) {
  const username = CacheService.getScriptCache().get('web_' + token);
  if (!username) return { success: false, error: '請重新登入' };
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheets()[0];
    const lastRow = sheet.getLastRow();
    if (rowIndex < 2 || rowIndex > lastRow) return { success: false, error: '記錄不存在' };
    const row = sheet.getRange(rowIndex, 1, 1, 10).getValues()[0];
    if (data.date !== undefined && data.date) {
      const parts = data.date.split('-');
      const d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
      sheet.getRange(rowIndex, 1).setValue(d);
    }
    if (data.item !== undefined) sheet.getRange(rowIndex, 2).setValue(data.item.trim());
    if (data.amount !== undefined) sheet.getRange(rowIndex, 3).setValue(parseFloat(data.amount));
    if (data.person !== undefined && data.person) sheet.getRange(rowIndex, 4).setValue(data.person);
    if (data.payment !== undefined) sheet.getRange(rowIndex, 5).setValue(data.payment);
    if (data.category !== undefined) sheet.getRange(rowIndex, 6).setValue(
      CONFIG.CATEGORIES.includes(data.category) ? data.category : CONFIG.DEFAULT_CATEGORY);
    if (data.card !== undefined) sheet.getRange(rowIndex, 10).setValue(data.card);
    return { success: true };
  } catch (e) {
    return { success: false, error: '修改失敗：' + e.message };
  }
}

// ── 取得設定（分類、付款方式、卡別）────────────
function webGetConfig() {
  const cardsMap = CONFIG.CARDS_MAP;
  const cards = Object.keys(cardsMap);
  const usersMap = JSON.parse(AppProps.getProperty('USERS_MAP') || '{}');
  const persons = [...new Set(Object.values(usersMap))];
  return {
    categories: CONFIG.CATEGORIES,
    payments: ['現金', '信用卡', '悠遊卡', 'Line Pay', '轉帳', '其他'],
    cards,
    budget: CONFIG.MONTHLY_BUDGET,
    persons
  };
}
