// ============================================
// Expense - LINE 家庭記帳機器人 v7.8
// ============================================

function processExpense(userMessage, userName, preParsedData = null, skipReview = false) {
  try {
    const aiParsed = preParsedData || parseAndClassifyExpenseWithGemini(userMessage, userName);

    // [v8.0] 文字路徑：在變數賦值前正規化 aiParsed.payment，防止 Gemini 把卡片品牌名放入 payment
    if (!preParsedData && aiParsed && aiParsed.payment) {
      const _p = aiParsed.payment;
      const _pL = _p.toLowerCase();
      const GENERIC = ['現金', '信用卡', '悠遊卡', 'line pay', '街口支付', 'apple pay', 'google pay', '轉帳'];
      if (!GENERIC.some(g => _pL === g)) {
        if (!aiParsed.card) aiParsed.card = _p;
        aiParsed.payment = _pL.includes('悠遊卡') ? '悠遊卡' : '信用卡';
      } else if (_pL === '刷卡') {
        aiParsed.payment = '信用卡';
      }
    }

    let item, amount, payment, aiCategory, aiProject, expenseDate;

    if (aiParsed && aiParsed.amount) {
      item = aiParsed.item;
      amount = aiParsed.amount;
      payment = aiParsed.payment || CONFIG.DEFAULT_PAYMENT;
      expenseDate = aiParsed.date ? new Date(aiParsed.date + 'T12:00:00+08:00') : new Date();
      aiCategory = CONFIG.CATEGORIES.includes(aiParsed.category) ? aiParsed.category : CONFIG.DEFAULT_CATEGORY;
      const validProjects = CONFIG.PROJECTS.map(p => p.name);
      aiProject = validProjects.includes(aiParsed.project) ? aiParsed.project : CONFIG.DEFAULT_PROJECT;
    } else {
      const parsedData = parseExpenseInput(userMessage);
      if (!parsedData.isValid) {
        return {
          status: 'error',
          message: '⚠️ 聽不太懂這筆花費的金額！\n您可以試著直接說：「全聯買菜花了 500」，或使用傳統格式：「全聯 500」'
        };
      }
      item = parsedData.item;
      amount = parsedData.amount;
      payment = parsedData.payment;
      expenseDate = new Date();
      const fallbackClassify = classifyWithGemini(item, userName);
      aiCategory = fallbackClassify.category;
      aiProject = fallbackClassify.project;
    }

    const travelInfo = getActiveTravelProject();
    const isTravelMode = !!travelInfo;
    const travelProject = travelInfo ? travelInfo.projectName : null;
    const destination = travelProject ? travelProject.replace('✈️ ', '') : null;
    const category = travelProject ? `✈️${destination}-${aiCategory}` : aiCategory;
    const project = travelProject || aiProject;

    let finalItem = item;
    let finalAmount = amount;
    let originalAmountStr = null;

    if (isTravelMode && travelInfo.currency && travelInfo.rate) {
      const symbol = getCurrencySymbol(travelInfo.currency);
      finalItem = `${item}(${symbol}${amount.toLocaleString('zh-TW')})`;
      finalAmount = Math.round(amount * travelInfo.rate);
      const rateDate = AppProps.getProperty(`rate_date_${travelInfo.currency}`) || '';
      const rateDateStr = rateDate ? ` ${rateDate}匯率` : '';
      originalAmountStr = `${symbol}${amount.toLocaleString('zh-TW')} × ${travelInfo.rate}${rateDateStr}`;
    }

    // [v8.0] OCR 多頁去重：同一張帳單分多頁傳送時，60 秒內相同金額的 OCR 視為同一張收據，略過重複頁
    if (preParsedData) {
      const dedupKey = `ocr_dedup_${finalAmount}`;
      const ocrCache = CacheService.getScriptCache();
      if (ocrCache.get(dedupKey)) {
        return { status: 'success', message: `📄 偵測到重複頁面（NT$ ${finalAmount.toLocaleString()}），已略過此張照片，帳本資料不變。` };
      }
      ocrCache.put(dedupKey, '1', 60); // 60 秒內相同金額的 OCR 視為同一張收據
    }

    // [v7.7] OCR 路徑：從 card_last4 解析對應的 cardId（必須在智能合併前宣告）
    let cardId = null;
    if (preParsedData && aiParsed.card_last4) {
      const cardsMap = CONFIG.CARDS_MAP;
      if (aiParsed.card_last4 === 'easycard') {
        // 悠遊卡感應：找 type=easycard 且 owner 符合的卡
        const match = Object.entries(cardsMap).find(([, info]) =>
          info.type === 'easycard' && info.owner === userName);
        if (match) cardId = match[0];
      } else {
        // 信用卡：比對末四碼（不限定 owner，因為有時幫對方刷）
        const match = Object.entries(cardsMap).find(([, info]) =>
          info.last4 === aiParsed.card_last4);
        if (match) cardId = match[0];
      }
    }

    // [v8.0] 文字路徑：支付方式正規化 + 卡別偵測
    if (!preParsedData && aiParsed) {
      const cardsMap = CONFIG.CARDS_MAP;

      // rawCard：優先用 Gemini 的 card 欄位，否則從 payment 提取
      const rawCard = (aiParsed.card || payment || '');
      const rawLower = rawCard.toLowerCase();

      // 去除常見通用後綴（如「刷卡」「信用卡」）取得純品牌關鍵字
      const cardKeyword = rawLower.replace(/刷卡$|信用卡$/, '').trim();

      if (rawLower.includes('悠遊卡') || rawLower.includes('easycard')) {
        // 悠遊卡路徑
        const match = Object.entries(cardsMap).find(([, info]) =>
          info.type === 'easycard' && info.owner === userName);
        if (match) { cardId = match[0]; payment = '悠遊卡'; }
      } else if (cardKeyword) {
        // 信用卡：用純品牌關鍵字比對 bank 名稱
        const matches = Object.entries(cardsMap).filter(([id, info]) =>
          cardKeyword.includes((info.bank || '').toLowerCase()) ||
          cardKeyword.includes(id.toLowerCase())
        );
        if (matches.length >= 1) {
          const ownerMatch = matches.find(([, info]) => info.owner === userName);
          cardId = (ownerMatch || matches[0])[0];
          payment = '信用卡';
        } else if (rawLower.includes('刷卡') || rawLower.includes('信用卡')) {
          // 未匹配到卡別，但確定是信用卡消費 → 至少正規化 payment
          payment = '信用卡';
        }
      }
      if (payment === '刷卡') payment = '信用卡';
    }

    // === [v5.5] 智能合併邏輯 (OCR 覆寫 CSV 舊帳) ===
    let isMerged = false;
    let mergedRowCount = 0;  // [v7.7] 記錄合併了幾筆，用於回報訊息
    let finalSource = preParsedData ? 'OCR圖片掃描' : 'LINE';
    const sheet = getSheet();
    const dateOnlyStr = Utilities.formatDate(expenseDate, CONFIG.TIMEZONE, 'yyyy/MM/dd');

    if (preParsedData) {
      const data = sheet.getDataRange().getValues();
      // [v5.5 fix] 收集所有符合條件的列，唯一匹配才執行合併，避免同日同額不同消費被誤合
      const matchCandidates = [];
      for (let i = data.length - 1; i > 0; i--) {
        const rowDateObj = new Date(data[i][0]);
        if (isNaN(rowDateObj)) continue;
        const rowDate = Utilities.formatDate(rowDateObj, CONFIG.TIMEZONE, 'yyyy/MM/dd');
        const rowAmount = parseInt(data[i][2]);
        // 條件：日期相同、金額相同
        if (rowDate === dateOnlyStr && rowAmount === finalAmount) {
          matchCandidates.push(i + 1);
        }
      }
      if (matchCandidates.length === 1) {
        const targetRow = matchCandidates[0];
        sheet.getRange(targetRow, 2).setValue(finalItem);          // 覆寫品名
        sheet.getRange(targetRow, 4).setValue(userName);           // [v5.5 fix] 更新記帳人為實際 OCR 掃描者
        sheet.getRange(targetRow, 5).setValue(payment);            // [v5.5 fix] 更新支付方式
        sheet.getRange(targetRow, 6).setValue(category);           // 覆寫分類
        sheet.getRange(targetRow, 7).setValue(project);            // 覆寫專案
        sheet.getRange(targetRow, 9).setValue('CSV+OCR智能合併');   // 更新來源狀態
        // [v7.7] 若 OCR 偵測到卡別也一併回填
        if (cardId) sheet.getRange(targetRow, 10).setValue(cardId);
        isMerged = true;
        mergedRowCount = 1;
      }

      // [v7.7] 發票號碼合併路徑（優先）：OCR 成功提取發票號碼，直接找對應列整批標記
      if (!isMerged && aiParsed.invoice_num) {
        const norm = s => String(s).replace(/[\s\-]/g, '').toUpperCase();
        const invoiceRows = [];
        for (let i = 1; i < data.length; i++) {
          if (norm(data[i][7]) === norm(aiParsed.invoice_num)) invoiceRows.push(i + 1);
        }
        if (invoiceRows.length > 0) {
          invoiceRows.forEach(row => {
            sheet.getRange(row, 9).setValue('CSV+OCR智能合併');
            if (cardId) sheet.getRange(row, 10).setValue(cardId);
          });
          isMerged = true;
          mergedRowCount = invoiceRows.length;
        }
      }

      // [v7.7] 發票分組加總合併路徑（備援）：不需 OCR 讀出發票號碼
      // 將試算表同日期有發票號碼的列依發票分組，找加總等於 OCR 金額的唯一那組
      if (!isMerged) {
        const invoiceGroups = {};
        for (let i = 1; i < data.length; i++) {
          const rowDateObj = new Date(data[i][0]);
          if (isNaN(rowDateObj)) continue;
          const rowDate = Utilities.formatDate(rowDateObj, CONFIG.TIMEZONE, 'yyyy/MM/dd');
          if (rowDate !== dateOnlyStr) continue;
          const inv = String(data[i][7] || '').trim();
          if (!inv) continue;  // 沒有發票號碼的列跳過
          if (!invoiceGroups[inv]) invoiceGroups[inv] = { rows: [], total: 0 };
          invoiceGroups[inv].rows.push(i + 1);
          invoiceGroups[inv].total += (parseInt(data[i][2]) || 0);
        }
        const matchGroups = Object.values(invoiceGroups).filter(
          g => g.total === finalAmount && g.rows.length > 1
        );
        if (matchGroups.length === 1) {  // 唯一匹配才執行，避免誤合
          matchGroups[0].rows.forEach(row => {
            sheet.getRange(row, 9).setValue('CSV+OCR智能合併');
            if (cardId) sheet.getRange(row, 10).setValue(cardId);
          });
          isMerged = true;
          mergedRowCount = matchGroups[0].rows.length;
        }
      }
    }

    // === [v6.0] 智能大額消費確認機制 (CacheService) ===
    const recordObj = {
      date: expenseDate,
      item: finalItem, amount: finalAmount, userName, payment, category, project, invoiceNum: '',
      source: isMerged ? 'CSV+OCR智能合併' : finalSource,
      cardId: cardId || '',  // [v7.7] 卡別
      // [v7.3+] 分期資訊隨 record 一起存入快取，供大額確認路徑使用
      installments: (aiParsed && aiParsed.installments > 1) ? aiParsed.installments : null,
      installment_day: (aiParsed && aiParsed.installments > 1) ? (aiParsed.installment_day || expenseDate.getDate()) : null
    };

    let dupWarning = null;
    if (!isMerged) {
      if (finalAmount >= 5000 && !preParsedData && !skipReview) { // 大筆且非掃描匯入，需二次確認；批次輸入時 skipReview=true
        const uuid = Utilities.getUuid();
        CacheService.getScriptCache().put(uuid, JSON.stringify(recordObj), 600); // 存入快取10分鐘

        return {
          status: 'review_needed',
          message: '⚠️ 收到一筆大額消費，請點擊確認是否寫入帳本。',
          flexMessage: buildReviewFlexMessage(recordObj, uuid)
        };
      }

      dupWarning = checkDuplicate(recordObj); // [v7.9] 重複偵測（寫入前掃描既有資料）
      const writeSuccess = writeToSheet(recordObj);
      if (!writeSuccess) {
        return { status: 'error', message: '❌ 記帳失敗！請稍後再試或聯絡管理員' };
      }
    }

    let mergeText = '';
    if (isMerged) {
      const countStr = mergedRowCount > 1 ? `共 ${mergedRowCount} 筆` : '';
      const cardStr = cardId ? `，標記【${cardId}】` : '';
      mergeText = `\n🔄 [已自動合併載具舊帳${countStr ? '，' + countStr : ''}${cardStr}]`;
    }
    let textMessage = buildSuccessMessage({
      userName, item: finalItem, category, project, amount: finalAmount, payment, isTravelMode, originalAmountStr, expenseDate
    }) + mergeText;

    // [v7.3+] 分期偵測：若 AI 解析到分期資訊，自動寫入固定收支設定表
    if (!isMerged && aiParsed && aiParsed.installments && aiParsed.installments > 1) {
      const installResult = addInstallmentToFixedSchedule({
        item: finalItem, totalAmount: finalAmount, installments: aiParsed.installments,
        day: aiParsed.installment_day || expenseDate.getDate(),
        category, project, payment, expenseDate
      });
      if (installResult) {
        textMessage += `\n\n💳 【分期自動入帳】\n已自動新增至固定收支 ${aiParsed.installments} 期\n📅 每月 ${installResult.day} 號扣 ${formatCurrency(installResult.monthlyAmount)}\n⏳ 扣款至 ${installResult.end_month} 止`;
      }
    }

    // [v6.0] 替換純文字為華麗的 Flex Receipt；dupWarning 作為獨立 warningText 訊息發送
    return { status: 'success', message: textMessage, flexMessage: buildReceiptFlexMessage(recordObj), warningText: dupWarning || null };
  } catch (error) {
    logError('processExpense', error);
    return { status: 'error', message: '❌ 記帳時發生錯誤\n' + error.toString() };
  }
}

function parseExpenseInput(userMessage) {
  const cleanMessage = userMessage.replace(/　/g, ' ').trim();
  const parts = cleanMessage.split(/\s+/);
  if (parts.length < 2) {
    return { isValid: false, errorMessage: '⚠️ 格式錯誤！\n正確格式：項目 金額 [支付方式]\n例如：午餐 150 信用卡' };
  }
  let amountIndex = -1;
  for (let i = parts.length - 1; i >= 1; i--) {
    if (isValidAmount(parts[i])) { amountIndex = i; break; }
  }
  if (amountIndex === -1) {
    return { isValid: false, errorMessage: '⚠️ 金額格式錯誤！\n請輸入有效的數字\n例如：150 或 1500' };
  }
  const item = parts.slice(0, amountIndex).join(' ');
  const rawAmount = parts[amountIndex];
  const payment = amountIndex < parts.length - 1 ? parts.slice(amountIndex + 1).join(' ') : CONFIG.DEFAULT_PAYMENT;
  return { isValid: true, item, amount: parseAmount(rawAmount), payment };
}

function buildSuccessMessage(data) {
  const { userName, item, category, project, amount, payment, isTravelMode, originalAmountStr, expenseDate } = data;
  let message = buildExpenseMessage(userName, item, category, project, amount, payment, originalAmountStr, expenseDate);
  if (isTravelMode) {
    const projectTotal = calculateProjectTotal(project);
    message += `\n✈️ 旅遊累計：${formatCurrency(projectTotal)}`;
    const props = AppProps;
    const budget = parseInt(props.getProperty('travel_budget') || '0');
    if (budget > 0) {
      const remaining = budget - projectTotal;
      if (remaining < 0) {
        message += `\n🚨 旅遊預算已超支 ${formatCurrency(Math.abs(remaining))}！`;
      } else {
        const pct = (projectTotal / budget * 100).toFixed(1);
        message += `\n📊 預算使用 ${pct}%，剩餘 ${formatCurrency(remaining)}`;
      }
    }
  } else if (project !== CONFIG.DEFAULT_PROJECT) {
    const projectTotal = calculateProjectTotal(project);
    message += `\n🔥 [${project}] 已累計：${formatCurrency(projectTotal)}`;

    // [v7.0 新增] 檢查專案專屬子預算防爆鎖（以本月累計比對月上限）
    if (CONFIG.PROJECT_BUDGETS && CONFIG.PROJECT_BUDGETS[project]) {
      const limit = CONFIG.PROJECT_BUDGETS[project];
      const d = expenseDate || new Date();
      const monthlyTotal = calculateProjectTotal(project, d.getMonth() + 1, d.getFullYear());
      const remaining = limit - monthlyTotal;
      const pct = (monthlyTotal / limit * 100).toFixed(1);
      if (remaining < 0) {
        message += `\n🚨 哎呀！本月這個專案已超支 ${formatCurrency(Math.abs(remaining))} 啦！(月上限 ${formatCurrency(limit)})`;
      } else if (pct > 80) {
        message += `\n⚠️ 【子預算提醒】本月已用 ${pct}%，額度只剩 ${formatCurrency(remaining)}！`;
      } else {
        message += `\n📊 【子預算】本月使用率：${pct}%，餘額 ${formatCurrency(remaining)}`;
      }
    }
  }
  const budgetAlert = checkBudgetAlert();
  if (budgetAlert) message += budgetAlert; // checkBudgetAlert 回傳值已含前置 \n，不重複添加
  // [v7.8] 分類預算超標警示
  const catBudgetAlert = checkCategoryBudgetAlert(category);
  if (catBudgetAlert) message += catBudgetAlert;
  return message;
}

// [v8.0] 刪除確認第一步：顯示 Flex 確認卡片，不直接刪除
function handleDeleteRequest(userName) {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: 'error', message: '⚠️ 目前沒有資料可以刪除！' };

    const lastRowData = sheet.getRange(lastRow, 1, 1, 9).getValues()[0];
    const recordUser = String(lastRowData[3] || '');
    if (!canDeleteRecord(recordUser, userName)) {
      return { status: 'error', message: `⚠️ 這筆是 ${recordUser} 記的，只有本人或 Ted 可以刪除！` };
    }

    const uuid = Utilities.getUuid();
    const record = {
      rowNum: lastRow,
      item: String(lastRowData[1] || ''),
      amount: parseInt(lastRowData[2]) || 0,
      date: lastRowData[0],
      category: String(lastRowData[5] || ''),
      payment: String(lastRowData[4] || ''),
      userName: recordUser
    };
    CacheService.getScriptCache().put(`del_${uuid}`, JSON.stringify(record), 120); // 2 分鐘有效

    return {
      status: 'delete_pending',
      message: `確認刪除「${record.item}」$${record.amount}？`,
      flexMessage: buildDeleteConfirmFlexMessage(record, uuid)
    };
  } catch (error) {
    logError('handleDeleteRequest', error);
    return { status: 'error', message: '❌ 操作失敗：' + error.toString() };
  }
}

// [v8.0] 刪除確認第二步：從 cache 取回記錄後執行真正的刪除
function executeDelete(uuid, userName) {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(`del_${uuid}`);
    if (!cached) return { status: 'error', message: '⚠️ 確認請求已逾時（2 分鐘），請重新輸入「刪除」' };

    const record = JSON.parse(cached);
    cache.remove(`del_${uuid}`);

    const sheet = getSheet();
    if (sheet.getLastRow() !== record.rowNum) {
      return { status: 'error', message: '⚠️ 帳本在確認期間有異動，請重新輸入「刪除」' };
    }
    if (!canDeleteRecord(record.userName, userName)) {
      return { status: 'error', message: `⚠️ 這筆是 ${record.userName} 記的，沒有權限刪除！` };
    }

    sheet.deleteRow(record.rowNum);

    // 同步清除對應分期固定收支
    let extraMsg = '';
    try {
      const fixedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('固定收支設定');
      if (fixedSheet) {
        const fixedData = fixedSheet.getDataRange().getValues();
        const pattern = new RegExp(`^${record.item.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\(分\\d+期\\)$`);
        for (let i = fixedData.length - 1; i >= 1; i--) {
          const name = String(fixedData[i][2] || '');
          if (pattern.test(name)) {
            fixedSheet.deleteRow(i + 1);
            extraMsg = `\n♻️ 同步清除固定收支：【${name}】`;
            break;
          }
        }
      }
    } catch (e) { logError('executeDelete - fixedSheet', e); }

    const lineSep = '━━━━━━━━━━';
    return {
      status: 'success',
      message: `🗑️ 已刪除上一筆記帳\n${lineSep}\n📝 項目：${record.item}\n💰 金額：${formatCurrency(record.amount)}\n👤 記錄人：${record.userName}${extraMsg}`
    };
  } catch (error) {
    logError('executeDelete', error);
    return { status: 'error', message: '❌ 刪除失敗：' + error.toString() };
  }
}

function deleteLastRow(userName) {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return '⚠️ 目前沒有資料可以刪除！';
    const lastRowData = sheet.getRange(lastRow, 1, 1, 9).getValues()[0];
    const recordUser = lastRowData[3];
    if (!canDeleteRecord(recordUser, userName)) {
      return `⚠️ 這筆是 ${recordUser} 記的，不能刪除喔！\n只有本人或Ted可以刪除`;
    }
    sheet.deleteRow(lastRow);

    const deletedItem = String(lastRowData[1] || '');
    const lineSep = '━━━━━━━━━━';
    let msg = `🗑️ 已刪除上一筆記帳\n${lineSep}\n📝 項目：${deletedItem}\n💰 金額：${formatCurrency(lastRowData[2])}\n👤 記錄人：${recordUser}`;

    // [v7.3+] 同步清除對應的分期固定收支（格式：品名(分N期)）
    try {
      const fixedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('固定收支設定');
      if (fixedSheet) {
        const fixedData = fixedSheet.getDataRange().getValues();
        // 尋找固定收支表中「品名以 deletedItem 開頭且帶有(分N期)」的列
        const installmentPattern = new RegExp(`^${deletedItem.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\(分\\d+期\\)$`);
        let matchRowIndex = -1;
        let matchItem = '';
        for (let i = fixedData.length - 1; i >= 1; i--) {
          const fixedItemName = String(fixedData[i][2] || '');
          if (installmentPattern.test(fixedItemName)) {
            matchRowIndex = i + 1; // Sheet 列號（1-indexed）
            matchItem = fixedItemName;
            break;
          }
        }
        if (matchRowIndex !== -1) {
          fixedSheet.deleteRow(matchRowIndex);
          msg += `\n\n♻️ 同步清除固定收支：\n📅 已刪除【${matchItem}】的分期設定`;
        }
      }
    } catch (fixedErr) {
      logError('deleteLastRow - fixedSheet cleanup', fixedErr);
      msg += '\n\n⚠️ 主帳本已刪除，但清除固定收支時發生錯誤，請手動執行「刪除固定」。';
    }

    return msg;
  } catch (error) { logError('deleteLastRow', error); return '❌ 刪除失敗：' + error.toString(); }
}

function canDeleteRecord(recordUser, currentUser) {
  return recordUser === currentUser || currentUser === 'Ted (老公)';
}

function writeToSheet(record) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = getSheet();

    // 🌟 關鍵修復：確保從 JSON 還原的 date 字串被強制轉回 Date 物件
    const recordDate = record.date ? new Date(record.date) : new Date();
    const dateOnlyStr = Utilities.formatDate(recordDate, CONFIG.TIMEZONE, 'yyyy/MM/dd');

    sheet.appendRow([dateOnlyStr, record.item, record.amount, record.userName, record.payment, record.category, record.project, record.invoiceNum || '', record.source || 'LINE', record.cardId || '']); // col 10: [v7.7] 卡別
    return true;
  } catch (error) {
    logError('writeToSheet', error);
    return false;
  } finally {
    lock.releaseLock();
  }
}

// [v7.9] 重複記帳偵測：掃描最近 50 筆，尋找同日同金額或同日同品名首詞
function checkDuplicate(recordObj) {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return null;
    const startRow = Math.max(2, lastRow - 50);
    const numRows = lastRow - startRow + 1;
    const data = sheet.getRange(startRow, 1, numRows, 3).getValues();
    const targetDate = Utilities.formatDate(new Date(recordObj.date), CONFIG.TIMEZONE, 'yyyy/MM/dd');
    const targetAmt = recordObj.amount;
    const targetKeyword = String(recordObj.item || '').split(/[\s\-（(]/)[0].trim();
    Logger.log(`[checkDuplicate] targetDate=${targetDate} targetAmt=${targetAmt} targetKeyword=${targetKeyword} rows=${numRows}`);

    for (let i = data.length - 1; i >= 0; i--) {
      const rawDate = data[i][0];
      if (!rawDate) continue;
      const rowDate = rawDate instanceof Date
        ? Utilities.formatDate(rawDate, CONFIG.TIMEZONE, 'yyyy/MM/dd')
        : String(rawDate).substring(0, 10).replace(/-/g, '/');
      if (rowDate !== targetDate) continue;
      const rowAmt = parseInt(data[i][2]) || 0;
      const rowItem = String(data[i][1] || '');
      const sameAmount = rowAmt === targetAmt && targetAmt > 0;
      const sameKeyword = targetKeyword.length >= 2 && rowItem.includes(targetKeyword);
      Logger.log(`[checkDuplicate] match candidate: rowDate=${rowDate} rowAmt=${rowAmt} rowItem=${rowItem} sameAmt=${sameAmount} sameKw=${sameKeyword}`);
      if (sameAmount || sameKeyword) {
        return `⚠️ 注意：今日已有相似記錄「${rowItem} $${rowAmt}」，如為重複請說「刪除」`;
      }
    }
    return null;
  } catch (e) {
    Logger.log(`[checkDuplicate] ERROR: ${e.toString()}`);
    return null;
  }
}

function handleEditLastEntry(userMessage, userName) {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: 'error', message: '⚠️ 目前沒有資料可以修改！' };

    const lastRowData = sheet.getRange(lastRow, 1, 1, 10).getValues()[0];
    const recordUser = String(lastRowData[3] || '');

    if (!canDeleteRecord(recordUser, userName)) {
      return { status: 'error', message: `⚠️ 這筆是 ${recordUser} 記的，只有本人或 Ted 可以修改！` };
    }

    const parsed = parseEditLastEntryWithGemini(userMessage);
    if (!parsed || !parsed.field || parsed.value === undefined || parsed.value === null) {
      return { status: 'error', message: '⚠️ 無法解析修改內容，請說明要改哪個欄位與新的值。\n範例：修改上一筆 金額 300' };
    }

    const { field, value } = parsed;
    const colMap = { item: 2, amount: 3, payment: 5, category: 6, card: 10 };
    const col = colMap[field];
    if (!col) return { status: 'error', message: `⚠️ 不支援修改「${field}」欄位！` };

    let finalValue = value;
    let displayOld, displayNew;

    if (field === 'amount') {
      finalValue = parseInt(value);
      if (isNaN(finalValue)) return { status: 'error', message: '⚠️ 金額格式錯誤，請輸入數字！' };
      displayOld = formatCurrency(lastRowData[2]);
      displayNew = formatCurrency(finalValue);
    } else if (field === 'category') {
      if (!CONFIG.CATEGORIES.includes(String(value))) {
        return { status: 'error', message: `⚠️ 分類「${value}」不在清單中！\n可用分類：${CONFIG.CATEGORIES.join('、')}` };
      }
      displayOld = lastRowData[5];
      displayNew = value;
    } else if (field === 'item') {
      displayOld = lastRowData[1];
      displayNew = value;
    } else if (field === 'card') {
      const cardsMap = CONFIG.CARDS_MAP;
      if (Object.keys(cardsMap).length > 0 && !cardsMap[String(value)]) {
        const validCards = Object.keys(cardsMap).join('、');
        return { status: 'error', message: `⚠️ 卡別「${value}」不在清單中！\n可用卡別：${validCards}` };
      }
      displayOld = String(lastRowData[9] || '（未設定）');
      displayNew = value;
    } else {
      displayOld = lastRowData[4];
      displayNew = value;
    }

    sheet.getRange(lastRow, col).setValue(finalValue);

    const fieldNames = { item: '品名', amount: '金額', payment: '支付方式', category: '分類', card: '卡別' };
    const lineSep = '━━━━━━━━━━';
    return {
      status: 'success',
      message: `✅ 修改成功！\n${lineSep}\n📝 項目：${lastRowData[1]}\n🔧 ${fieldNames[field]}：${displayOld} ➔ ${displayNew}`
    };
  } catch (error) {
    logError('handleEditLastEntry', error);
    return { status: 'error', message: '❌ 修改失敗：' + error.toString() };
  }
}
