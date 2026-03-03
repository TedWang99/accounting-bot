// ============================================
// Main - LINE 家庭記帳機器人 v7.8
// ============================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🤖 記帳小幫手')
    .addItem('📥 立即掃描雲端載具資料夾', 'processInvoices')
    .addSeparator()
    .addItem('⚙️ 1. 建立「固定收支設定」工作表', 'initFixedScheduleSheet')
    .addItem('⏰ 2. 設定固定收支【每日】自動入帳', 'setupDailyTrigger')
    .addItem('⏰ 3. 設定載具發票【週末】自動掃描', 'setupWeekendInvoiceTrigger')
    .addSeparator()
    .addItem('⏰ 設定每月自動月報推播', 'setupMonthlyTrigger')
    .addItem('🧪 立即測試月報推播', 'sendMonthlyReport')
    .addItem('📊 設定每週自動週報推播（週一 8:00）', 'setupWeeklyTrigger')
    .addItem('🧪 立即測試週報推播', 'sendWeeklyReport')
    .addItem('❌ 取消全部自動推播', 'removeAllTriggers')
    .addSeparator()
    .addItem('🔄 重新分類【其他】項目（AI 重新判斷）', 'reclassifyAllFromMenu')
    .addSeparator()
    .addItem('💳 設定卡別下拉選單（J 欄）', 'setupCardDropdown')
    .addItem('⏰ 設定卡別下拉【每日 05:00 自動更新】', 'setupCardDropdownTrigger')
    .addSeparator()
    .addItem('🎊 設定年度回顧推播（每年 12/31）', 'setupAnnualTrigger')
    .addItem('🧪 立即測試年度回顧推播', 'testSendAnnualReport')
    .addToUi();
}

function doPost(e) {
  try {
    // --- 🛡️ 安全性驗證：確認請求是否來自 LINE 官方 ---
    const channelSecret = CONFIG.LINE_CHANNEL_SECRET;
    if (channelSecret) {
      const requestBody = e.postData.contents; // 原始請求內容
      const hash = Utilities.computeHmacSha256Signature(requestBody, channelSecret, Utilities.Charset.UTF_8);
      const expectedSignature = Utilities.base64Encode(hash);

      // LINE 會在 header 放 x-line-signature（小寫或大寫都有可能）
      // 以安全方式存取 e.headers，避免 GAS 某些環境下 headers 不存在時拋出 TypeError
      const headers = e.headers || {};
      const clientSignature = headers['x-line-signature'] || headers['X-Line-Signature'];

      if (clientSignature !== expectedSignature) {
        logError('doPost', '⚠️ Webhook 簽章驗證失敗！可能有非 LINE 來源的請求。');
        return createResponse('error', 'Invalid signature');
      }
    }

    const events = parseLineEvents(e);
    if (!events.length) return createResponse('no_event');

    // 多張照片同時送出：全部交給 Gemini 一次判斷成一筆帳
    const imageEvents = events.filter(ev => ev.type === 'message' && ev.message?.type === 'image');
    if (imageEvents.length > 1) {
      const lastEv = imageEvents[imageEvents.length - 1];
      const userId = lastEv.source.userId;
      const userName = getUserName(userId);
      const messageIds = imageEvents.map(ev => ev.message.id);
      const response = handleMultiPageImages(messageIds, userName);
      replyLine(lastEv.replyToken, response.message, response.imageUrl, response.flexMessage);
      return createResponse(response.status);
    }

    const event = events[0];
    if (!event) return createResponse('no_event');

    const { replyToken, userMessage, userId } = extractEventData(event);
    const userName = getUserName(userId);

    if (event.type === 'message' && event.message.type === 'image') {
      const messageId = event.message.id;
      const response = handleImageMessage(messageId, userName);
      replyLine(replyToken, response.message, response.imageUrl, response.flexMessage);
      return createResponse(response.status);
    }

    // 攔截按鈕回傳事件 (Flex Message Postback)
    if (event.type === 'postback') {
      const response = handlePostbackEvent(event, userName);
      replyLine(replyToken, response.message, response.imageUrl, response.flexMessage);
      return createResponse(response.status);
    }

    if (!isTextMessage(event)) return createResponse('non_text');

    const response = handleCommand(userMessage, userName);
    replyLine(replyToken, response.message, response.imageUrl, response.flexMessage, response.warningText);
    return createResponse(response.status);
  } catch (error) {
    logError('doPost', error);
    return createResponse('error', error.toString());
  }
}

function handleMultipleExpenses(lines, userName) {
  const results = [];
  let successCount = 0;
  let failCount = 0;
  for (const line of lines) {
    const msg = line.toLowerCase().trim();
    let result;
    if (/^(更新|修改)/.test(msg)) {
      result = handleUpdateFixedSchedule(line);
    } else if (/^(刪除固定|移除固定|取消固定)/.test(msg)) {
      result = handleDeleteFixedSchedule(line);
    } else if (/^(新增固定|加固定|每月固定|每個?月\d{1,2}號)/.test(msg)) {
      result = handleAddFixedSchedule(line);
    } else {
      result = processExpense(line, userName, null, true); // 批次輸入跳過大額確認，直接寫入
    }
    if (result.status === 'success') {
      successCount++;
      const itemMatch = result.message.match(/(?:📝|📅) 項目：(.+)/);
      const amountMatch = result.message.match(/💰 金額：(.+)/);

      const item = itemMatch ? itemMatch[1] : line;
      const amount = amountMatch ? amountMatch[1] : '處理成功';

      results.push(`✅ ${item} → ${amount}`);
    } else {
      failCount++;
      const errorMsg = result.message.split('\n')[0].replace('⚠️ ', '').replace('❌ ', '');
      results.push(`❌ 「${line}」→ ${errorMsg}`);
    }
  }
  const lineSep = '━━━━━━━━━━';
  let summary = `📋 批次處理完成！共 ${lines.length} 筆\n`;
  summary += `✅ 成功 ${successCount} 筆　❌ 失敗 ${failCount} 筆\n`;
  summary += `${lineSep}\n`;
  summary += results.join('\n');
  const travelInfo = getActiveTravelProject();
  if (travelInfo) {
    const projectTotal = calculateProjectTotal(travelInfo.projectName);
    summary += `\n${lineSep}\n✈️ 旅遊累計：${formatCurrency(projectTotal)}`;
    const budget = parseInt(AppProps.getProperty('travel_budget') || '0');
    if (budget > 0) {
      const remaining = budget - projectTotal;
      if (remaining < 0) {
        summary += `\n🚨 已超支 ${formatCurrency(Math.abs(remaining))}！`;
      } else {
        summary += `\n📊 預算使用 ${(projectTotal / budget * 100).toFixed(1)}%，剩餘 ${formatCurrency(remaining)}`;
      }
    }
  }
  return { status: 'multi_expense', message: summary };
}

function handleCommand(userMessage, userName) {
  const lines = userMessage.split(/\n/).map(l => l.trim()).filter(l => l.length > 0);
  if (lines.length > 1) {
    return handleMultipleExpenses(lines, userName);
  }
  // 批次輸入不支援「刪除固定」，這裡僅處理單行的固定收支指令
  const msg = userMessage.toLowerCase().trim();
  // [v7.7] 補記卡別：「這筆用XX卡」「上一筆用XX卡」
  if (/^(這筆|上一筆|剛才那筆)用.+卡/.test(msg)) return handleUpdateLastEntryCard(userMessage, userName);
  // [v7.7] 多品項整批標記卡別確認：「確認標記」(由 handleUpdateLastEntryCard 發出的確認流程)
  if (/^確認標記卡別/.test(msg)) return confirmBatchCardTag(userMessage, userName);
  if (/^(修改上一筆|改上一筆|更正上一筆)/.test(msg)) return handleEditLastEntry(userMessage, userName);
  if (/^(更新|修改)/.test(msg)) return handleUpdateFixedSchedule(userMessage);
  if (/^(刪除固定|移除固定|取消固定)/.test(msg)) return handleDeleteFixedSchedule(userMessage);
  if (/^(新增固定|加固定|每月固定|每個?月\d{1,2}號)/.test(msg)) return handleAddFixedSchedule(userMessage);
  if (/^(查固定|查詢固定)/.test(msg)) return handleQueryFixedSchedule(userMessage);
  if (/^(查預算|預算狀態|預算表)/.test(msg)) return handleQueryBudgets(userMessage);
  if (/^(查分類預算|分類預算)/.test(msg)) return handleQueryCategoryBudgets(); // [v7.8]

  // [v6.0 新增] 支援「指定月份」或「上個月」的報表查詢
  const monthMatch = msg.match(/^(?:([1-9]|1[0-2])|上個?)\s*(?:月|月份)\s*的?\s*(報表|月報|查詢|統計|總計|分類)$/);
  if (monthMatch) {
    const rawMonth = monthMatch[1]; // 若 match 到「上個」則為 undefined
    const action = monthMatch[2];
    let { month, year } = getCurrentMonthYear();

    if (rawMonth) {
      month = parseInt(rawMonth, 10);
      // 若查詢指定月大於當前的現實月份(例如1月時查12月)，推斷為去年
      if (month > (new Date().getMonth() + 1)) year -= 1;
    } else {
      month -= 1;
      if (month === 0) {
        month = 12;
        year -= 1;
      }
    }

    if (/(報表|月報)/.test(action)) return getMonthlyReport(month, year);
    if (/(查詢|總計)/.test(action)) return getMonthlyTotal(month, year); // getMonthlyTotal 回傳 { status, message, flexMessage }，不再二次包裹
    if (/(統計|分類)/.test(action)) return { status: 'stats', message: getCategoryStats(month, year) };
  }

  const commandHandlers = {
    help: () => ({ status: 'help', message: getHelpMessage() }),
    last: () => ({ status: 'last', message: getLastEntry() }),
    monthly: () => getMonthlyTotal(), // getMonthlyTotal 回傳 { status, message, flexMessage }，不再二次包裹
    delete: () => handleDeleteRequest(userName), // [v8.0] 先顯示確認 Flex，不直接刪
    report: () => getMonthlyReport(), // [v5.6] getMonthlyReport 會直接回傳 { status, message, imageUrl }
    stats: () => ({ status: 'stats', message: getCategoryStats() }),
    query: () => handleProjectQuery(userMessage)
  };

  for (const [key, commands] of Object.entries(CONFIG.COMMANDS)) {
    if (commands.includes(msg)) {
      return commandHandlers[key.toLowerCase()]();
    }
  }
  // [v7.5] 日期範圍查詢：偵測到日期/週別關鍵字才走新路由，否則維持原有關鍵字查詢
  if ((msg.startsWith('查 ') || msg.startsWith('查詢 ')) && parseDateRangeInput(msg.replace(/^查詢?\s+/, '')) !== null) {
    return handleDateRangeQuery(userMessage);
  }
  // [v7.7] 按卡查詢：偵測到 CARDS_MAP 中的卡別/銀行關鍵字時走卡別路由
  if (msg.startsWith('查 ') || msg.startsWith('查詢 ')) {
    const cardParsed = parseCardQuery(msg.replace(/^查詢?\s+/, ''));
    if (cardParsed) return handleCardQuery(userMessage, userName);
  }
  if (msg.startsWith('查 ') || msg.startsWith('查詢 ') || msg.startsWith('query ')) {
    return commandHandlers.query();
  }
  if (/(旅遊狀態|travel status|目前花費|旅遊進度|花多少了|花了多少)/.test(msg)) {
    const props = AppProps;
    const activeProject = props.getProperty('travel_project');
    if (activeProject) {
      return { status: 'travel_status', message: getTravelStatus() };
    }

    let keyword = msg.replace(/(這趟|我們|目前|花多少了|花了多少|總共|查詢|查|的|費用|旅遊|狀態|\?|？)/g, '').trim();
    if (keyword) {
      return { status: 'query_project', message: queryProjectTotal(keyword) };
    } else {
      return { status: 'travel_status', message: getTravelStatus() };
    }
  }
  if (/(結束旅遊|回國|旅遊結束|結束旅行|回台灣|end travel)/.test(msg)) {
    return handleTravelEnd();
  }

  if (/(啟動旅遊|開啟旅遊|開始旅遊|幫我開啟|設定旅遊|旅遊模式|出國模式)/.test(msg) || msg.includes('之旅')) {
    return handleTravelStart(userMessage);
  }
  return processExpense(userMessage, userName);
}

function handlePostbackEvent(event, userName) {
  try {
    const dataParts = event.postback.data.split('&');
    const actionObj = {};
    dataParts.forEach(part => {
      const [k, v] = part.split('=');
      actionObj[k] = v;
    });

    if (actionObj.action === 'confirm') return confirmTransaction(actionObj.id, userName);
    if (actionObj.action === 'cancel') return cancelTransaction(actionObj.id, userName);
    // [v8.0] 刪除確認按鈕
    if (actionObj.action === 'delete_confirm') return executeDelete(actionObj.id, userName);
    if (actionObj.action === 'delete_cancel') {
      CacheService.getScriptCache().remove(`del_${actionObj.id}`);
      return { status: 'success', message: '✅ 已取消刪除，記帳資料保留。' };
    }

    return { status: 'error', message: '⚠️ 無效的操作' };
  } catch (err) {
    logError('handlePostbackEvent', err);
    return { status: 'error', message: '⚠️ 處理按鈕回傳時發生錯誤' };
  }
}

function confirmTransaction(uuid, _userName) {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(uuid);

  if (!cachedData) {
    return { status: 'error', message: '⚠️ 此筆確認資料已過期或已處理！' };
  }

  const record = JSON.parse(cachedData);
  const writeSuccess = writeToSheet(record);
  cache.remove(uuid);

  if (!writeSuccess) {
    return { status: 'error', message: '❌ 寫入這筆大額消費時失敗！請稍後再試。' };
  }

  const mergeText = record.source.includes('自動合併') ? '\n🔄 [系統已自動合併載具舊帳]' : '';
  const budgetAlert = checkBudgetAlert() || '';

  // [v7.0 fix] 補上子預算警告（大額確認按鈕路徑原本遺漏此邏輯）
  let projectBudgetAlert = '';
  if (CONFIG.PROJECT_BUDGETS && CONFIG.PROJECT_BUDGETS[record.project]) {
    const limit = CONFIG.PROJECT_BUDGETS[record.project];
    const d = new Date(record.date);
    const monthlyTotal = calculateProjectTotal(record.project, d.getMonth() + 1, d.getFullYear());
    const remaining = limit - monthlyTotal;
    const pct = (monthlyTotal / limit * 100).toFixed(1);
    if (remaining < 0) {
      projectBudgetAlert = `\n🚨 哎呀！本月這個專案已超支 ${formatCurrency(Math.abs(remaining))} 啦！(月上限 ${formatCurrency(limit)})`;
    } else if (pct > 80) {
      projectBudgetAlert = `\n⚠️ 【子預算提醒】本月已用 ${pct}%，額度只剩 ${formatCurrency(remaining)}！`;
    } else {
      projectBudgetAlert = `\n📊 【子預算】本月使用率：${pct}%，餘額 ${formatCurrency(remaining)}`;
    }
  }

  // [v7.3+] 大額確認路徑：同樣偵測分期並寫入固定收支
  let installmentNotice = '';
  if (record.installments && record.installments > 1) {
    const expDate = new Date(record.date);
    const installResult = addInstallmentToFixedSchedule({
      item: record.item, totalAmount: record.amount, installments: record.installments,
      day: record.installment_day || expDate.getDate(),
      category: record.category, project: record.project, payment: record.payment, expenseDate: expDate
    });
    if (installResult) {
      installmentNotice = `\n\n💳 【分期自動入帳】\n已自動新增至固定收支 ${record.installments} 期\n📅 每月 ${installResult.day} 號扣 ${formatCurrency(installResult.monthlyAmount)}\n⏳ 扣款至 ${installResult.end_month} 止`;
    }
  }

  const textMessage = `✅ 大額消費寫入成功！${mergeText}${projectBudgetAlert}${budgetAlert}${installmentNotice}`;
  return { status: 'success', message: textMessage, flexMessage: buildReceiptFlexMessage(record) };
}

function cancelTransaction(uuid, _userName) {
  const cache = CacheService.getScriptCache();
  cache.remove(uuid);
  return { status: 'success', message: '🗑️ 已取消該筆記帳動作' };
}

// ============================================
// [v7.7] 補記卡別：「這筆用XX卡」「上一筆用XX卡」
// ============================================
function handleUpdateLastEntryCard(userMessage, userName) {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: 'error', message: '⚠️ 目前沒有資料可以補記！' };

    // 解析使用者說的是哪張卡
    const cardsMap = CONFIG.CARDS_MAP;
    if (!cardsMap || Object.keys(cardsMap).length === 0) {
      return { status: 'error', message: '⚠️ 尚未設定 CARDS_MAP！\n請在指令碼屬性中新增 CARDS_MAP。' };
    }

    // 比對 CARDS_MAP 中的 bank 名或 cardKey
    let matchedCardKey = null;
    const msgLower = userMessage.toLowerCase();
    for (const [key, info] of Object.entries(cardsMap)) {
      if (msgLower.includes(info.bank.toLowerCase()) ||
          msgLower.includes(key.toLowerCase()) ||
          (info.last4 && msgLower.includes(info.last4))) {
        matchedCardKey = key;
        break;
      }
    }
    if (!matchedCardKey) {
      const cardList = Object.entries(cardsMap).map(([k, v]) => `• ${k}（${v.owner}）`).join('\n');
      return { status: 'error', message: `⚠️ 找不到符合的卡別！\n目前設定的卡：\n${cardList}` };
    }

    const lastRowData = sheet.getRange(lastRow, 1, 1, 10).getValues()[0];
    const lastDate = lastRowData[0];
    const lastAmount = parseInt(lastRowData[2]) || 0;
    const lastItem = String(lastRowData[1] || '');
    const existingCard = String(lastRowData[9] || '');

    // 檢查是否為多品項情況（同日期、同商家前綴有多筆）
    const merchantPrefix = lastItem.includes(' - ') ? lastItem.split(' - ')[0].replace(/股份有限公司|有限公司/g, '').substring(0, 6) : '';
    if (merchantPrefix) {
      const allData = sheet.getDataRange().getValues();
      const lastDateStr = Utilities.formatDate(new Date(lastDate), CONFIG.TIMEZONE, 'yyyy/MM/dd');
      const sameGroup = [];
      for (let i = 1; i < allData.length; i++) {
        const rDateStr = Utilities.formatDate(new Date(allData[i][0]), CONFIG.TIMEZONE, 'yyyy/MM/dd');
        const rItem = String(allData[i][1] || '');
        const rCard = String(allData[i][9] || '');
        if (rDateStr === lastDateStr && rItem.startsWith(merchantPrefix) && !rCard) {
          sameGroup.push(i + 1); // 1-indexed row number
        }
      }
      if (sameGroup.length > 1) {
        const groupTotal = sameGroup.reduce((sum, rowNum) => {
          return sum + (parseInt(allData[rowNum - 1][2]) || 0);
        }, 0);
        // 存入快取等待確認
        const cacheKey = `card_batch_${Date.now()}`;
        CacheService.getScriptCache().put(cacheKey, JSON.stringify({ rows: sameGroup, cardKey: matchedCardKey }), 300);
        return {
          status: 'card_batch_confirm',
          message: `🔍 找到 ${sameGroup.length} 筆【${merchantPrefix}…】同日消費，合計 ${formatCurrency(groupTotal)}\n\n確認全部標記為「${matchedCardKey}」嗎？\n\n回傳「確認標記卡別 ${cacheKey}」確認，或直接忽略取消`
        };
      }
    }

    // 單筆直接更新
    sheet.getRange(lastRow, 10).setValue(matchedCardKey);
    const lineSep = '━━━━━━━━━━';
    const oldStr = existingCard ? `（原：${existingCard}）` : '';
    return {
      status: 'success',
      message: `💳 卡別補記完成！${lineSep}\n📝 ${lastItem}\n💰 ${formatCurrency(lastAmount)}\n💳 卡別：${matchedCardKey} ${oldStr}`
    };
  } catch (error) {
    logError('handleUpdateLastEntryCard', error);
    return { status: 'error', message: '❌ 補記卡別失敗：' + error.toString() };
  }
}

// [v7.7] 整批標記卡別確認
function confirmBatchCardTag(userMessage, userName) {
  try {
    // 從訊息中提取 cacheKey
    const keyMatch = userMessage.match(/確認標記卡別\s+(\S+)/);
    if (!keyMatch) return { status: 'error', message: '⚠️ 指令格式錯誤，請重新操作。' };

    const cacheKey = keyMatch[1];
    const cached = CacheService.getScriptCache().get(cacheKey);
    if (!cached) return { status: 'error', message: '⚠️ 確認請求已逾時（5 分鐘），請重新輸入。' };

    const { rows, cardKey } = JSON.parse(cached);
    const sheet = getSheet();
    rows.forEach(rowNum => sheet.getRange(rowNum, 10).setValue(cardKey));
    CacheService.getScriptCache().remove(cacheKey);

    return {
      status: 'success',
      message: `✅ 已將 ${rows.length} 筆消費標記為「${cardKey}」！`
    };
  } catch (error) {
    logError('confirmBatchCardTag', error);
    return { status: 'error', message: '❌ 整批標記失敗：' + error.toString() };
  }
}
