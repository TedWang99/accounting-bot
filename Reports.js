// ============================================
// Reports - LINE 家庭記帳機器人 v7.8
// ============================================

function handleQueryBudgets() {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const d = new Date();
    const currentMonth = d.getMonth() + 1;
    const currentYear = d.getFullYear();

    // 1. & 2. 使用底層核心計算所有數據，保證與本月、報表指令完全一致
    const stats = aggregateMonthlyData(data, currentMonth, currentYear);
    const budgetData = generateBudgetData(currentMonth, stats);

    // fallback 純文字訊息
    let textMsg = `🎯 【${currentMonth}月份】預算狀態總覽\n━━━━━━━━━━\n`;
    textMsg += `💰 本月總進度：${budgetData.main.pct.toFixed(1)}% (餘額 ${formatCurrency(budgetData.main.remaining)})\n`;
    budgetData.projects.forEach(p => {
      textMsg += `📌 ${p.name}：${p.pct.toFixed(1)}% (餘額 ${formatCurrency(p.remaining)})\n`;
    });

    return {
      status: 'success',
      message: textMsg.trim(),
      flexMessage: buildBudgetFlexMessage(budgetData)
    };
  } catch (e) {
    logError('handleQueryBudgets', e);
    return { status: 'error', message: '❌ 查詢子預算失敗：' + e.toString() };
  }
}

function getMonthlyTotal(targetMonth = null, targetYear = null) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const current = getCurrentMonthYear();
    const month = targetMonth || current.month;
    const year = targetYear || current.year;

    const stats = aggregateMonthlyData(data, month, year);
    return formatMonthlyTotalMessage(stats, month);
  } catch (error) { logError('getMonthlyTotal', error); return { status: 'error', message: '❌ 查詢失敗：' + error.toString() }; }
}

function formatMonthlyTotalMessage(stats, month) {
  const { totalExpense, totalIncome, projectStats } = stats;
  const remaining = CONFIG.MONTHLY_BUDGET - totalExpense;
  const percentage = CONFIG.MONTHLY_BUDGET > 0 ? (totalExpense / CONFIG.MONTHLY_BUDGET * 100) : 0;

  let alert = totalExpense > CONFIG.MONTHLY_BUDGET
    ? `\n🚨 警告：已超支 ${formatCurrency(Math.abs(remaining))}！`
    : `\n✅ 預算剩餘：${formatCurrency(remaining)} (${(100 - percentage).toFixed(1)}%)`;

  const incomeStr = totalIncome > 0 ? `\n📥 總收入：${formatCurrency(totalIncome)}` : '';
  const lineSep = '━━━━━━━━━━';
  let message = `📊 ${month}月 收支摘要\n${lineSep}${incomeStr}\n💸 一般支出：${formatCurrency(totalExpense)}\n📈 預算使用率：${percentage.toFixed(1)}%\n${lineSep}${alert}`;

  // 使用共用的 budgetData
  const budgetData = generateBudgetData(month, stats);

  const specialProjects = Object.entries(projectStats)
    .filter(([name, amt]) => name !== CONFIG.DEFAULT_PROJECT && amt > 0)
    .sort((a, b) => b[1] - a[1]);
  if (specialProjects.length > 0) {
    message += '\n\n📌 特殊專案（不含在預算內）';
    specialProjects.forEach(([name, amt]) => {
      message += `\n• ${name}：${formatCurrency(amt)}`;
    });
  }

  return {
    status: 'success',
    message: message.trim(),
    flexMessage: CONFIG.MONTHLY_BUDGET > 0 ? buildBudgetFlexMessage(budgetData) : null
  };
}

function getLastEntry() {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return '⚠️ 目前還沒有任何記帳紀錄喔！';

    // 取得最後50筆資料來檢查是否有同發票的明細 (為了效能不全讀)
    const startRow = Math.max(2, lastRow - 50);
    const numRows = lastRow - startRow + 1;
    const data = sheet.getRange(startRow, 1, numRows, 10).getValues();

    const lastRecord = data[data.length - 1];
    const targetInvoice = lastRecord[7];

    let totalAmount = 0;
    let items = [];
    let recordCount = 0;

    // 如果最後一筆有發票號碼，往前找同發票的所有明細
    if (targetInvoice) {
      for (let i = data.length - 1; i >= 0; i--) {
        if (data[i][7] === targetInvoice) {
          totalAmount += parseInt(data[i][2]) || 0;
          items.push(data[i][1].split('-').pop().trim()); // 只取品名摘要
          recordCount++;
        }
      }
    } else {
      // 沒有發票號碼，就只顯示最後單筆
      totalAmount = parseInt(lastRecord[2]) || 0;
      items.push(lastRecord[1]);
      recordCount = 1;
    }

    const dateStr = Utilities.formatDate(new Date(lastRecord[0]), CONFIG.TIMEZONE, 'MM/dd');
    const lineSep = '━━━━━━━━━━';

    let itemDisplay = items[0];
    if (recordCount > 1) {
      itemDisplay = `${lastRecord[1].split('-')[0].trim()} (共 ${recordCount} 項明細)`;
    }

    const cardId = String(lastRecord[9] || '').trim();  // [v7.8] col 10 卡別
    const cardStr = cardId ? `\n💳 卡別：${cardId}` : '';
    let msg = `🔍 上一筆記帳詳情\n${lineSep}\n📅 日期：${dateStr}\n📝 項目：${itemDisplay}\n📂 分類：${lastRecord[5]}\n🚀 專案：${lastRecord[6]}\n💳 支付：${lastRecord[4]}${cardStr}\n💰 金額：${formatCurrency(totalAmount)}\n👤 記錄人：${lastRecord[3]}`;

    if (targetInvoice && recordCount > 1) {
      msg += `\n🧾 發票：${targetInvoice}`;
    }
    msg += `\n${lineSep}`;
    return msg;
  } catch (error) { logError('getLastEntry', error); return '❌ 查詢失敗：' + error.toString(); }
}

function handleProjectQuery(userMessage) {
  const keyword = userMessage.split(/\s+/)[1];
  if (!keyword) return { status: 'error', message: '⚠️ 請輸入要查詢的關鍵字\n例如：查 飲食、查 全聯、查 養車' };
  // [v7.9] 同時回傳 flexMessage 以顯示 Flex 卡片
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const result = aggregateProjectData(data, keyword);
  if (result.count === 0) return { status: 'error', message: `⚠️ 找不到有關「${keyword}」的紀錄\n請確認關鍵字是否正確\n\n💡 提示：可搜尋分類、專案或特定項目名稱` };
  const textMsg = formatProjectQueryResult(keyword, result);
  const flexMsg = buildSearchResultFlexMessage(keyword, result);
  return { status: 'query_project', message: textMsg, flexMessage: flexMsg };
}

function queryProjectTotal(keyword) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const result = aggregateProjectData(data, keyword);
    if (result.count === 0) return `⚠️ 找不到有關「${keyword}」的紀錄\n請確認關鍵字是否正確\n\n💡 提示：可搜尋分類、專案或特定項目名稱`;
    return formatProjectQueryResult(keyword, result);
  } catch (error) { logError('queryProjectTotal', error); return '❌ 查詢失敗：' + error.toString(); }
}

function aggregateProjectData(data, keyword) {
  let total = 0, count = 0;
  const matchedGroups = {}; // 用 invoiceNum (或 date+item 如果沒 invoice) 當 key

  for (let i = 1; i < data.length; i++) {
    const dateStr = Utilities.formatDate(new Date(data[i][0]), CONFIG.TIMEZONE, 'MM/dd');
    const itemDesc = String(data[i][1] || '');
    const amount = parseInt(data[i][2]) || 0;
    const category = String(data[i][5] || '');
    const project = String(data[i][6] || '');
    const invoiceNum = String(data[i][7] || '');

    if (itemDesc.includes(keyword) || category.includes(keyword) || project.includes(keyword)) {
      total += amount;
      count++;

      const groupKey = invoiceNum ? `INV_${invoiceNum}` : `FB_${i}`;

      if (!matchedGroups[groupKey]) {
        matchedGroups[groupKey] = {
          date: dateStr,
          storeName: itemDesc.split('-')[0].trim(),
          amount: 0,
          itemCount: 0,
          originalItem: itemDesc
        };
      }
      matchedGroups[groupKey].amount += amount;
      matchedGroups[groupKey].itemCount++;
    }
  }

  // 整理成陣列並取最後8筆顯示
  const allMatched = Object.values(matchedGroups).map(g => {
    const displayItem = g.itemCount > 1 ? `${g.storeName} (共 ${g.itemCount} 項)` : g.originalItem;
    return { date: g.date, item: displayItem, amount: g.amount };
  });

  const items = allMatched.slice(-8).reverse();
  // 注意這裡回傳的 count 仍然是「真實的明細總筆數」，但畫面上 items 會折疊
  return { total, count, foldedCount: allMatched.length, items };
}

function formatProjectQueryResult(keyword, result) {
  const { total, count, foldedCount, items } = result;
  const lineSep = '━━━━━━━━━━';
  let message = `🔍 關鍵字查詢：${keyword}\n${lineSep}\n💰 累計花費：${formatCurrency(total)}\n📝 相關筆數：${foldedCount} 筆消費 (共 ${count} 項明細)\n${lineSep}`;
  if (items.length > 0) {
    message += '\n\n📋 最近紀錄：';
    items.forEach((item, index) => { message += `\n${index + 1}. ${item.date} ${item.item} $${item.amount}`; });
  }
  return message;
}

function getMonthlyReport(targetMonth = null, targetYear = null) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const current = getCurrentMonthYear();
    const month = targetMonth || current.month;
    const year = targetYear || current.year;

    const stats = aggregateMonthlyData(data, month, year);
    const aiAdvice = getAIFinancialAdvice(month, stats); // [v7.0] AI 財務顧問
    const message = formatMonthlyReport(month, year, stats) + `\n\n💡 AI 顧問點評：\n${aiAdvice}`;
    const imageUrl = generatePieChartUrl(stats.categoryStats, `${month}月各項支出佔比`);
    const flexReport = buildReportFlexMessage(month, stats, imageUrl, aiAdvice);

    // [v6.0/v7.0] 優先傳送 Flex Message 與圖片
    return { status: 'report', message, imageUrl, flexMessage: flexReport };
  } catch (error) {
    logError('getMonthlyReport', error);
    return { status: 'error', message: '❌ 報表產生失敗：' + error.toString() };
  }
}

function aggregateMonthlyData(data, month, year) {
  const categoryStats = {}, projectStats = {}, paymentStats = {}, cardStats = {}, personStats = {};
  let totalExpense = 0, totalIncome = 0, count = 0;

  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][0]);
    if (!isCurrentMonth(rowDate, month, year)) continue;

    const amount = parseInt(data[i][2]) || 0;
    const category = data[i][5] || CONFIG.DEFAULT_CATEGORY;
    const project = data[i][6] || CONFIG.DEFAULT_PROJECT;
    const payment = data[i][4] || CONFIG.DEFAULT_PAYMENT;
    const cardId = String(data[i][9] || '').trim();  // [v7.7] col 10 卡別
    const recorder = String(data[i][3] || '').trim(); // [v7.8] col 4 記錄人
    count++;

    if (category === '工資') {
      totalIncome += amount;
    } else {
      projectStats[project] = (projectStats[project] || 0) + amount;
      if (project === CONFIG.DEFAULT_PROJECT) {
        totalExpense += amount;
        categoryStats[category] = (categoryStats[category] || 0) + amount;
        paymentStats[payment] = (paymentStats[payment] || 0) + amount;
      }
      // [v7.7] 各卡統計（含所有專案，只含支出正數）
      if (cardId && amount > 0) {
        cardStats[cardId] = (cardStats[cardId] || 0) + amount;
      }
      // [v7.8] 各人統計（含所有專案，只含支出正數）
      if (recorder && amount > 0) {
        personStats[recorder] = (personStats[recorder] || 0) + amount;
      }
    }
  }
  return { total: totalExpense, totalExpense, totalIncome, count, categoryStats, projectStats, paymentStats, cardStats, personStats };
}

// [v7.7] 按卡別查詢：解析「查 國泰 本月」「查 富邦Costco 上月」「查 台新」等指令
function parseCardQuery(queryText) {
  const cardsMap = CONFIG.CARDS_MAP;
  if (!cardsMap || Object.keys(cardsMap).length === 0) return null;

  let matchedKeys = [];
  let matchLabel = '';
  let remaining = queryText;

  // 優先比對完整 cardKey（較長的優先，避免「末1234」誤匹配「末123」）
  const sortedKeys = Object.keys(cardsMap).sort((a, b) => b.length - a.length);
  for (const key of sortedKeys) {
    if (remaining.includes(key)) {
      matchedKeys = [key];
      matchLabel = key;
      remaining = remaining.replace(key, '').trim();
      break;
    }
  }
  // 若沒有完整 key，改用 bank 名稱比對
  if (matchedKeys.length === 0) {
    const banksSorted = Object.entries(cardsMap).sort((a, b) => b[1].bank.length - a[1].bank.length);
    for (const [key, info] of banksSorted) {
      if (info.bank && remaining.includes(info.bank)) {
        matchedKeys.push(key);
      }
    }
    if (matchedKeys.length > 0) {
      // 找最長的 bank 名進行移除
      const longestBank = matchedKeys.map(k => cardsMap[k].bank).sort((a, b) => b.length - a.length)[0];
      matchLabel = longestBank;
      remaining = remaining.replace(longestBank, '').trim();
    }
  }
  if (matchedKeys.length === 0) return null;

  // 解析剩餘部分的日期（parseDateRangeInput 不支援本月/上月，需自行處理）
  let dateRange = null;
  if (/本月/.test(remaining)) {
    const now = new Date();
    dateRange = {
      startDate: new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0),
      endDate: new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59, 999),
      label: '本月'
    };
  } else if (/上月/.test(remaining)) {
    const now = new Date();
    dateRange = {
      startDate: new Date(now.getFullYear(), now.getMonth() - 1, 1, 0, 0, 0, 0),
      endDate: new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59, 999),
      label: '上月'
    };
  } else {
    dateRange = parseDateRangeInput(remaining.trim());
  }
  // 若沒有日期關鍵字，預設本月
  if (!dateRange) {
    const now = new Date();
    dateRange = {
      startDate: new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0),
      endDate: new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59, 999),
      label: '本月'
    };
  }
  return { matchedKeys, matchLabel, dateRange };
}

function handleCardQuery(userMessage, userName) {
  try {
    const queryText = userMessage.replace(/^查詢?\s+/, '').trim();
    const parsed = parseCardQuery(queryText);
    if (!parsed) return { status: 'error', message: '⚠️ 查不到符合的卡別，請確認 CARDS_MAP 設定。' };

    const { matchedKeys, matchLabel, dateRange } = parsed;
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const start = dateRange.startDate;
    const end = dateRange.endDate;

    const lineSep = '━━━━━━━━━━';
    const dateLabel = dateRange.label || `${Utilities.formatDate(start, CONFIG.TIMEZONE, 'MM/dd')} ~ ${Utilities.formatDate(end, CONFIG.TIMEZONE, 'MM/dd')}`;

    let total = 0;
    const items = [];

    for (let i = 1; i < data.length; i++) {
      const rowDate = new Date(data[i][0]);
      if (isNaN(rowDate) || rowDate < start || rowDate > end) continue;
      const cardId = String(data[i][9] || '').trim();
      if (!matchedKeys.some(k => cardId === k)) continue;
      const amount = parseInt(data[i][2]) || 0;
      if (amount <= 0) continue;  // 排除負數（折扣/退款）
      total += amount;
      items.push({ date: rowDate, item: String(data[i][1] || ''), amount });
    }

    if (items.length === 0) {
      return { status: 'success', message: `💳 ${matchLabel}（${dateLabel}）\n${lineSep}\n查無消費記錄。` };
    }

    // 最近 8 筆（倒序）
    const recent = items.sort((a, b) => b.date - a.date).slice(0, 8);
    const itemLines = recent.map(it =>
      `${Utilities.formatDate(it.date, CONFIG.TIMEZONE, 'MM/dd')}  ${it.item}  NT$${it.amount.toLocaleString()}`
    ).join('\n');
    const moreHint = items.length > 8 ? `\n  ...共 ${items.length} 筆，僅顯示最近 8 筆` : '';

    const msg = `💳 ${matchLabel}（${dateLabel}）\n${lineSep}\n${itemLines}${moreHint}\n${lineSep}\n合計 NT$ ${total.toLocaleString()}（共 ${items.length} 筆）`;
    return { status: 'success', message: msg };
  } catch (error) {
    logError('handleCardQuery', error);
    return { status: 'error', message: '❌ 查詢失敗：' + error.toString() };
  }
}

function getAIFinancialAdvice(month, stats) {
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    const topCat = Object.entries(stats.categoryStats)
      .filter(([k, _]) => k !== '工資')
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(x => `${x[0]}:${x[1]}元`)
      .join(', ');

    const prompt = `你是一個專業、親切且幽默的家庭財務顧問。請根據以下 ${month} 月份的收支統計，寫一段「50~100字」的總結評語。如果特定項目花太多可以稍微幽默吐槽，若有控制在預算內(${CONFIG.MONTHLY_BUDGET})請給予大大稱讚。
    【本月數據】
    總支出：${stats.totalExpense}
    總收入：${stats.totalIncome}
    前三大開銷：${topCat}
    請直接回傳純文字評語，不用包含任何開頭稱呼或 Markdown 標記，直接講重點。`;

    const response = fetchWithRetry(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.7 } })
    });

    const result = safeJSONParse(response.getContentText());
    if (result.error) throw new Error(result.error.message);
    return (result.candidates?.[0]?.content?.parts?.[0]?.text || '').replace(/```/g, '').trim();
  } catch (error) {
    logError('getAIFinancialAdvice', error);
    return "您的 AI 顧問趕去休假了，稍後再為您分析報表！";
  }
}

function getCategoryStats(targetMonth = null, targetYear = null) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const current = getCurrentMonthYear();
    const month = targetMonth || current.month;
    const year = targetYear || current.year;

    return formatCategoryStats(aggregateCategoryStats(data, month, year), month);
  } catch (error) { logError('getCategoryStats', error); return '❌ 統計失敗：' + error.toString(); }
}

function aggregateCategoryStats(data, month, year) {
  const stats = {};
  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][0]);
    if (!isCurrentMonth(rowDate, month, year)) continue;
    const category = data[i][5] || CONFIG.DEFAULT_CATEGORY;
    if (category === '工資') continue; // [v5.5 fix] 排除收入，避免工資混入支出統計排行
    const amount = parseInt(data[i][2]) || 0;
    if (!stats[category]) stats[category] = { total: 0, count: 0 };
    stats[category].total += amount;
    stats[category].count++;
  }
  return stats;
}

function formatCategoryStats(stats, month = null) {
  const lineSep = '━━━━━━━━━━';
  const monthTitle = month ? `${month}月` : '本月';
  let message = `📊 ${monthTitle} 分類統計\n${lineSep}\n`;
  Object.entries(stats).sort((a, b) => b[1].total - a[1].total).forEach(([category, data]) => {
    message += `\n${category}\n  💰 ${formatCurrency(data.total)}\n  📝 ${data.count} 筆 (均 ${formatCurrency(Math.round(data.total / data.count))})\n`;
  });
  return message;
}

function parseDateRangeInput(text) {
  const t = text.trim();

  // 上週
  if (/^上週?$/.test(t)) {
    const now = new Date();
    const dayOfWeek = now.getDay() || 7; // 1=Mon … 7=Sun
    const lastMonday = new Date(now);
    lastMonday.setDate(now.getDate() - dayOfWeek - 6);
    lastMonday.setHours(0, 0, 0, 0);
    const lastSunday = new Date(lastMonday);
    lastSunday.setDate(lastMonday.getDate() + 6);
    lastSunday.setHours(23, 59, 59, 999);
    const fmt = d => Utilities.formatDate(d, CONFIG.TIMEZONE, 'MM/dd');
    return { startDate: lastMonday, endDate: lastSunday, label: `上週（${fmt(lastMonday)} ~ ${fmt(lastSunday)}）` };
  }

  // 本週
  if (/^本週?$/.test(t)) {
    const now = new Date();
    const dayOfWeek = now.getDay() || 7;
    const monday = new Date(now);
    monday.setDate(now.getDate() - dayOfWeek + 1);
    monday.setHours(0, 0, 0, 0);
    const today = new Date(now);
    today.setHours(23, 59, 59, 999);
    const fmt = d => Utilities.formatDate(d, CONFIG.TIMEZONE, 'MM/dd');
    return { startDate: monday, endDate: today, label: `本週（${fmt(monday)} ~ ${fmt(today)}）` };
  }

  // YYYY/MM（整月）
  const monthOnly = t.match(/^(\d{4})[\/\-](\d{1,2})$/);
  if (monthOnly) {
    const y = parseInt(monthOnly[1]), m = parseInt(monthOnly[2]) - 1;
    const start = new Date(y, m, 1, 0, 0, 0, 0);
    const end = new Date(y, m + 1, 0, 23, 59, 59, 999);
    return { startDate: start, endDate: end, label: `${y}年${m + 1}月` };
  }

  // YYYY/MM/DD ~ YYYY/MM/DD（含 ～ 或「至」）
  const rangeMatch = t.match(/^(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})\s*[~～至\-]\s*(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})$/);
  if (rangeMatch) {
    const start = new Date(rangeMatch[1].replace(/\//g, '-') + 'T00:00:00+08:00');
    const end = new Date(rangeMatch[2].replace(/\//g, '-') + 'T23:59:59+08:00');
    if (isNaN(start) || isNaN(end)) return null;
    const fmt = d => Utilities.formatDate(d, CONFIG.TIMEZONE, 'MM/dd');
    return { startDate: start, endDate: end, label: `${fmt(start)} ~ ${fmt(end)}` };
  }

  return null;
}

// [v7.8] 查詢分類預算狀態
function handleQueryCategoryBudgets() {
  try {
    const catBudgets = CONFIG.CATEGORY_BUDGETS;
    if (!catBudgets || Object.keys(catBudgets).length === 0) {
      return { status: 'info', message: '⚠️ 尚未設定分類預算！\n請在指令碼屬性中新增 CATEGORY_BUDGETS\n格式：{"飲食": 15000, "娛樂": 3000}' };
    }
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const { month, year } = getCurrentMonthYear();
    const stats = aggregateMonthlyData(data, month, year);

    const lineSep = '━━━━━━━━━━';
    let msg = `🏷️ ${month}月 分類預算狀態\n${lineSep}\n`;
    Object.entries(catBudgets).sort((a, b) => b[1] - a[1]).forEach(([cat, limit]) => {
      const spent = stats.categoryStats[cat] || 0;
      const pct = (spent / limit * 100).toFixed(1);
      const remaining = limit - spent;
      const icon = spent > limit ? '🚨' : (pct > 80 ? '⚠️' : '✅');
      msg += `\n${icon} ${cat}：${formatCurrency(spent)} / ${formatCurrency(limit)} (${pct}%)`;
      if (remaining < 0) msg += `  超支 ${formatCurrency(Math.abs(remaining))}！`;
    });
    return { status: 'success', message: msg.trim() };
  } catch (e) {
    logError('handleQueryCategoryBudgets', e);
    return { status: 'error', message: '❌ 查詢分類預算失敗：' + e.toString() };
  }
}

function aggregateDateRangeData(data, startDate, endDate) {
  const categoryStats = {}, projectStats = {}, paymentStats = {};
  let totalExpense = 0, totalIncome = 0, count = 0;

  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][0]);
    if (isNaN(rowDate) || rowDate < startDate || rowDate > endDate) continue;

    const amount = parseInt(data[i][2]) || 0;
    const category = data[i][5] || CONFIG.DEFAULT_CATEGORY;
    const project = data[i][6] || CONFIG.DEFAULT_PROJECT;
    const payment = data[i][4] || CONFIG.DEFAULT_PAYMENT;
    count++;

    if (category === '工資') {
      totalIncome += amount;
    } else {
      projectStats[project] = (projectStats[project] || 0) + amount;
      if (project === CONFIG.DEFAULT_PROJECT) {
        totalExpense += amount;
        categoryStats[category] = (categoryStats[category] || 0) + amount;
        paymentStats[payment] = (paymentStats[payment] || 0) + amount;
      }
    }
  }
  return { totalExpense, totalIncome, count, categoryStats, projectStats, paymentStats };
}

function handleDateRangeQuery(userMessage) {
  try {
    const keyword = userMessage.replace(/^查詢?\s+/i, '').trim();
    const range = parseDateRangeInput(keyword);
    if (!range) return { status: 'error', message: '⚠️ 無法解析日期範圍！\n支援格式：\n• 查 上週\n• 查 本週\n• 查 2026/01\n• 查 2026/01/01 ~ 2026/01/31' };

    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const stats = aggregateDateRangeData(data, range.startDate, range.endDate);

    if (stats.count === 0) return { status: 'success', message: `📊 ${range.label}\n⚠️ 此區間沒有任何記帳資料。` };

    const lineSep = '━━━━━━━━━━';
    const topCats = Object.entries(stats.categoryStats).sort((a, b) => b[1] - a[1]).slice(0, 3);
    const incomeStr = stats.totalIncome > 0 ? `\n📥 收入：${formatCurrency(stats.totalIncome)}` : '';
    let msg = `📊 ${range.label}\n${lineSep}${incomeStr}\n💸 支出：${formatCurrency(stats.totalExpense)}\n📝 共 ${stats.count} 筆\n${lineSep}`;

    if (topCats.length > 0) {
      msg += '\n🏆 支出分類 Top 3：';
      topCats.forEach(([cat, amt], i) => {
        const pct = stats.totalExpense > 0 ? (amt / stats.totalExpense * 100).toFixed(1) : '0.0';
        msg += `\n${i + 1}. ${cat}：${formatCurrency(amt)} (${pct}%)`;
      });
    }

    const specialProjects = Object.entries(stats.projectStats)
      .filter(([name, amt]) => name !== CONFIG.DEFAULT_PROJECT && amt > 0)
      .sort((a, b) => b[1] - a[1]);
    if (specialProjects.length > 0) {
      msg += `\n${lineSep}\n📌 特殊專案：`;
      specialProjects.forEach(([name, amt]) => { msg += `\n• ${name}：${formatCurrency(amt)}`; });
    }

    return { status: 'success', message: msg };
  } catch (error) {
    logError('handleDateRangeQuery', error);
    return { status: 'error', message: '❌ 查詢失敗：' + error.toString() };
  }
}
