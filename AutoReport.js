// ============================================
// AutoReport - LINE 家庭記帳機器人 v7.8
// ============================================

function generateBudgetData(month, stats) {
  const { totalExpense, projectStats } = stats;
  const remaining = CONFIG.MONTHLY_BUDGET - totalExpense;
  const percentage = CONFIG.MONTHLY_BUDGET > 0 ? (totalExpense / CONFIG.MONTHLY_BUDGET * 100) : 0;

  const budgetData = {
    month: month,
    main: {
      spent: totalExpense,
      limit: CONFIG.MONTHLY_BUDGET || 0,
      pct: percentage,
      remaining: remaining
    },
    projects: []
  };

  // 永遠列出所有設定了子預算的專案 (即使本月花費為 0)，確保圖卡一體性
  if (CONFIG.PROJECT_BUDGETS && Object.keys(CONFIG.PROJECT_BUDGETS).length > 0) {
    for (const [name, limit] of Object.entries(CONFIG.PROJECT_BUDGETS)) {
      const amt = projectStats[name] || 0;
      budgetData.projects.push({
        name: name,
        spent: amt,
        limit: limit,
        pct: limit > 0 ? (amt / limit * 100) : 0,
        remaining: limit - amt
      });
    }
  }

  // [v7.8] 分類預算
  const catBudgets = CONFIG.CATEGORY_BUDGETS;
  budgetData.categories = [];
  if (catBudgets && Object.keys(catBudgets).length > 0) {
    for (const [catName, limit] of Object.entries(catBudgets)) {
      const amt = stats.categoryStats[catName] || 0;
      budgetData.categories.push({
        name: catName,
        spent: amt,
        limit: limit,
        pct: limit > 0 ? (amt / limit * 100) : 0,
        remaining: limit - amt
      });
    }
  }

  return budgetData;
}

// [v7.8] 分類預算警示（記帳後呼叫）
function checkCategoryBudgetAlert(category) {
  try {
    const catBudgets = CONFIG.CATEGORY_BUDGETS;
    if (!catBudgets || !catBudgets[category]) return null;
    const limit = catBudgets[category];
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const { month, year } = getCurrentMonthYear();
    let total = 0;
    for (let i = 1; i < data.length; i++) {
      const rowDate = new Date(data[i][0]);
      if (!isCurrentMonth(rowDate, month, year)) continue;
      if (String(data[i][5]) === category) total += parseInt(data[i][2]) || 0;
    }
    const pct = (total / limit * 100);
    if (total > limit) return `\n🚨 【${category}預算】本月已超支 ${formatCurrency(total - limit)}！（上限 ${formatCurrency(limit)}）`;
    if (pct > 80) return `\n⚠️ 【${category}預算】本月已用 ${pct.toFixed(1)}%（${formatCurrency(total)} / ${formatCurrency(limit)}）`;
    return null;
  } catch (e) { return null; }
}

function checkBudgetAlert() {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const { month, year } = getCurrentMonthYear();
    let total = 0;
    for (let i = 1; i < data.length; i++) {
      const rowDate = new Date(data[i][0]);
      if (isCurrentMonth(rowDate, month, year)) {
        const cat = data[i][5] || '';
        const proj = data[i][6] || CONFIG.DEFAULT_PROJECT;
        if (cat !== '工資' && proj === CONFIG.DEFAULT_PROJECT) {
          total += parseInt(data[i][2]) || 0;
        }
      }
    }
    const percentage = (total / CONFIG.MONTHLY_BUDGET * 100);
    if (total > CONFIG.MONTHLY_BUDGET) return `\n⚠️ 本月已超支 ${formatCurrency(total - CONFIG.MONTHLY_BUDGET)}！`;
    else if (percentage > 90) return `\n⚠️ 預算即將用完！已使用 ${percentage.toFixed(1)}%`;
    else if (percentage > 80) return `\n💡 提醒：預算已使用 ${percentage.toFixed(1)}%`;
    return null;
  } catch (error) { return null; }
}

// ============================================
// 📊 圖表產生器 (QuickChart)
// ============================================
function generatePieChartUrl(categoryStats, title) {
  try {
    // 過濾掉金額為0或工資(收入)的項目
    const validData = Object.entries(categoryStats)
      .filter(([name, amount]) => amount > 0 && name !== '工資')
      .sort((a, b) => b[1] - a[1]);

    if (validData.length === 0) return null;

    const labels = validData.map(d => d[0]);
    const data = validData.map(d => d[1]);

    // 生成隨機色彩帶或是使用特定色彩策略
    const reqData = {
      type: 'outlabeledPie',
      data: {
        labels: labels,
        datasets: [{
          backgroundColor: [
            "#FF6384", "#36A2EB", "#FFCE56", "#4BC0C0", "#9966FF",
            "#FF9F40", "#C9CBCF", "#00A6B4", "#F3A183", "#28B463"
          ],
          data: data
        }]
      },
      options: {
        title: {
          display: true,
          text: title,
          fontSize: 24,
          fontColor: '#333'
        },
        plugins: {
          legend: false,
          outlabels: {
            text: '%l %p\n$%v',
            color: 'white',
            stretch: 35,
            font: {
              resizable: true,
              minSize: 14,
              maxSize: 18
            }
          }
        }
      }
    };

    const chartUrl = `https://quickchart.io/chart?c=${encodeURIComponent(JSON.stringify(reqData))}&w=600&h=400&bkg=white`;
    return chartUrl;
  } catch (error) {
    logError('generatePieChartUrl', error);
    return null;
  }
}

function sendMonthlyReport() {
  const now = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const month = lastMonth.getMonth() + 1;
  const year = lastMonth.getFullYear();
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const stats = aggregateMonthlyData(data, month, year);
  const aiAdvice = getAIFinancialAdvice(month, stats); // [v7.0] 取得 AI 建議

  let message = formatAutoMonthlyReport(month, year, stats);
  message += `\n━━━━━━━━━━\n💡 AI 顧問點評：\n${aiAdvice}`;

  const chartUrl = generatePieChartUrl(stats.categoryStats, `${month}月各項支出佔比`);
  const flexReport = buildReportFlexMessage(month, stats, chartUrl, aiAdvice);

  Object.keys(CONFIG.USERS).forEach(userId => {
    pushLine(userId, message, chartUrl, flexReport);
  });
  Logger.log(`[月報推播] ${year}/${month} 月報與圖表已推播`);
}

function formatAutoMonthlyReport(month, year, stats) {
  const { total, count, categoryStats, projectStats, paymentStats } = stats;
  const budget = CONFIG.MONTHLY_BUDGET;
  const overBudget = total > budget;
  const percentage = (total / budget * 100).toFixed(1);
  const daysInMonth = new Date(year, month, 0).getDate();
  const avgDaily = Math.round(total / daysInMonth);
  const lineSep = '━━━━━━━━━━';
  const topCategories = Object.entries(categoryStats).sort((a, b) => b[1] - a[1]).slice(0, 3);
  const topProjects = Object.entries(projectStats).filter(([name]) => name !== CONFIG.DEFAULT_PROJECT).sort((a, b) => b[1] - a[1]).slice(0, 3);

  let message = `📊 ${year}年 ${month}月份消費報表\n${lineSep}\n💰 總支出：${formatCurrency(total)}\n📝 總筆數：${count} 筆\n📊 日均支出：${formatCurrency(avgDaily)}\n📈 預算使用率：${percentage}%\n`;
  if (overBudget) message += `🚨 警告：已超支 ${formatCurrency(total - budget)}！\n`;
  else message += `✅ 預算剩餘：${formatCurrency(budget - total)} (${(100 - parseFloat(percentage)).toFixed(1)}%)\n`;

  message += `\n🏆 支出排行 TOP 3：\n`;
  topCategories.forEach((item, index) => {
    const pct = total > 0 ? (item[1] / total * 100).toFixed(1) : '0.0';
    message += `${index + 1}. ${item[0]}：${formatCurrency(item[1])} (${pct}%)\n`;
  });

  if (topProjects.length > 0) {
    message += `\n🚀 特殊專案支出：\n`;
    topProjects.forEach(([proj, amt]) => { message += `• ${proj}：${formatCurrency(amt)}\n`; });
  }
  message += `\n💳 支付方式：\n`;
  Object.entries(paymentStats).sort((a, b) => b[1] - a[1]).forEach(([key, value]) => { message += `• ${key}：${formatCurrency(value)}\n`; });

  const specialProjects = Object.entries(projectStats).filter(([name, amt]) => name !== CONFIG.DEFAULT_PROJECT && amt > 0).sort((a, b) => b[1] - a[1]);
  if (specialProjects.length > 0) {
    message += `\n📌 特殊專案（不含在預算內）：\n`;
    specialProjects.forEach(([name, amt]) => { message += `• ${name}：${formatCurrency(amt)}\n`; });
  }
  message += `${lineSep}\n輸入「分類」可查看詳細統計`;
  return message;
}

function setupMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendMonthlyReport') ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger('sendMonthlyReport').timeBased().onMonthDay(1).atHour(8).create();
  SpreadsheetApp.getUi().alert('✅ 已成功設定每月 1 號早上 8:00 自動推播月報！');
}

function removeAllTriggers() {
  let count = 0;
  ScriptApp.getProjectTriggers().forEach(t => { ScriptApp.deleteTrigger(t); count++; });
  SpreadsheetApp.getUi().alert(`✅ 已清除全部 ${count} 個自動排程。`);
}

function removeMonthlyTrigger() {
  let count = 0;
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendMonthlyReport') { ScriptApp.deleteTrigger(trigger); count++; }
  });
  SpreadsheetApp.getUi().alert(`已刪除 ${count} 個月報觸發器。`);
}

// ============================================
// 📊 [v7.5] 每週支出摘要推播
// ============================================
function setupWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendWeeklyReport') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendWeeklyReport').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  SpreadsheetApp.getUi().alert('✅ 已設定：每週一早上 8:00 自動推播上週支出摘要！');
}

function sendWeeklyReport() {
  const now = new Date();
  const dayOfWeek = now.getDay() || 7; // 1=Mon … 7=Sun

  // 計算上週一 00:00 ~ 上週日 23:59
  const lastMonday = new Date(now);
  lastMonday.setDate(now.getDate() - dayOfWeek - 6);
  lastMonday.setHours(0, 0, 0, 0);
  const lastSunday = new Date(lastMonday);
  lastSunday.setDate(lastMonday.getDate() + 6);
  lastSunday.setHours(23, 59, 59, 999);

  const fmt = d => Utilities.formatDate(d, CONFIG.TIMEZONE, 'MM/dd');
  const weekLabel = `${fmt(lastMonday)} ~ ${fmt(lastSunday)}`;

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const stats = aggregateDateRangeData(data, lastMonday, lastSunday);

  const lineSep = '━━━━━━━━━━';
  let message;

  if (stats.count === 0) {
    message = `📊 上週（${weekLabel}）支出摘要\n${lineSep}\n✅ 上週沒有任何消費紀錄，太厲害了！`;
  } else {
    const topCats = Object.entries(stats.categoryStats).sort((a, b) => b[1] - a[1]).slice(0, 3);
    const incomeStr = stats.totalIncome > 0 ? `\n📥 收入：${formatCurrency(stats.totalIncome)}` : '';

    // 本月預算使用率
    const { month, year } = getCurrentMonthYear();
    const monthStats = aggregateMonthlyData(data, month, year);
    const budgetPct = CONFIG.MONTHLY_BUDGET > 0 ? (monthStats.totalExpense / CONFIG.MONTHLY_BUDGET * 100).toFixed(1) : null;
    const budgetStr = budgetPct !== null ? `\n${lineSep}\n💰 本月預算已用 ${budgetPct}%（${formatCurrency(monthStats.totalExpense)} / ${formatCurrency(CONFIG.MONTHLY_BUDGET)}）` : '';

    message = `📊 上週（${weekLabel}）支出摘要\n${lineSep}${incomeStr}\n💸 支出：${formatCurrency(stats.totalExpense)}\n📝 共 ${stats.count} 筆\n${lineSep}`;

    if (topCats.length > 0) {
      message += '\n🏆 支出分類 Top 3：';
      topCats.forEach(([cat, amt], i) => {
        const pct = stats.totalExpense > 0 ? (amt / stats.totalExpense * 100).toFixed(1) : '0.0';
        message += `\n${i + 1}. ${cat}：${formatCurrency(amt)} (${pct}%)`;
      });
    }

    const specialProjects = Object.entries(stats.projectStats)
      .filter(([name, amt]) => name !== CONFIG.DEFAULT_PROJECT && amt > 0)
      .sort((a, b) => b[1] - a[1]);
    if (specialProjects.length > 0) {
      message += `\n${lineSep}\n📌 特殊專案：`;
      specialProjects.forEach(([name, amt]) => { message += `\n• ${name}：${formatCurrency(amt)}`; });
    }

    message += budgetStr;

    // [v7.8] 統計上週未標卡別的信用卡消費筆數
    let untaggedCardCount = 0;
    for (let i = 1; i < data.length; i++) {
      const rowDate = new Date(data[i][0]);
      if (isNaN(rowDate) || rowDate < lastMonday || rowDate > lastSunday) continue;
      const payment = String(data[i][4] || '').toLowerCase();
      const cardId = String(data[i][9] || '').trim();
      if ((payment.includes('信用卡') || payment.includes('credit')) && !cardId) {
        untaggedCardCount++;
      }
    }
    if (untaggedCardCount > 0) {
      message += `\n${lineSep}\n💳 提醒：本週有 ${untaggedCardCount} 筆信用卡消費未標記卡別`;
    }
  }

  const chartUrl = stats.count > 0 ? generatePieChartUrl(stats.categoryStats, `上週（${weekLabel}）支出`) : null;
  Object.keys(CONFIG.USERS).forEach(userId => pushLine(userId, message, chartUrl));
  Logger.log(`[週報推播] 上週 ${weekLabel} 週報已推播`);
}

// ============================================
// 🎊 [v7.8] 年度回顧推播
// ============================================
function aggregateAnnualData(data, year) {
  const monthlyTotals = Array.from({ length: 12 }, (_, i) => ({ month: i + 1, total: 0, income: 0 }));
  const categoryStats = {}, personStats = {};
  const allExpenses = [];

  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][0]);
    if (isNaN(rowDate) || rowDate.getFullYear() !== year) continue;
    const m = rowDate.getMonth(); // 0-indexed
    const amount = parseInt(data[i][2]) || 0;
    const category = String(data[i][5] || '');
    const recorder = String(data[i][3] || '').trim();
    const project = String(data[i][6] || CONFIG.DEFAULT_PROJECT);

    if (category === '工資') {
      monthlyTotals[m].income += amount;
    } else if (amount > 0) {
      if (project === CONFIG.DEFAULT_PROJECT) {
        monthlyTotals[m].total += amount;
        categoryStats[category] = (categoryStats[category] || 0) + amount;
      }
      if (recorder) personStats[recorder] = (personStats[recorder] || 0) + amount;
      allExpenses.push({ date: rowDate, item: String(data[i][1] || ''), amount, category });
    }
  }

  allExpenses.sort((a, b) => b.amount - a.amount);
  return { monthlyTotals, categoryStats, personStats, allExpenses };
}

function sendAnnualReport(forceYear = null) {
  const now = new Date();
  const year = forceYear || now.getFullYear();
  // 觸發器呼叫時，只在 12 月執行（避免每月 31 日都執行）
  if (!forceYear && now.getMonth() !== 11) {
    Logger.log('[年報推播] 非 12 月，跳過年報推播');
    return;
  }

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const { monthlyTotals, categoryStats, personStats, allExpenses } = aggregateAnnualData(data, year);

  const yearTotal = monthlyTotals.reduce((s, m) => s + m.total, 0);
  const yearIncome = monthlyTotals.reduce((s, m) => s + m.income, 0);
  const top10 = allExpenses.slice(0, 10);
  const topCats = Object.entries(categoryStats).sort((a, b) => b[1] - a[1]).slice(0, 5);

  const lineSep = '━━━━━━━━━━';
  let msg = `🎊 ${year} 年度消費回顧\n${lineSep}\n`;
  msg += `💸 全年總支出：${formatCurrency(yearTotal)}\n`;
  if (yearIncome > 0) {
    msg += `📥 全年收入：${formatCurrency(yearIncome)}\n`;
    const netSaving = yearIncome - yearTotal;
    msg += `💰 全年淨儲蓄：${netSaving >= 0 ? '' : '-'}${formatCurrency(Math.abs(netSaving))}\n`;
  }
  msg += `📅 月均支出：${formatCurrency(Math.round(yearTotal / 12))}`;

  msg += `\n${lineSep}\n📅 各月份支出：`;
  monthlyTotals.forEach(m => {
    if (m.total > 0) msg += `\n${m.month}月：${formatCurrency(m.total)}`;
  });

  if (topCats.length > 0) {
    msg += `\n${lineSep}\n📊 全年分類 Top 5：`;
    topCats.forEach(([cat, amt], i) => { msg += `\n${i + 1}. ${cat}：${formatCurrency(amt)}`; });
  }

  if (Object.keys(personStats).length >= 2) {
    msg += `\n${lineSep}\n👫 兩人全年支出：`;
    Object.entries(personStats).sort((a, b) => b[1] - a[1]).forEach(([person, amt]) => {
      msg += `\n• ${person}：${formatCurrency(amt)}`;
    });
  }

  if (top10.length > 0) {
    msg += `\n${lineSep}\n🏆 Top 10 單筆消費：`;
    top10.forEach((e, i) => {
      const dateStr = Utilities.formatDate(e.date, CONFIG.TIMEZONE, 'MM/dd');
      const itemShort = e.item.length > 12 ? e.item.substring(0, 12) + '…' : e.item;
      msg += `\n${i + 1}. ${dateStr} ${itemShort}  NT$${e.amount.toLocaleString()}`;
    });
  }

  const chartUrl = Object.keys(categoryStats).length > 0
    ? generatePieChartUrl(categoryStats, `${year}年全年支出分布`)
    : null;

  Object.keys(CONFIG.USERS).forEach(userId => pushLine(userId, msg, chartUrl));
  Logger.log(`[年報推播] ${year} 年度回顧已推播`);
}

function setupAnnualTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendAnnualReport') ScriptApp.deleteTrigger(t);
  });
  // 每月 31 日 23:00 執行，函數內部判斷 12 月才實際推播
  ScriptApp.newTrigger('sendAnnualReport').timeBased().onMonthDay(31).atHour(23).create();
  SpreadsheetApp.getUi().alert('✅ 已設定：每年 12/31 23:00 自動推播年度回顧！');
}

function testSendAnnualReport() {
  sendAnnualReport(new Date().getFullYear()); // 強制以當年資料執行，不受月份限制
}
