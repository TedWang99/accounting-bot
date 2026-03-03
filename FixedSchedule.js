// ============================================
// FixedSchedule - LINE 家庭記帳機器人 v7.7
// ============================================

function initFixedScheduleSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('固定收支設定');
  if (!sheet) {
    sheet = ss.insertSheet('固定收支設定');
    sheet.appendRow(['每月日期', '收/支', '項目', '金額', '分類', '專案', '支付方式', '📌 結束月份(YYYY-MM)']);
    sheet.getRange('A1:H1').setBackground('#d9ead3').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.appendRow([5, '收入', '薪水', 80000, '工資', '🏠 一般開銷', '銀行轉帳', '']);
    sheet.appendRow([10, '支出', '房貸', 25000, '居家', '🏠 一般開銷', '銀行扣款', '']);
    sheet.appendRow([15, '支出', '老婆電話費', 999, '電話網路', '🏠 一般開銷', '信用卡', '']);
    SpreadsheetApp.getUi().alert('✅ 已建立「固定收支設定」工作表！\n(支援設定結束月份，如 2026-10)');
  } else {
    SpreadsheetApp.getUi().alert('⚠️ 「固定收支設定」工作表已存在！');
  }
}

function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'processDailySchedules') { ScriptApp.deleteTrigger(t); }
  });
  ScriptApp.newTrigger('processDailySchedules').timeBased().everyDays(1).atHour(8).create();
  SpreadsheetApp.getUi().alert('✅ 已設定：每天早上 8:00 固定收支自動入帳');
}

function setupWeekendInvoiceTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyAutoImportInvoices') { ScriptApp.deleteTrigger(t); }
  });
  [ScriptApp.WeekDay.FRIDAY, ScriptApp.WeekDay.SATURDAY, ScriptApp.WeekDay.SUNDAY].forEach(day => {
    ScriptApp.newTrigger('dailyAutoImportInvoices').timeBased().onWeekDay(day).atHour(22).create();
  });
  SpreadsheetApp.getUi().alert('✅ 已設定：每週五、六、日 晚上 10:00 自動掃描雲端資料夾並匯入發票！');
}

function processDailySchedules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fixedSheet = ss.getSheetByName('固定收支設定');
  if (!fixedSheet) return;
  const data = fixedSheet.getDataRange().getValues();
  if (data.length <= 1) return;
  const now = new Date();
  const today = now.getDate();
  const currentMonth = now.getMonth() + 1;
  let addedCount = 0;
  const lineSep = '━━━━━━━━━━';
  let summary = `🤖 【系統自動入帳通知】\n今天是 ${currentMonth} 月 ${today} 日\n${lineSep}`;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const [day, type, item, amount, category, project, payment, endMonth] = row;
    if (parseInt(day) !== today || !item || !amount || isNaN(parseInt(amount))) continue;

    // [v7.1] 檢查期限
    if (endMonth) {
      const eMonthStr = String(endMonth).trim();
      if (eMonthStr.match(/^\d{4}-\d{2}$/)) {
        const [eYear, eMon] = eMonthStr.split('-').map(Number);
        if (now.getFullYear() > eYear || (now.getFullYear() === eYear && currentMonth > eMon)) {
          // 已過期，不記錄
          continue;
        }
      }
    }

    writeToSheet({
      date: now,
      item, amount: parseInt(amount), userName: '系統自動',
      payment: payment || (type === '收入' ? '銀行轉帳' : CONFIG.DEFAULT_PAYMENT),
      category: category || (type === '收入' ? '工資' : CONFIG.DEFAULT_CATEGORY),
      project: project || CONFIG.DEFAULT_PROJECT,
      invoiceNum: '', source: '自動排程'
    });

    // [v7.1] 使用 buildSuccessMessage 來取得防爆鎖警告 (拔除第一行的記帳成功字樣)
    const rawMsg = buildSuccessMessage({
      userName: '系統', item, category: category || CONFIG.DEFAULT_CATEGORY,
      project: project || CONFIG.DEFAULT_PROJECT, amount: parseInt(amount),
      payment: payment || CONFIG.DEFAULT_PAYMENT, isTravelMode: false,
      originalAmountStr: null, expenseDate: now
    });
    // 擷取 `buildSuccessMessage` 裡面可能包含的 🔥、🚨、⚠️ 或 📊 等預算警告行
    const alertMatch = rawMsg.match(/(\n🔥.*|\n🚨.*|\n⚠️.*|\n📊.*|\n💡.*)/g);
    const alertText = alertMatch ? alertMatch.join('') : '';

    summary += `\n${type === '收入' ? '💰' : '💸'} [${type}] ${item}：${formatCurrency(parseInt(amount))}${alertText}`;

    // 如果剛好是最後一個月，加上提醒
    if (endMonth) {
      const eMonthStr = String(endMonth).trim();
      if (/^\d{4}-\d{2}$/.test(eMonthStr)) {
        const [eYear, eMon] = eMonthStr.split('-').map(Number);
        if (now.getFullYear() === eYear && currentMonth === eMon) {
          summary += `\n  ⚠️ (注意：此筆項目本月為最後一期)`;
        }
      }
    }

    addedCount++;
  }

  if (addedCount > 0) {
    summary += `\n${lineSep}\n共自動入帳 ${addedCount} 筆！`;
    Object.keys(CONFIG.USERS).forEach(userId => pushLine(userId, summary));
  }
}

function dailyAutoImportInvoices() {
  Logger.log('[自動排程] 開始掃描雲端載具資料夾...');
  processInvoices(true);
  Logger.log('[自動排程] 掃描完成');
}

function handleUpdateFixedSchedule(userMessage) {
  try {
    const parsedData = parseUpdateFixedScheduleWithGemini(userMessage);
    if (!parsedData || !parsedData.target_keyword) {
      return { status: 'error', message: '⚠️ 無法解析修改指令，請說明要修改「哪個項目」以及「新的金額或日期」！\n範例：幫我把房貸改成 26000 元' };
    }
    const keyword = parsedData.target_keyword;
    const fixedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('固定收支設定');
    if (!fixedSheet) return { status: 'error', message: '⚠️ 找不到固定收支設定表！' };

    const data = fixedSheet.getDataRange().getValues();
    let targetRowIndex = -1;
    let oldData = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).includes(keyword)) {
        targetRowIndex = i + 1;
        oldData = data[i];
        break;
      }
    }
    if (targetRowIndex === -1) { return { status: 'error', message: `⚠️ 找不到名稱包含「${keyword}」的項目！請確認名稱是否正確。` }; }

    let updateMsg = `✅ 修改成功！(🤖 AI 智慧解析)\n📝 項目：${oldData[2]}`;
    let isUpdated = false;

    if (parsedData.new_amount !== null && parsedData.new_amount !== undefined) {
      fixedSheet.getRange(targetRowIndex, 4).setValue(parsedData.new_amount);
      updateMsg += `\n💰 金額：${formatCurrency(oldData[3])} ➔ ${formatCurrency(parsedData.new_amount)}`;
      isUpdated = true;
    }
    if (parsedData.new_day !== null && parsedData.new_day !== undefined) {
      fixedSheet.getRange(targetRowIndex, 1).setValue(parsedData.new_day);
      updateMsg += `\n🗓️ 日期：每月 ${oldData[0]} 號 ➔ 每月 ${parsedData.new_day} 號`;
      isUpdated = true;
    }
    if (parsedData.new_payment !== null && parsedData.new_payment !== undefined) {
      fixedSheet.getRange(targetRowIndex, 7).setValue(parsedData.new_payment);
      updateMsg += `\n💳 支付：${oldData[6]} ➔ ${parsedData.new_payment}`;
      isUpdated = true;
    }
    if (parsedData.new_end_month !== null && parsedData.new_end_month !== undefined) {
      if (parsedData.new_end_month === 'CLEAR') {
        fixedSheet.getRange(targetRowIndex, 8).setValue('');
        updateMsg += `\n⏳ 期限：已移除 (變更為無限期)`;
      } else {
        fixedSheet.getRange(targetRowIndex, 8).setValue(parsedData.new_end_month);
        updateMsg += `\n⏳ 期限：變更為 ${parsedData.new_end_month} 止`;
      }
      isUpdated = true;
    }
    if (!isUpdated) {
      return { status: 'error', message: `⚠️ 找到「${oldData[2]}」，但沒有偵測到要修改的內容。\n範例：把房貸改成 15 號，或是 房貸繳到 2026-05 為止` };
    }
    return { status: 'success', message: updateMsg };
  } catch (error) { logError('handleUpdateFixedSchedule', error); return { status: 'error', message: '❌ 修改失敗：' + error.toString() }; }
}

function handleAddFixedSchedule(userMessage) {
  try {
    const parsedData = parseFixedScheduleWithGemini(userMessage);
    if (!parsedData || !parsedData.day || !parsedData.amount) {
      return { status: 'error', message: '⚠️ 無法解析固定收支設定，請再說清楚一點！\n範例：新增固定 每個月15號扣卡費2000元' };
    }
    const { day, type, item, amount, category, project, payment, end_month } = parsedData;
    const fixedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('固定收支設定');
    if (!fixedSheet) return { status: 'error', message: '⚠️ 找不到固定收支設定表！請先從試算表選單建立。' };

    const validProjects = CONFIG.PROJECTS.map(p => p.name);
    const finalProject = (project && validProjects.includes(project)) ? project : CONFIG.DEFAULT_PROJECT;

    fixedSheet.appendRow([
      day, type, item, amount,
      category || (type === '收入' ? '工資' : CONFIG.DEFAULT_CATEGORY),
      finalProject,
      payment || (type === '收入' ? '銀行轉帳' : '自動扣款'),
      end_month || ''
    ]);
    const endRuleStr = end_month ? `\n⏳ 扣款至 ${end_month} 止` : '';
    const projectStr = finalProject !== CONFIG.DEFAULT_PROJECT ? `\n🚀 專案：${finalProject}` : '';
    return {
      status: 'success',
      message: `✅ 固定收支新增成功！(🤖 AI 智慧解析)\n🗓️ 每月 ${day} 號 [${type}]\n📝 項目：${item}\n📂 分類：${category}${projectStr}\n💳 支付：${payment}\n💰 金額：${formatCurrency(amount)}${endRuleStr}`
    };
  } catch (error) { logError('handleAddFixedSchedule', error); return { status: 'error', message: '❌ 設定失敗：' + error.toString() }; }
}

function handleQueryFixedSchedule(userMessage) {
  try {
    const keyword = userMessage.trim().split(/\s+/)[1] || '';
    const fixedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('固定收支設定');
    if (!fixedSheet) return { status: 'error', message: '⚠️ 找不到固定收支設定表！' };
    const data = fixedSheet.getDataRange().getValues();
    if (data.length <= 1) return { status: 'success', message: '📋 目前沒有設定任何固定收支！' };

    const lineSep = '━━━━━━━━━━';
    let message = keyword ? `🔍 固定收支查詢：${keyword}\n${lineSep}\n` : `📋 所有固定收支設定\n${lineSep}\n`;
    let totalIncome = 0, totalExpense = 0, matchCount = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const [day, type, item, amount, category, project, _payment, endMonth] = row;
      if (!item) continue;
      if (keyword && !(String(item) + String(category) + String(project) + String(type)).includes(keyword)) continue;

      let validStr = '';
      if (endMonth) {
        const eMonthStr = String(endMonth).trim();
        if (eMonthStr.match(/^\d{4}-\d{2}$/)) {
          const [eY, eM] = eMonthStr.split('-').map(Number);
          const now = new Date();
          if (now.getFullYear() > eY || (now.getFullYear() === eY && (now.getMonth() + 1) > eM)) {
            validStr = ' (❌已過期停扣)';
          } else {
            validStr = ` (⏳至 ${eMonthStr} 止)`;
          }
        }
      }

      matchCount++;
      message += `${type === '收入' ? '📥' : '💸'} 每月 ${day} 號 | [${type}] ${item}：${formatCurrency(amount)}${validStr}\n`;
      if (!validStr.includes('已過期')) {
        if (type === '收入') totalIncome += parseInt(amount) || 0;
        else totalExpense += parseInt(amount) || 0;
      }
    }
    if (keyword && matchCount === 0) return { status: 'success', message: `⚠️ 找不到包含「${keyword}」的固定收支！` };
    message += `${lineSep}\n💰 預計總收入：${formatCurrency(totalIncome)}\n💳 預計總支出：${formatCurrency(totalExpense)}`;
    return { status: 'success', message };
  } catch (error) { logError('handleQueryFixedSchedule', error); return { status: 'error', message: '❌ 查詢失敗：' + error.toString() }; }
}

// ============================================
// 💳 [v7.3+] 刷卡分期自動寫入固定收支
// ============================================
/**
 * 將分期付款資訊寫入「固定收支設定」工作表。
 */
function addInstallmentToFixedSchedule({ item, totalAmount, installments, day, category, project, payment, expenseDate }) {
  try {
    const fixedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('固定收支設定');
    if (!fixedSheet) {
      logError('addInstallmentToFixedSchedule', '找不到「固定收支設定」工作表，請先從選單初始化。');
      return null;
    }

    const monthlyAmount = Math.round(totalAmount / installments);

    // 到期月份：購買當月算第 1 期，共 installments 期
    const start = expenseDate instanceof Date ? expenseDate : new Date(expenseDate);
    const endDate = new Date(start.getFullYear(), start.getMonth() + installments - 1, 1);
    const end_month = Utilities.formatDate(endDate, CONFIG.TIMEZONE, 'yyyy-MM');

    // 扣款日上限 28 防月底溢位問題
    const safeDay = Math.min(day || start.getDate(), 28);

    // 品名格式：原品名 + (分X期)
    const installmentItem = `${item}(分${installments}期)`;

    fixedSheet.appendRow([
      safeDay,
      '支出',
      installmentItem,
      monthlyAmount,
      category || CONFIG.DEFAULT_CATEGORY,
      project || CONFIG.DEFAULT_PROJECT,
      payment || '信用卡',
      end_month
    ]);

    Logger.log(`[分期自動入帳] ${installmentItem} × ${installments} 期，每月 ${monthlyAmount} 元，至 ${end_month} 止`);
    return { monthlyAmount, end_month, day: safeDay };
  } catch (error) {
    logError('addInstallmentToFixedSchedule', error);
    return null;
  }
}

function handleDeleteFixedSchedule(userMessage) {
  try {
    const keyword = userMessage.replace(/^(刪除固定|移除固定|取消固定)/, '').trim();
    if (!keyword) {
      return { status: 'error', message: '⚠️ 請提供要刪除的固定收支關鍵字！\n例如：刪除固定 房貸' };
    }

    const fixedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('固定收支設定');
    if (!fixedSheet) return { status: 'error', message: '⚠️ 找不到固定收支設定表！' };

    const data = fixedSheet.getDataRange().getValues();
    let targetRowIndex = -1;
    let targetData = null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).includes(keyword)) {
        targetRowIndex = i + 1; // getValues 是 0-indexed，Sheet 的 Row 是 1-indexed
        targetData = data[i];
        break;
      }
    }

    if (targetRowIndex === -1) {
      return { status: 'error', message: `⚠️ 找不到名稱包含「${keyword}」的固定項目！請輸入「查固定」確認名稱。` };
    }

    fixedSheet.deleteRow(targetRowIndex);
    return {
      status: 'success',
      message: `🗑️ 成功刪除固定項目：\n📝 項目：${targetData[2]}\n💰 金額：${formatCurrency(targetData[3])}\n🗓️ 日期：每月 ${targetData[0]} 號\n(若要恢復，請重新新增)`
    };

  } catch (error) {
    logError('handleDeleteFixedSchedule', error);
    return { status: 'error', message: '❌ 刪除固定項目失敗：' + error.toString() };
  }
}
