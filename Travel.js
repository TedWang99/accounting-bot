// ============================================
// Travel - LINE 家庭記帳機器人 v7.7
// ============================================

function handleTravelStart(userMessage) {
  try {
    const parsedData = parseTravelStartWithGemini(userMessage);
    if (!parsedData || !parsedData.destination) {
      return { status: 'error', message: '⚠️ 無法解析旅遊設定！\n請說明要去哪裡玩、去幾天。例如：「開啟北海道 8 天之旅，預算 8 萬」' };
    }
    const destination = parsedData.destination;
    const days = parsedData.days || 5;
    const budgetRaw = parsedData.budget;
    const budgetCurrency = parsedData.budget_currency || 'TWD';
    let currency = parsedData.currency;
    let exchangeRate = parsedData.exchange_rate;
    let rateSource = '';

    if (currency && !exchangeRate) {
      const rateResult = getExchangeRate(currency);
      if (rateResult) {
        exchangeRate = Math.round(rateResult.rate * 10000) / 10000;
        const rateDate = AppProps.getProperty(`rate_date_${currency}`) || '今日';
        rateSource = `（${rateDate} 自動取得最新匯率）`;
      } else {
        currency = null;
      }
    } else if (currency && exchangeRate) {
      rateSource = '（手動設定）';
    }

    let budgetTwd = null;
    let budgetStr = '';
    if (budgetRaw) {
      if (budgetCurrency !== 'TWD' && exchangeRate) {
        budgetTwd = Math.round(budgetRaw * exchangeRate);
        budgetStr = `\n💰 旅遊預算：${budgetCurrency} ${budgetRaw.toLocaleString()} (約台幣 ${formatCurrency(budgetTwd)})`;
      } else if (budgetCurrency === 'TWD') {
        budgetTwd = budgetRaw;
        budgetStr = `\n💰 旅遊預算：${formatCurrency(budgetTwd)}`;
      } else {
        budgetStr = `\n⚠️ 匯率取得失敗，預算未設定（原始：${budgetCurrency} ${budgetRaw.toLocaleString()}）\n   可手動指定匯率後重新設定`;
      }
    }

    const endDate = new Date();
    endDate.setDate(endDate.getDate() + days - 1);
    const projectName = `✈️ ${destination}`;
    const props = AppProps;
    props.setProperty('travel_project', projectName);
    props.setProperty('travel_end', endDate.toISOString());
    if (budgetTwd) props.setProperty('travel_budget', budgetTwd.toString());
    else props.deleteProperty('travel_budget');

    if (currency && exchangeRate) {
      props.setProperty('travel_currency', currency);
      props.setProperty('travel_rate', exchangeRate.toString());
    } else {
      props.deleteProperty('travel_currency');
      props.deleteProperty('travel_rate');
    }

    const endStr = Utilities.formatDate(endDate, CONFIG.TIMEZONE, 'MM/dd');
    const currencyStr = currency ? `\n💱 幣別：${currency}（1 ${currency} = ${exchangeRate} 台幣）${rateSource}` : '';
    const lineSep = '━━━━━━━━━━';
    return {
      status: 'travel_start',
      message: `✈️ 旅遊模式已順利開啟！(🤖 AI 智慧設定)\n${lineSep}\n📍 目的地：${destination}\n📅 天數：${days} 天\n🗓️ 自動結束：${endStr}${currencyStr}${budgetStr}\n${lineSep}\n接下來的記帳將自動歸入「${projectName}」\n並會自動幫您換算回台幣！\n(若提早回國，請輸入「結束旅遊」)`
    };
  } catch (error) { logError('handleTravelStart', error); return { status: 'error', message: '❌ 設定旅遊模式失敗：' + error.toString() }; }
}

function handleTravelEnd() {
  const props = AppProps;
  const projectName = props.getProperty('travel_project');
  if (!projectName) return { status: 'error', message: '⚠️ 目前沒有進行中的旅遊模式' };

  const projectTotal = calculateProjectTotal(projectName);
  const budget = parseInt(props.getProperty('travel_budget') || '0');
  const lineSep = '━━━━━━━━━━';
  let summaryStr = `✨ 本趟【${projectName.replace('✈️ ', '')}】總結算：${formatCurrency(projectTotal)}`;

  if (budget > 0) {
    const remaining = budget - projectTotal;
    if (remaining < 0) {
      summaryStr += `\n🚨 哎呀！預算超支了 ${formatCurrency(Math.abs(remaining))}！`;
    } else {
      summaryStr += `\n📊 預算達成率 ${(projectTotal / budget * 100).toFixed(1)}%\n🎉 替荷包省下了 ${formatCurrency(remaining)}！`;
    }
  }

  const travelCurr = props.getProperty('travel_currency');
  props.deleteProperty('travel_project');
  props.deleteProperty('travel_end');
  props.deleteProperty('travel_budget');
  props.deleteProperty('travel_currency');
  props.deleteProperty('travel_rate');
  if (travelCurr) {
    props.deleteProperty(`rate_cache_${travelCurr}`);
    props.deleteProperty(`rate_cache_time_${travelCurr}`);
    props.deleteProperty(`rate_date_${travelCurr}`); // [v5.5 fix] 原本漏刪，導致舊匯率日期殘留
  }
  return {
    status: 'travel_end',
    message: `🏠 歡迎回家！旅遊模式已結束\n已為您切換回「一般開銷」模式\n${lineSep}\n${summaryStr}\n${lineSep}\n日後可用「查 ${projectName.replace('✈️ ', '')}」隨時回顧這趟旅程花費！`
  };
}

function getActiveTravelProject() {
  try {
    const props = AppProps;
    const projectName = props.getProperty('travel_project');
    const endDateStr = props.getProperty('travel_end');
    if (!projectName || !endDateStr) return null;

    const endDate = new Date(endDateStr);
    endDate.setHours(23, 59, 59);
    if (new Date() > endDate) {
      props.deleteProperty('travel_project');
      props.deleteProperty('travel_end');
      props.deleteProperty('travel_budget');   // [v5.5 fix] 原本漏刪，導致舊預算殘留
      props.deleteProperty('travel_currency');
      props.deleteProperty('travel_rate');
      return null;
    }

    const currency = props.getProperty('travel_currency') || null;
    const rate = props.getProperty('travel_rate') ? parseFloat(props.getProperty('travel_rate')) : null;
    return { projectName, currency, rate };
  } catch (error) { return null; }
}

function getTravelStatus() {
  const props = AppProps;
  const projectName = props.getProperty('travel_project');
  const endDateStr = props.getProperty('travel_end');

  if (!projectName) return '🏠 目前沒有進行中的旅遊\n\n開啟旅遊模式：幫我開啟去日本 5 天';

  const endDate = new Date(endDateStr);
  const daysLeft = Math.ceil((endDate - new Date()) / (1000 * 60 * 60 * 24));
  const endStr = Utilities.formatDate(endDate, CONFIG.TIMEZONE, 'MM/dd');
  const projectTotal = calculateProjectTotal(projectName);
  const budget = parseInt(props.getProperty('travel_budget') || '0');

  let budgetInfo = '';
  if (budget > 0) {
    const remaining = budget - projectTotal;
    budgetInfo = remaining >= 0
      ? `\n💰 預算：${formatCurrency(budget)}\n📊 已用 ${(projectTotal / budget * 100).toFixed(1)}%，剩餘 ${formatCurrency(remaining)}`
      : `\n💰 預算：${formatCurrency(budget)}\n🚨 已超支 ${formatCurrency(Math.abs(remaining))}`;
  }

  const currency = props.getProperty('travel_currency');
  const rate = props.getProperty('travel_rate');
  let currencyLine = '';
  if (currency && rate) {
    const rateDate = props.getProperty(`rate_date_${currency}`) || '';
    const rateDateStr = rateDate ? `，${rateDate} 更新` : '';
    currencyLine = `\n💱 幣別：${currency}（1 ${currency} = ${rate} 台幣${rateDateStr}）`;
  }

  const lineSep = '━━━━━━━━━━';
  return `✈️ 旅遊模式進行中！\n${lineSep}\n📍 目的地：${projectName}\n🗓️ 結束日期：${endStr}\n⏳ 剩餘：${daysLeft} 天${currencyLine}\n✨ 目前累計：${formatCurrency(projectTotal)}${budgetInfo}\n${lineSep}\n提早回國請傳「結束旅遊」`;
}
