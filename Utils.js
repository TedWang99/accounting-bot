// ============================================
// Utils - LINE 家庭記帳機器人 v7.8
// ============================================

function parseLineEvent(e) { try { return JSON.parse(e.postData.contents).events?.[0]; } catch (error) { logError('parseLineEvent', error); return null; } }
function parseLineEvents(e) { try { return JSON.parse(e.postData.contents).events || []; } catch (error) { logError('parseLineEvents', error); return []; } }
function isTextMessage(event) { return event.type === 'message' && event.message.type === 'text'; }
function extractEventData(event) {
  return {
    replyToken: event.replyToken,
    userMessage: event.message?.text?.trim() ?? '',
    userId: event.source.userId
  };
}
function getUserName(userId) { return CONFIG.USERS[userId] || `訪客 (${userId.substring(0, 8)}...)`; }
function getSheet() { return SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheets()[0]; }
function isValidAmount(amount) { return !isNaN(amount) && parseFloat(amount) !== 0; }
function parseAmount(rawAmount) { return Math.round(parseFloat(rawAmount)); }
function formatCurrency(amount) { return `${amount.toLocaleString('zh-TW')} 元`; }
function getCurrentMonthYear() { const now = new Date(); return { month: now.getMonth() + 1, year: now.getFullYear() }; }
function isCurrentMonth(date, month, year) { return date.getMonth() + 1 === month && date.getFullYear() === year; }
function createResponse(status, details = '') { return ContentService.createTextOutput(JSON.stringify({ status, details, timestamp: new Date() })).setMimeType(ContentService.MimeType.JSON); }

function safeJSONParse(text) {
  try {
    return JSON.parse(text);
  } catch (e) {
    logError('safeJSONParse', `JSON 解析失敗。截斷內容：${text.substring(0, 100)}`);
    return { error: { message: '系統回傳非預期的資料格式' } }; // 兼容現有 if(result.error) 邏輯
  }
}

// 加入重試機制的 Fetch Wrapper (用以解決 API 偶發性 500, 503, 429 等問題)
function fetchWithRetry(url, options = {}, maxRetries = 3) {
  // 強制加入 muteHttpExceptions，讓 4xx/5xx 回傳 response 物件而非拋出 Exception
  // 如此才能正確執行 getResponseCode() 判斷，而不是全靠 catch 兜底
  const fetchOptions = { ...options, muteHttpExceptions: true };
  let attempt = 0;
  let lastResponse = null;
  while (attempt < maxRetries) {
    try {
      lastResponse = UrlFetchApp.fetch(url, fetchOptions);
      const code = lastResponse.getResponseCode();

      // 200~299 為成功，若是 400 系列 (如驗證錯誤) 通常重試也沒用，直接回傳讓外層處理
      if (code < 400 || (code >= 400 && code < 500 && code !== 429)) {
        return lastResponse;
      }

      // 遇到 429 (Too Many Requests), 500 等伺服器錯誤才重試
      Logger.log(`⚠️ API 請求失敗 (狀態碼 ${code})，準備第 ${attempt + 1} 次重試...`);
    } catch (e) {
      Logger.log(`⚠️ API 請求發生例外狀況：${e.toString()}，準備第 ${attempt + 1} 次重試...`);
      // 若是最後一次嘗試仍失敗，將錯誤拋出
      if (attempt === maxRetries - 1) throw e;
    }

    attempt++;
    if (attempt < maxRetries) {
      // 指數型 Backoff 等待： 1秒, 2秒, 4秒...
      Utilities.sleep(Math.pow(2, attempt - 1) * 1000);
    }
  }
  return lastResponse; // 回傳最後一次的結果，不再額外發第 4 次請求
}

function calculateProjectTotal(targetProject, month = null, year = null) {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    let total = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][6] !== targetProject) continue;
      if (month !== null && year !== null) {
        const rowDate = new Date(data[i][0]);
        if (rowDate.getMonth() + 1 !== month || rowDate.getFullYear() !== year) continue;
      }
      total += parseInt(data[i][2]) || 0;
    }
    return total;
  } catch (error) { logError('calculateProjectTotal', error); return 0; }
}

function getCurrencySymbol(currency) {
  const symbols = { JPY: '¥', USD: '$', EUR: '€', KRW: '₩', HKD: 'HK$', SGD: 'S$', GBP: '£', THB: '฿', AUD: 'A$' };
  return symbols[currency] || (currency + ' ');
}

// ============================================
// 💱 匯率自動取得
// ============================================
function getExchangeRate(currency) {
  try {
    const props = AppProps;
    const cacheKey = `rate_cache_${currency}`;
    const cacheTimeKey = `rate_cache_time_${currency}`;
    const cachedRate = props.getProperty(cacheKey);
    const cachedTime = props.getProperty(cacheTimeKey);
    if (cachedRate && cachedTime && (Date.now() - parseInt(cachedTime)) < 24 * 60 * 60 * 1000) {
      return { rate: parseFloat(cachedRate), fromCache: true };
    }
    const url = `https://open.er-api.com/v6/latest/${currency}`;
    const response = fetchWithRetry(url, { muteHttpExceptions: true });
    const data = safeJSONParse(response.getContentText());
    if (data.result !== 'success' || !data.rates) {
      logError('getExchangeRate', `API 回傳失敗：${data['error-type'] || '格式異常'}`);
      return cachedRate ? { rate: parseFloat(cachedRate), fromCache: true } : null;
    }
    const rate = data.rates?.['TWD'];
    if (!rate) return null;
    props.setProperty(cacheKey, rate.toString());
    props.setProperty(cacheTimeKey, Date.now().toString());
    props.setProperty(`rate_date_${currency}`, Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'MM/dd'));
    return { rate, fromCache: false };
  } catch (error) { logError('getExchangeRate', error); return null; }
}

function logError(functionName, error) { Logger.log(`[錯誤] ${functionName}: ${error.toString()}`); }

// [v7.7] 動態設定卡別下拉選單（從 CARDS_MAP 取得選項）
function setupCardDropdown() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheets()[0];
  const cardsMap = CONFIG.CARDS_MAP;
  if (!cardsMap || Object.keys(cardsMap).length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ CARDS_MAP 是空的，請先在指令碼屬性中設定卡別資料');
    return;
  }
  const cardKeys = Object.keys(cardsMap);
  const lastRow = Math.max(sheet.getLastRow(), 100);
  const range = sheet.getRange(2, 10, lastRow, 1); // J2 開始（第10欄 = 卡別）
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(cardKeys, true)  // true = 顯示下拉箭頭
    .setAllowInvalid(true)               // 允許手動輸入（彈性）
    .build();
  range.setDataValidation(rule);
  SpreadsheetApp.getUi().alert(`✅ 已為 J 欄設定卡別下拉選單\n共 ${cardKeys.length} 個選項：${cardKeys.join('、')}`);
}

// [v8.0] 每日自動更新用（無 UI alert，可由 Trigger 呼叫）
function setupCardDropdownSilent() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheets()[0];
    const cardsMap = CONFIG.CARDS_MAP;
    if (!cardsMap || Object.keys(cardsMap).length === 0) return;
    const cardKeys = Object.keys(cardsMap);
    const lastRow = Math.max(sheet.getLastRow(), 100);
    const range = sheet.getRange(2, 10, lastRow, 1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(cardKeys, true)
      .setAllowInvalid(true)
      .build();
    range.setDataValidation(rule);
    Logger.log(`[setupCardDropdownSilent] 已更新 J 欄下拉選單，共 ${cardKeys.length} 個選項`);
  } catch (e) {
    logError('setupCardDropdownSilent', e);
  }
}

// [v8.0] 設定每日 05:00 自動更新 J 欄卡別下拉選單的 Trigger
function setupCardDropdownTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'setupCardDropdownSilent') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('setupCardDropdownSilent').timeBased().everyDays(1).atHour(5).create();
  SpreadsheetApp.getUi().alert('✅ 已設定：每天早上 05:00 自動更新 J 欄卡別下拉選單');
}

function replyLine(replyToken, message, imageUrl = null, flexContents = null, extraText = null) {
  try {
    const messagesArray = [];

    // 如果有 flexContents，優先加入 Flex Message 作為主訊息
    if (flexContents) {
      messagesArray.push({
        type: 'flex',
        altText: typeof message === 'string' ? message.substring(0, 300) : '系統通知',
        contents: flexContents
      });
    } else {
      // 保留原本純文字防過長邏輯
      const safeMessage = typeof message === 'string' && message.length > 5000
        ? message.substring(0, 4990) + '\n...(訊息過長已截斷)'
        : message || '';
      if (safeMessage) messagesArray.push({ type: 'text', text: safeMessage });
    }

    // [v7.9] 額外文字訊息（如重複記帳警示），在 Flex 之後獨立發送
    if (extraText) {
      messagesArray.push({ type: 'text', text: extraText });
    }

    if (imageUrl) {
      messagesArray.push({ type: 'image', originalContentUrl: imageUrl, previewImageUrl: imageUrl });
    }

    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      headers: { 'Content-Type': 'application/json; charset=UTF-8', 'Authorization': 'Bearer ' + CONFIG.LINE_ACCESS_TOKEN },
      method: 'post',
      payload: JSON.stringify({ replyToken, messages: messagesArray }),
      muteHttpExceptions: true
    });
  } catch (error) { logError('replyLine', error); }
}

function pushLine(userId, message, imageUrl = null, flexContents = null) {
  try {
    const messagesArray = [];

    if (flexContents) {
      messagesArray.push({
        type: 'flex',
        altText: typeof message === 'string' ? message.substring(0, 300) : '系統通知',
        contents: flexContents
      });
    } else {
      const safeMessage = typeof message === 'string' && message.length > 5000
        ? message.substring(0, 4990) + '\n...(訊息過長已截斷)'
        : message || '';
      if (safeMessage) messagesArray.push({ type: 'text', text: safeMessage });
    }

    if (imageUrl) {
      messagesArray.push({ type: 'image', originalContentUrl: imageUrl, previewImageUrl: imageUrl });
    }

    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      headers: { 'Content-Type': 'application/json; charset=UTF-8', 'Authorization': 'Bearer ' + CONFIG.LINE_ACCESS_TOKEN },
      method: 'post', payload: JSON.stringify({ to: userId, messages: messagesArray }), muteHttpExceptions: true
    });
  } catch (error) { logError('pushLine', error); }
}

function getCategory(item) {
  const categoryMap = {
    '飲食': ['餐', '飯', '麵', '吃', '便當', '火鍋', '燒烤', '壽司', '早餐', '午餐', '晚餐', '麥當勞', '肯德基', '摩斯', '全聯', '好市多'],
    '飲料': ['飲料', '珍奶', '手搖', '咖啡', '星巴克', '可樂', '果汁'],
    '點心': ['點心', '蛋糕', '餅乾', '零食', '甜點', '冰淇淋'],
    '酒': ['酒', '啤酒', '紅酒', '清酒', '威士忌'],
    '汽車機車': ['加油', '停車', '保養', '輪胎', 'etag', '過路費', '洗車'],
    '交通': ['捷運', '公車', '計程車', 'uber', '高鐵', '台鐵', '悠遊卡', '機票', 'ubike'],
    '水電瓦斯': ['電費', '水費', '瓦斯'],
    '電話網路': ['手機費', '網路費', '電信', '第四台'],
    '居家': ['家具', '裝潢', '收納', '清潔用品', '衛生紙', '洗衣精'],
    '家電': ['家電', '冰箱', '洗衣機', '冷氣', '電視', '電鍋', '微波爐', '烤箱', '除濕機', '空氣清淨機', '吸塵器', '掃地機', '熱水器', '洗碗機', 'electrical', 'appliances'],
    '日常用品': ['日用品', '生活用品', '雜貨', '蝦皮', 'momo', 'pchome', '文具'],
    '老公服飾': ['男裝', '男鞋', '西裝', '老公'],
    '老婆服飾': ['女裝', '女鞋', '裙', '老婆', 'uniqlo', 'zara', '包包', '洋裝', '內衣'],
    '醫療保健': ['診所', '掛號', '醫院', '牙', '藥局', '看病', '健檢'],
    '保健食品': ['保健食品', '維他命', '益生菌', '魚油'],
    '保養品': ['保養品', '面膜', '乳液', '防曬', '精華液'],
    '美容美髮': ['剪髮', '燙髮', '染髮', '美甲', '美容'],
    '娛樂': ['電影', 'netflix', 'spotify', '遊戲', 'steam', 'ktv', '展覽', '演唱會'],
    '運動': ['健身', '游泳', '運動', '球場', '瑜珈'],
    '3c數位': ['手機', '電腦', '平板', '耳機', '充電器', '3c'],
    '學習深造': ['課程', '書', '學費', '補習'],
    '保險': ['保險'],
    '稅金': ['稅', '稅金', '牌照稅', '燃料稅'],
    '生寶寶': ['取卵', '卵泡', '婦產科', '備孕', '試管', '葉酸'],
    '捐款': ['捐款', '捐贈', '公益'],
    '折扣回饋': ['折扣', '優惠', '回饋金', '回饋', '點數兌換', '點數', '紅利', '折抵', '現金回饋', 'cashback']
  };
  const lowerItem = item.toLowerCase();
  for (const [category, keywords] of Object.entries(categoryMap)) {
    if (keywords.some(k => lowerItem.includes(k.toLowerCase()))) return category;
  }
  return CONFIG.DEFAULT_CATEGORY;
}

function getProject(item) {
  const lowerItem = item.toLowerCase();
  const projectRules = [
    { name: '👶 生寶寶／備孕', keywords: ['取卵', '卵泡', '婦產科', '備孕', '試管', '葉酸', '產檢'] },
    { name: '🎮 胖嘟嘟玩具', keywords: ['ted玩具', '老公玩具', 'switch', 'steam deck', '模型', 'f1', 'rog', 'macbook', 'mac ', 'apple', 'iphone', 'ipad', 'm4', 'm3', 'm2'] },
    { name: '🪆 可愛嘟嘟玩具', keywords: ['可愛嘟嘟', '老婆玩具', '胖丁', '吉伊卡哇', '扭蛋'] }
  ];
  for (const project of projectRules) {
    if (project.keywords.some(k => lowerItem.includes(k))) return project.name;
  }
  return CONFIG.DEFAULT_PROJECT;
}

function getHelpMessage() {
  const lineSep = '━━━━━━━━━━';
  return `🤖 家庭記帳小幫手 v7.8\n${lineSep}\n\n🌸 記帳（直接說就好！）\n  今天去全聯買菜 500\n  吃火鍋 800 刷卡\n  📸 直接傳收據/刷卡明細照片\n  多行訊息 = 批次記多筆\n\n💳 卡別記錄\n  📸 傳刷卡明細照片（自動識別末四碼）\n  這筆用國泰卡 → 補記上一筆的卡別\n  上一筆用老婆永豐悠遊卡\n\n📊 查詢這個月\n  本月 / 查詢　→ 本月總花費\n  月報 / 報表　→ 圓餅圖＋AI 建議＋兩人對比\n  分類 / 統計　→ 各類明細\n  查預算　　　→ 子預算進度條\n  查分類預算　→ 各分類花費上限狀態\n  上一筆　　　→ 最後一筆記錄（含卡別）\n\n🔍 查特定花費\n  查 [關鍵字]　→ 查 北海道 / 查 娃娃\n  查 上週 / 查 本週\n  查 2025-10　→ 指定年月\n\n✈️ 旅遊模式\n  幫我開啟去日本 5 天，預算 3 萬\n  直接輸入當地數字，自動換算台幣！\n  目前花費多少了 → 查旅遊累計\n  回國囉 → 結束並顯示總結算\n\n😱 記錯了怎麼辦\n  刪除　→ 刪掉上一筆\n  修改上一筆 金額 300\n  修改上一筆 分類 飲食\n  修改上一筆 卡別 國泰末1234\n\n🗓️ 固定收支\n  查固定　→ 查看所有固定扣款\n\n${lineSep}\n💰 月預算：${formatCurrency(CONFIG.MONTHLY_BUDGET)}`;
}
