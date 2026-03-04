// ============================================
// AI_Parse - LINE 家庭記帳機器人 v8.1
// ============================================

// ============================================
// [v8.0] 逐項明細掃描：每個商品獨立記錄
// ============================================
function parseReceiptItemsWithGeminiVision(base64Image, mimeType, knownCurrency, knownRate) {
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    const categoryList = CONFIG.CATEGORIES.join('、');
    const todayStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');

    // 若旅遊模式已設定匯率，直接注入 prompt，不讓 Gemini 自行估算
    const rateInstruction = (knownCurrency && knownRate && knownCurrency !== 'TWD')
      ? `2. 幣別換算：本收據為 ${knownCurrency}，請使用匯率 1 ${knownCurrency} = ${knownRate} TWD 換算成台幣（已由旅遊模式取得當日匯率，不要自行估算）。exchange_rate 填 ${knownRate}。`
      : `2. 幣別換算：若為外幣（如 JPY），請估算当前匯率換算成台幣。若為 TWD 則 exchange_rate 為 1。`;

    const prompt = `你是專業的家庭記帳助理。請從這張收據圖片中，識別出「每一條商品明細」，並個別記錄為獨立項目。

【核心任務】
這不是一般記帳（不只記總金額），要把每一項商品都列出來。

【規則】
1. 品名翻譯：若為日文或外文，請翻譯成台灣慣用繁體中文。知名品牌（如 DHC、LYSOL、GATSBY）保留英文。
${rateInstruction}
3. 個別分類：每一項商品都要根據內容個別判斷分類，不要統一套用同一個分類。
4. 折扣/COUPON：若有折扣行，price_original 為負數，category 填「折扣回饋」。
5. 小計/合計行：不要記錄，因為它不是獨立商品。
6. 今天是 ${todayStr}，請根據帳單日期填寫 date。

【分類清單】${categoryList}

請回傳以下 JSON 格式（只回傳 JSON，不要 Markdown）：
{
  "store": "店家名稱（繁體中文）",
  "date": "YYYY-MM-DD",
  "currency": "${knownCurrency || 'TWD 或 JPY 等'}",
  "exchange_rate": ${knownRate || 1},
  "payment": "現金或信用卡",
  "card_last4": "1234（選填）",
  "items": [
    {
      "name": "翻譯後品名",
      "price_original": 數字（原幣金額，折扣為負數）,
      "price_twd": 數字（台幣金額），
      "category": "從分類清單擇一"
    }
  ],
  "total_original": 數字（原幣總計）,
  "total_twd": 數字（台幣總計）
}`;

    const payload = {
      contents: [{
        parts: [
          { text: prompt },
          { inline_data: { mime_type: mimeType, data: base64Image } }
        ]
      }],
      generationConfig: { temperature: 0, responseMimeType: 'application/json' }
    };

    const response = fetchWithRetry(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });

    const result = safeJSONParse(response.getContentText());
    if (result.error) throw new Error(result.error.message);
    const cleanText = (result.candidates?.[0]?.content?.parts?.[0]?.text || '{}')
      .replace(/```json|```/g, '').trim();
    return JSON.parse(cleanText);
  } catch (error) {
    logError('parseReceiptItemsWithGeminiVision', error);
    return null;
  }
}


// 多頁帳單：同時送多張圖 → 全部傳給 Gemini 一次判斷成一筆
function handleMultiPageImages(messageIds, userName) {
  try {
    const imagesData = messageIds.map(msgId => {
      const url = `https://api-data.line.me/v2/bot/message/${msgId}/content`;
      const blob = fetchWithRetry(url, {
        headers: { 'Authorization': 'Bearer ' + CONFIG.LINE_ACCESS_TOKEN },
        method: 'get'
      }).getBlob();
      return { mimeType: blob.getContentType() || 'image/jpeg', data: Utilities.base64Encode(blob.getBytes()) };
    });

    const aiParsed = parseMultiPageReceiptWithGeminiVision(imagesData);
    if (aiParsed && aiParsed.amount) {
      return processExpense('', userName, aiParsed);
    } else {
      return { status: 'error', message: `🔍 收到 ${messageIds.length} 張圖片，但無法解析金額，請手動輸入！` };
    }
  } catch (error) {
    logError('handleMultiPageImages', error);
    return { status: 'error', message: '❌ 多頁帳單解析失敗：\n' + error.toString() };
  }
}

function parseMultiPageReceiptWithGeminiVision(imagesData) {
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    const categoryList = CONFIG.CATEGORIES.join('、');
    const projectList = CONFIG.PROJECTS.map(p => `- ${p.name}（${p.desc}）`).join('\n');
    const todayStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');

    const prompt = `你是一個專業的家庭記帳與收據掃描助理。使用者傳來的是「同一張帳單的多頁照片」，請綜合所有頁面的資訊，提取出「整張帳單」的記帳資訊。

    【重要】這是多頁帳單，請將所有頁面視為同一筆消費：
    1. 金額：取所有頁面中「最終總計金額」（通常在最後一頁），不要加總各頁面的小計。
    2. 品名：根據最主要的消費項目或店家名稱摘要。
    3. 翻譯：如為外文帳單，品名與店家名請翻譯成繁體中文。
    4. 今天是 ${todayStr}，請根據帳單上的日期填寫。
    5. 分類清單：${categoryList}
    6. 專案清單：\n${projectList}
    7. 發票號碼：若有統一發票號碼（2英文+8數字）請提取 "invoice_num"。
    8. 卡片資訊：
       - 信用卡：若有末四碼（如「VISA ****1234」）請提取 "card_last4"（4 位數字字串）。
       - 悠遊卡：若出現「悠遊卡」「悠遊付」「EasyCard」並附有卡號數字（如「悠遊卡 2084581057」），請提取末四碼（如 "1057"）。若僅有悠遊卡字樣無卡號，請填 "card_last4": "easycard"。

    請回傳以下 JSON 格式（只回傳 JSON，不要 Markdown）：
    {
      "date": "YYYY-MM-DD",
      "item": "店家 - 主要品名摘要",
      "amount": 數字,
      "payment": "現金或信用卡",
      "category": "從分類清單擇一",
      "project": "從專案清單擇一",
      "card_last4": "1234（選填）",
      "invoice_num": "XC81084758（選填）"
    }`;

    const parts = [
      { text: prompt },
      ...imagesData.map(img => ({ inline_data: { mime_type: img.mimeType, data: img.data } }))
    ];

    const response = fetchWithRetry(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts }],
        generationConfig: { temperature: 0, responseMimeType: 'application/json' }
      }),
      muteHttpExceptions: true
    });

    const result = safeJSONParse(response.getContentText());
    if (result.error) throw new Error(result.error.message);
    const cleanText = (result.candidates?.[0]?.content?.parts?.[0]?.text || '{}').replace(/```json|```/g, '').trim();
    return JSON.parse(cleanText);
  } catch (error) {
    logError('parseMultiPageReceiptWithGeminiVision', error);
    return null;
  }
}

function handleImageMessage(messageId, userName) {
  try {
    const url = `https://api-data.line.me/v2/bot/message/${messageId}/content`;
    const response = fetchWithRetry(url, {
      headers: { 'Authorization': 'Bearer ' + CONFIG.LINE_ACCESS_TOKEN },
      method: 'get'
    });
    const blob = response.getBlob();
    const mimeType = blob.getContentType() || 'image/jpeg';
    const base64Image = Utilities.base64Encode(blob.getBytes());

    const aiParsed = parseReceiptWithGeminiVision(base64Image, mimeType);

    if (aiParsed && aiParsed.amount) {
      return processExpense('', userName, aiParsed);
    } else {
      return { status: 'error', message: '🔍 機器人看不清楚這張收據，請換張照片試試或手動輸入金額！' };
    }
  } catch (error) {
    logError('handleImageMessage', error);
    return { status: 'error', message: '❌ 圖片解析失敗：\n' + error.toString() };
  }
}

function parseReceiptWithGeminiVision(base64Image, mimeType = 'image/jpeg') {
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    const categoryList = CONFIG.CATEGORIES.join('、');
    const projectList = CONFIG.PROJECTS.map(p => `- ${p.name}（${p.desc}）`).join('\n');
    const todayStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');

    // [v5.5] 優化 Prompt，針對商場明細表做特別處理
    const prompt = `你是一個專業的家庭記帳與收據掃描助理。請從這張收據圖片中精準提取資訊。

    【🚨 跨國與翻譯規則 🚨】
    1. 翻譯要求：如果偵測到這是一張「日文或外文收據」，請務必將「商店名稱」與「消費明細/品名」自動翻譯成台灣慣用的繁體中文！(註：知名連鎖品牌如 BIC CAMERA, UNIQLO, Lawson, ENEOS 請保留英文原名即可)。
    2. 金額處理：請只提取「最終總計金額」的純數字。
    3. 日期推算：今天是 ${todayStr}，請根據收據上的資訊填寫正確的消費日期。
    4. 品名摘要：如果收據上有很長的多個品名，請用逗號隔開列出前 2-3 個主要項目即可。
    5. 商場明細特別處理：若看到「交易明細表」或類似字眼，請優先提取上方真正的品牌/專櫃名稱（例如：カルディ/咖樂迪），並將下方列出的品名合併為摘要，忽略商場或百貨公司本身的名字（如啦啦寶都、三井）。
    6. 分類與專案判斷：
       分類清單：${categoryList}
       專案清單：\n${projectList}
    7. [v7.7] 卡片資訊：
       - 信用卡：若出現末四碼（如「VISA ****1234」、「末四碼：1234」、「卡號後四碼 1234」、「卡號:356778******9217」末四碼為 9217），請提取 "card_last4"（4 位數字字串）。
       - 悠遊卡：若出現「悠遊卡」「悠遊付」「EasyCard」字樣，並附有卡號數字（如「悠遊卡 2084581057」），請提取該數字的「末四碼」作為 "card_last4"（如 "1057"）。若有悠遊卡字樣但完全沒有卡號數字，請填 "card_last4": "easycard"。
       - 若無任何卡片資訊，省略此欄位。
    8. [v7.7] 發票號碼：若收據上出現統一發票號碼（格式為 2 個英文字母 + 8 個數字，例如 XC81084758、AB12345678），請提取 "invoice_num"（字串）。若無則省略此欄位。

    請回傳以下 JSON 格式：
    {
      "date": "YYYY-MM-DD",
      "item": "翻譯後的商店名稱 - 翻譯後的品名摘要",
      "amount": 數字,
      "payment": "現金或信用卡(若看不出則預設為現金)",
      "category": "從分類清單擇一",
      "project": "從專案清單擇一",
      "card_last4": "1234 或 easycard（選填，無則省略）",
      "invoice_num": "XC81084758（選填，無則省略）"
    }
    ⛔ 嚴格限制：請只回傳合法的 JSON 物件，絕對不要包含任何 Markdown 標記（如 \`\`\`json）或額外說明文字。`;

    const payload = {
      contents: [{
        parts: [
          { text: prompt },
          { inline_data: { mime_type: mimeType, data: base64Image } }
        ]
      }],
      generationConfig: { temperature: 0, responseMimeType: "application/json" }
    };

    const response = fetchWithRetry(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const result = safeJSONParse(response.getContentText());
    if (result.error) throw new Error(result.error.message);
    const cleanText = (result.candidates?.[0]?.content?.parts?.[0]?.text || '{}').replace(/```json|```/g, '').trim();
    return JSON.parse(cleanText);
  } catch (error) {
    logError('parseReceiptWithGeminiVision', error);
    return null;
  }
}

function parseAndClassifyExpenseWithGemini(text, userName) {
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    const categoryList = CONFIG.CATEGORIES.join('、');
    const projectList = CONFIG.PROJECTS.map(p => `- ${p.name}（${p.desc}）`).join('\n');

    const todayStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const prompt = `你是一個聰明的家庭記帳助理。請從使用者的「日常對話」中，擷取記帳資訊並自動分類。
    【重要系統資訊】
    1. 今天是 ${todayStr}
    2. 目前說話的使用者是：「${userName}」 (若對話提到"我"、"自己"，就是指他本人)
    【🚨 嚴格分類規則 🚨】
    - 「汽車機車」：任何關於自家汽機車的花費，例如「停車費」、「加油」、「洗車」、「ETC過路費」、「保養」。千萬不可以歸類到交通！
    - 「交通」：僅限大眾運輸工具，例如「捷運」、「公車」、「計程車」、「高鐵」、「機票」、「悠遊卡加值」。
    - 「老公服飾」/「老婆服飾」：若是買衣服鞋包，請依據對話判斷是誰買的並分別歸類。
    - 「家電」：家用電器，如冰箱、洗衣機、冷氣、電視、電鍋、微波爐、除濕機、空氣清淨機、吸塵器、掃地機器人、熱水器等。不可歸類到居家！
    - 「居家」：裝潢、家具、收納、清潔用品等非電器的居家物品。
    - 「折扣回饋」：折扣、優惠、回饋金、點數兌換、紅利折抵等，金額必須為負數。
    - 「收入」：薪水、獎金、年終、進帳、退款、退稅、補貼、利息、租金收入等一切收入。金額必須為負數（代表資金流入）。例如：「收入薪水50000」→ amount: -50000, category: "收入"。
    對話內容：「${text}」
    請提取出以下 8 個欄位，並以 JSON 格式回傳：
    1. "date": 字串，消費日期 (YYYY-MM-DD)。請根據對話推算。若無特別提及時間，請回傳 "${todayStr}"。
    2. "item": 字串，精簡的消費/收入品名。
    3. "amount": 數字，金額（若無數字回傳 null）。若為折扣/回饋/收入，請回傳負數（例如：-50000）。
    4. "payment": 字串，支付方式的「通用類別」，只能填：現金、信用卡、悠遊卡、Line Pay、街口支付、Apple Pay 等通用詞。⚠️ 禁止填入具體卡片品牌名稱（如「國泰Cube」），那應放在欄位 9 的 card 中。預設「${CONFIG.DEFAULT_PAYMENT}」。
    5. "category": 字串，從清單擇一：${categoryList}
    6. "project": 字串，從專案清單擇一：
    ${projectList}
    7. "installments": 數字，若對話中明確提到「分期」、「分X期」、「分X個月」之類的字眼，請回傳分期期數（例如 36），否則請回傳 null。
    8. "installment_day": 數字 (1~28)，每月扣款日。若有提及請填入，否則根據消費日期的日數填入（即 date 欄位的日期數字）。
    9. "card": 字串，若對話中有提到用哪張卡或哪種支付方式（如「國泰Cube」、「玉山卡」、「Line Pay」、「悠遊卡」、「永豐卡」），請回傳該卡片／支付工具的關鍵字；否則回傳 null。
    ⛔ 嚴格限制：請只回傳合法的 JSON 物件，絕對不要包含任何 Markdown 標記或額外文字。`;

    const response = fetchWithRetry(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0 } }),
      muteHttpExceptions: true
    });

    const result = safeJSONParse(response.getContentText());
    if (result.error) throw new Error(result.error.message);
    const cleanText = (result.candidates?.[0]?.content?.parts?.[0]?.text || '{}').replace(/```json|```/g, '').trim();
    return JSON.parse(cleanText);
  } catch (error) {
    logError('parseAndClassifyExpenseWithGemini', error);
    return null;
  }
}

function parseTravelStartWithGemini(text) {
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    const prompt = `你是一個專業的家庭記帳助手。請從以下使用者的輸入中，擷取「開啟旅遊模式」的相關設定。
    使用者輸入：「${text}」
    請提取出以下 6 個欄位，並以 JSON 格式回傳：
    1. "destination": 字串，旅遊目的地（例如：北海道、日本、沖繩、韓國）。
    2. "days": 數字，旅遊天數（若未提及，預設回傳 5）。
    3. "budget": 數字，旅遊總預算金額（請直接提取數字，例如 20萬 回傳 200000。若無明確提及，請回傳 null）。
    4. "budget_currency": 字串，預算的幣別代碼（例如 "JPY", "TWD"）。若使用者有提及「日幣/日圓」請回傳 "JPY"，未特別指明則預設回傳 "TWD"。
    5. "currency": 字串，當地消費幣別的國際代碼（請根據目的地自動推斷，例如去日本回傳 "JPY"，去韓國回傳 "KRW"，去歐洲回傳 "EUR"，若未提及或在台灣則回傳 null）。
    6. "exchange_rate": 數字，使用者自訂的外幣匯率（若對話中沒明確提到匯率數字，請回傳 null）。
    ⛔ 嚴格限制：請只回傳合法的 JSON 物件，絕對不要包含任何 Markdown 標記或額外文字。`;

    const response = fetchWithRetry(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0 } }),
      muteHttpExceptions: true
    });

    const result = safeJSONParse(response.getContentText());
    if (result.error) throw new Error(result.error.message);
    const cleanText = (result.candidates?.[0]?.content?.parts?.[0]?.text || '{}').replace(/```json|```/g, '').trim();
    return JSON.parse(cleanText);
  } catch (error) {
    logError('parseTravelStartWithGemini', error);
    return null;
  }
}

function parseFixedScheduleWithGemini(text) {
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    const categoryList = CONFIG.CATEGORIES.join('、');
    const projectList = CONFIG.PROJECTS.map(p => `- ${p.name}（${p.desc}）`).join('\n');
    const prompt = `你是一個專業的家庭記帳助手。請從以下使用者的輸入中，擷取「每月固定收支」的設定資訊。
    使用者輸入：「${text}」
    請提取出以下 8 個欄位，並以 JSON 格式回傳：
    1. "day": 數字 (1~31)，代表每個月的幾號。
    2. "type": 字串，必須是 "收入" 或 "支出"。
    3. "item": 字串，項目名稱。
    4. "amount": 數字，金額大小。
    5. "category": 字串，請從以下清單擇一：${categoryList}
    6. "project": 字串，請從以下專案清單擇一：\n${projectList} （若無法判斷請回傳 null）
    7. "payment": 字串，判斷支付方式 (例如：銀行轉帳、自動扣款、信用卡、現金等)
    8. "end_month": 字串，代表這筆固定收支的結束月份（格式為 "YYYY-MM"）。若無提及到期日、分期期數或期限，請務必回傳 null。若有說分幾期，請從現在算起推算出結束的年月。
    ⛔ 嚴格限制：請只回傳合法的 JSON 物件，絕對不要包含任何 Markdown 標記或額外文字。`;

    const response = fetchWithRetry(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0 } }),
      muteHttpExceptions: true
    });

    const result = safeJSONParse(response.getContentText());
    if (result.error) throw new Error(result.error.message);
    const cleanText = (result.candidates?.[0]?.content?.parts?.[0]?.text || '{}').replace(/```json|```/g, '').trim();
    return JSON.parse(cleanText);
  } catch (error) {
    logError('parseFixedScheduleWithGemini', error);
    return null;
  }
}

function parseUpdateFixedScheduleWithGemini(text) {
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    const prompt = `你是一個專業的家庭記帳助手。請從以下使用者的輸入中，擷取「修改固定收支」的意圖與資訊。
    使用者輸入：「${text}」
    請提取出以下資訊，並以 JSON 格式回傳：
    1. "target_keyword": 字串，使用者想修改的目標項目名稱（請提取關鍵字即可，例如：房貸、Netflix、孝親費）。
    2. "new_amount": 數字，修改後的新金額（若無提及請填 null）。
    3. "new_day": 數字 (1~31)，修改後的新扣款日期（若無提及請填 null）。
    4. "new_payment": 字串，修改後的新支付方式（若無提及請填 null）。
    5. "new_end_month": 字串，修改後的新期限/結束月份，格式為 "YYYY-MM"（例如：只繳到明年三月 -> 填入對應的年月。若直接說"無限期"或"取消期限"，請回傳 "CLEAR"。若完全沒提到期限，請填 null）。
    ⛔ 嚴格限制：請只回傳合法的 JSON 物件，絕對不要包含任何 Markdown 標記或額外文字。`;

    const response = fetchWithRetry(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0 } }),
      muteHttpExceptions: true
    });

    const result = safeJSONParse(response.getContentText());
    if (result.error) throw new Error(result.error.message);
    const cleanText = (result.candidates?.[0]?.content?.parts?.[0]?.text || '{}').replace(/```json|```/g, '').trim();
    return JSON.parse(cleanText);
  } catch (error) {
    logError('parseUpdateFixedScheduleWithGemini', error);
    return null;
  }
}

function classifyWithGemini(item, userName) {
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    const categoryList = CONFIG.CATEGORIES.map(c => `- ${c}`).join('\n');
    const projectList = CONFIG.PROJECTS.map(p => `- ${p.name}（${p.desc}）`).join('\n');
    const prompt = `你是一個家庭記帳分類助理。請根據這筆消費的實際用途和目的判斷分類。
    目前說話的使用者是：「${userName}」
    消費項目：「${item}」

    【🚨 嚴格分類規則 🚨】
    - 「汽車機車」：停車費、加油、洗車、ETC過路費、保養等。（注意：停車費與加油絕不是交通）
    - 「交通」：大眾運輸，如計程車、公車、捷運、高鐵、機票。
    - 「家電」：冰箱、洗衣機、冷氣、電視、電鍋、微波爐、除濕機、空氣清淨機、吸塵器、掃地機器人、熱水器等家用電器。不可歸類到居家！
    - 「居家」：裝潢、家具、收納、清潔用品等，不含電器。
    - 「折扣回饋」：折扣、優惠、回饋金、點數兌換、紅利折抵等，此類金額為負數。
    分類（擇一）：\n${categoryList}
    專案（擇一）：\n${projectList}
    請只回傳 JSON，不要有其他文字：
    {"category": "飲食", "project": "🏠 一般開銷"}`;

    const response = fetchWithRetry(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0 } }),
      muteHttpExceptions: true
    });

    const result = safeJSONParse(response.getContentText());
    if (result.error) { return fallbackClassify(item); }
    const cleanText = (result.candidates?.[0]?.content?.parts?.[0]?.text || '{}').replace(/```json|```/g, '').trim();
    const parsed = JSON.parse(cleanText);
    const category = CONFIG.CATEGORIES.includes(parsed.category) ? parsed.category : CONFIG.DEFAULT_CATEGORY;
    const validProjectNames = CONFIG.PROJECTS.map(p => p.name);
    const project = validProjectNames.includes(parsed.project) ? parsed.project : CONFIG.DEFAULT_PROJECT;
    return { category, project };
  } catch (error) {
    logError('classifyWithGemini', error);
    return fallbackClassify(item);
  }
}

function fallbackClassify(item) {
  return { category: getCategory(item), project: getProject(item) };
}

function batchClassifyWithGemini(itemsArray) {
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    const categoryList = CONFIG.CATEGORIES.join('、');
    const projectList = CONFIG.PROJECTS.map(p => p.name).join('、');
    const prompt = `你是一個專業的家庭記帳分類助理。請將以下 JSON 陣列中的消費明細進行分類。
    分類選項：${categoryList}
    專案選項：${projectList}
    【🚨 嚴格分類規則 🚨】
    - 「汽車機車」：停車費、加油、洗車、ETC過路費、保養等。（注意：停車費與加油千萬不可分到交通）
    - 「交通」：大眾運輸，如計程車、公車、捷運、高鐵、機票。
    - 「家電」：冰箱、洗衣機、冷氣、電視、電鍋、微波爐、除濕機、空氣清淨機、吸塵器、掃地機器人、熱水器等家用電器。不可歸類到居家！
    - 「居家」：裝潢、家具、收納、清潔用品等，不含電器。
    - 「折扣回饋」：折扣、優惠、回饋金、點數兌換、紅利折抵等，此類金額為負數。
    請只回傳合法的 JSON 陣列，絕對不要包含任何 Markdown 標記或額外文字。
    實際輸入資料：\n${JSON.stringify(itemsArray)}`;

    const response = fetchWithRetry(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0 } }),
      muteHttpExceptions: true
    });

    const result = safeJSONParse(response.getContentText());
    if (result.error) throw new Error(result.error.message);
    const cleanText = (result.candidates?.[0]?.content?.parts?.[0]?.text || '[]').replace(/```json|```/g, '').trim();
    return JSON.parse(cleanText);
  } catch (error) {
    logError('batchClassifyWithGemini', error);
    return null;
  }
}

function parseEditLastEntryWithGemini(text) {
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    const categoryList = CONFIG.CATEGORIES.join('、');
    const prompt = `你是一個家庭記帳助手。使用者想修改「上一筆」記帳記錄的某個欄位。
    使用者輸入：「${text}」
    請判斷使用者想修改哪個欄位與新的值，以 JSON 格式回傳：
    - "field": 必須是 "item"（品名）、"amount"（金額）、"category"（分類）、"payment"（支付方式）、"card"（卡別）其中一個
    - "value": 新的值（若 field 為 amount，請回傳純數字；若 field 為 category，請從以下清單擇一：${categoryList}；若 field 為 card，直接回傳卡別名稱字串）
    ⛔ 嚴格限制：只回傳合法的 JSON 物件，不要有任何 Markdown 標記或說明文字。`;

    const response = fetchWithRetry(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0 } }),
      muteHttpExceptions: true
    });

    const result = safeJSONParse(response.getContentText());
    if (result.error) throw new Error(result.error.message);
    const cleanText = (result.candidates?.[0]?.content?.parts?.[0]?.text || '{}').replace(/```json|```/g, '').trim();
    return JSON.parse(cleanText);
  } catch (error) {
    logError('parseEditLastEntryWithGemini', error);
    return null;
  }
}
