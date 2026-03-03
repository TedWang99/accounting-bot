// ============================================
// 🤖 LINE 家庭記帳機器人 v8.0 (模組化版)
// ============================================
// v8.0 更新項目:
// [FEATURE] 刪除前確認 Flex：說「刪除」改為先顯示確認卡片，按鈕確認後才真正刪除，防手滑誤刪。
// ============================================
// v7.9 更新項目:
// [FEATURE] 重複記帳偵測：同日相同金額或相同品名首詞，記帳後自動附加警示。
// [FEATURE] 查詢結果 Flex 化：「查 關鍵字」改為 Flex 卡片顯示累計金額與最近清單。
// ============================================
// v7.8 更新項目:
// [FEATURE] 上一筆顯示卡別：「上一筆」指令回傳現在包含卡別欄位。
// [FEATURE] 未標卡別提示：週報附帶「本週有 X 筆信用卡未標卡別」提醒。
// [FEATURE] Ted vs 老婆消費對比：月報新增兩人各自花費對比區塊。
// [FEATURE] 分類預算：為特定消費分類設月上限（CATEGORY_BUDGETS），記帳超標即警示。
// [FEATURE] 年度回顧推播：每年 12/31 23:00 自動推播全年 Top 10 消費 + 各月趨勢 + 兩人對比。
// ============================================
// v7.7 更新項目:
// [FEATURE] 信用卡/悠遊卡卡別記錄：帳本新增第 10 欄「卡別」。
// [FEATURE] OCR 自動偵測末四碼：掃描刷卡明細照片時，自動提取信用卡末四碼並比對 CARDS_MAP 標記卡別。
// [FEATURE] 悠遊卡感應補記：支援輸入「這筆用老婆永豐悠遊卡」手動補記卡別。
// [FEATURE] 多品項整批標記：刷卡明細對應多筆發票品項時，互動式確認後整批更新卡別。
// ============================================
// v7.6 更新項目:
// [REFACTOR] 模組化拆分：將單一 7.5.js (3,122 行) 拆分為 11 個獨立 GAS 模組，提升可維護性。
// ============================================
// v7.5 更新項目:
// [FEATURE] 修改上一筆記帳：可修正最後一筆的金額/分類/支付/品名，無需刪除重打。
// [FEATURE] 日期範圍查詢：支援「查 上週」「查 本週」「查 2026/01」「查 日期~日期」。
// [FEATURE] 每週支出摘要推播：每週一 8:00 自動推送上週花費統計與圓餅圖。
// ============================================
// v7.4 更新項目:
// [FEATURE] 新增「折扣回饋」分類：支援負數金額，適用折扣/優惠/回饋金/點數兌換/紅利等。
// [FEATURE] 強化「家電」分類：修正家電用品被誤歸入「居家」的問題，補充完整關鍵字與 AI 規則。
// [FIX] USERS_MAP 改從 PropertiesService 讀取，不再硬編碼 LINE User ID。
// [FIX] fetchWithRetry 重試條件修正：5xx 與 429 才重試，其餘 4xx 直接回傳。
// [FIX] 重新分類功能改為只針對分類為「其他」的項目執行，節省 API 用量。
// ============================================
// v7.3 新增項目:
// [FEATURE] 視覺化圖表支援：透過 QuickChart 產生每月分類圓餅圖。
// [FEATURE] 互動式審核機制：大額消費 (>5000) 利用 CacheService 先傳送 Flex Message 供使用者確認。
// [FEATURE] 全面 Flex Message 化：記帳成功與報表皆改由精美的 LINE UI 呈現。
// [UPDATE] 新增 handlePostbackEvent 攔截器，完善處理按鈕交互。
// ============================================

// ============================================
// Config - LINE 家庭記帳機器人 v7.7
// ============================================

let _sysPropsDict = null;
const AppProps = {
  getProperty: (key) => {
    if (!_sysPropsDict) _sysPropsDict = PropertiesService['getScriptProperties']().getProperties();
    return _sysPropsDict[key] || null;
  },
  setProperty: (key, val) => {
    if (!_sysPropsDict) _sysPropsDict = PropertiesService['getScriptProperties']().getProperties();
    _sysPropsDict[key] = String(val);
    PropertiesService['getScriptProperties']().setProperty(key, String(val));
    return AppProps;
  },
  deleteProperty: (key) => {
    if (!_sysPropsDict) _sysPropsDict = PropertiesService['getScriptProperties']().getProperties();
    delete _sysPropsDict[key];
    PropertiesService['getScriptProperties']().deleteProperty(key);
    return AppProps;
  }
};

const CONFIG = {
  get SHEET_ID() { return AppProps.getProperty('SHEET_ID'); },
  get LINE_ACCESS_TOKEN() { return AppProps.getProperty('LINE_ACCESS_TOKEN'); },
  get GEMINI_API_KEY() { return AppProps.getProperty('GEMINI_API_KEY'); },
  get LINE_CHANNEL_SECRET() { return AppProps.getProperty('LINE_CHANNEL_SECRET'); },

  // 雲端硬碟資料夾 ID
  get UNIMPORTED_FOLDER_ID() { return AppProps.getProperty('UNIMPORTED_FOLDER_ID'); },
  get IMPORTED_FOLDER_ID() { return AppProps.getProperty('IMPORTED_FOLDER_ID'); },

  MONTHLY_BUDGET: 60000,
  TIMEZONE: 'Asia/Taipei',
  // 請在「指令碼屬性」新增 USERS_MAP，格式為 JSON 字串：
  // {"U8286dadf...": "Ted (老公)", "U11ba13ba...": "老婆大人"}
  get USERS() {
    try {
      const raw = AppProps.getProperty('USERS_MAP');
      return raw ? JSON.parse(raw) : {};
    } catch (e) {
      return {};
    }
  },
  DEFAULT_PAYMENT: '現金',
  DEFAULT_CATEGORY: '其他',
  DEFAULT_PROJECT: '🏠 一般開銷',
  CATEGORIES: [
    '飲食', '日常用品', '交通', '水電瓦斯',
    '電話網路', '居家', '家電', '老公服飾', '老婆服飾', '汽車機車',
    '娛樂', '美容美髮', '交際應酬', '學習深造',
    '保險', '稅金', '醫療保健', '教育',
    'Travel', '3c數位', '點心', '酒',
    '飲料', '保健食品', '保養品', '運動',
    '胖嘟嘟玩具', '可愛嘟嘟玩具', '捐款', '生寶寶',
    '工資', '轉帳手續費', '折扣回饋', '收入', '其他'
  ],
  PROJECTS: [
    { name: '👶 生寶寶／備孕', desc: '與備孕、婦產科、試管、生產相關的醫療費用' },
    { name: '🎮 胖嘟嘟玩具', desc: 'Ted(老公)專屬娛樂花費：電玩、Steam、ROG、F1周邊、男孩公仔模型；以及老公個人使用的 Apple 裝置與電腦 (如 MacBook、iPhone、iPad)、高效能電腦周邊等 3C 產品' },
    { name: '🪆 可愛嘟嘟玩具', desc: '老婆專屬娛樂花費：胖丁、吉伊卡哇、可愛扭蛋、娃娃公仔' },
    { name: '🏠 一般開銷', desc: '以上都不符合則選此項' }
  ],
  PROJECT_BUDGETS: {
    '🎮 胖嘟嘟玩具': 2000,
    '🪆 可愛嘟嘟玩具': 2000
  },
  INVOICE_BATCH_SIZE: 20,  // [v5.5 fix] 原為魔術數字，移至 CONFIG 統一管理
  // [v7.7] 信用卡/悠遊卡設定（存於 ScriptProperties，Key: CARDS_MAP）
  // 格式範例：{"國泰末1234": {"owner": "Ted (老公)", "bank": "國泰", "last4": "1234", "type": "credit"}, ...}
  get CARDS_MAP() {
    try { return JSON.parse(AppProps.getProperty('CARDS_MAP') || '{}'); }
    catch (e) { return {}; }
  },
  // [v7.8] 分類預算（存於 ScriptProperties，Key: CATEGORY_BUDGETS）
  // 格式範例：{"飲食": 15000, "娛樂": 3000, "飲料": 2000}
  get CATEGORY_BUDGETS() {
    try { return JSON.parse(AppProps.getProperty('CATEGORY_BUDGETS') || '{}'); }
    catch (e) { return {}; }
  },
  COMMANDS: {
    HELP: ['說明', '指令', 'help', '?'],
    LAST: ['上一筆', 'last'],
    MONTHLY: ['查詢', '本月', 'total'],
    DELETE: ['刪除', 'delete', 'del'],
    REPORT: ['月報', '報表', 'report'],
    STATS: ['分類', '統計', 'stats']
  }
};
