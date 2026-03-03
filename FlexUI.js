// ============================================
// FlexUI - LINE 家庭記帳機器人 v7.8
// ============================================

function buildExpenseMessage(userName, item, category, project, amount, payment, originalAmountStr, expenseDate) {
  const dateStr = Utilities.formatDate(expenseDate || new Date(), CONFIG.TIMEZONE, 'MM/dd');
  const line = '━━━━━━━━━━';

  // 收入記帳：綠色標示，金額顯示為正數
  if (category === '收入' && amount < 0) {
    const incomeAmt = formatCurrency(Math.abs(amount));
    return `💵 ${userName} 收入記帳！\n${line}\n🗓️ 日期：${dateStr}\n📝 項目：${item}\n📂 分類：${category}\n💰 金額：+${incomeAmt}\n${line}`;
  }

  let amountLine;
  if (originalAmountStr) {
    amountLine = `${formatCurrency(amount)} (${originalAmountStr})`;
  } else { amountLine = formatCurrency(amount); }

  return `✅ ${userName} 記帳成功！\n${line}\n🗓️ 日期：${dateStr}\n📝 項目：${item}\n📂 分類：${category}\n🚀 專案：${project}\n💳 支付：${payment}\n💰 金額：${amountLine}\n${line}`;
}

function buildProgressBarBox(pct) {
  // 防呆與計算寬度
  const validPct = Math.max(0, pct);
  const fillWidth = validPct >= 100 ? '100%' : `${validPct.toFixed(1)}%`;

  // 決定顏色：超過100%為紅色，超過80%為橘黃色，其餘為綠色
  let fillColor = '#00B900'; // 預設 LINE 綠
  if (validPct >= 100) fillColor = '#FF334B';
  else if (validPct >= 80) fillColor = '#FF9900';

  return {
    type: "box",
    layout: "vertical",
    contents: [
      {
        type: "box",
        layout: "horizontal",
        contents: [
          {
            type: "box",
            layout: "vertical",
            contents: [],
            width: fillWidth,
            backgroundColor: fillColor,
            cornerRadius: "10px"
          }
        ],
        backgroundColor: "#E0E0E0",
        height: "8px",
        cornerRadius: "10px",
        margin: "sm"
      }
    ]
  };
}

// --- v7.2 預算狀態儀表板 Flex Message ---
function buildBudgetFlexMessage(data) {
  const contents = [];

  // 標題區域
  contents.push({
    type: "box",
    layout: "vertical",
    contents: [
      { type: "text", text: "📊 預算狀態儀表板", weight: "bold", size: "xl", color: "#1DB446" },
      { type: "text", text: `${data.month} 月份結算`, size: "xs", color: "#aaaaaa", margin: "md" }
    ],
    paddingAll: "20px",
    paddingBottom: "10px"
  });

  // --- 1. 總預算區塊 ---
  const mainStatusColor = data.main.pct >= 100 ? "#FF334B" : (data.main.pct >= 80 ? "#FF9900" : "#555555");
  contents.push({
    type: "box",
    layout: "vertical",
    contents: [
      {
        type: "box",
        layout: "horizontal",
        contents: [
          { type: "text", text: "💰 總預算", size: "md", weight: "bold", color: "#333333", flex: 1 },
          { type: "text", text: `${data.main.pct.toFixed(1)}%`, size: "md", weight: "bold", color: mainStatusColor, align: "end", flex: 1 }
        ]
      },
      buildProgressBarBox(data.main.pct),
      {
        type: "box",
        layout: "horizontal",
        contents: [
          { type: "text", text: `已用 ${formatCurrency(data.main.spent)}`, size: "xs", color: "#aaaaaa", flex: 1 },
          { type: "text", text: data.main.remaining >= 0 ? `剩餘 ${formatCurrency(data.main.remaining)}` : `超支 ${formatCurrency(Math.abs(data.main.remaining))}`, size: "xs", color: mainStatusColor, align: "end", flex: 1 }
        ],
        margin: "sm"
      }
    ],
    paddingAll: "20px",
    paddingTop: "10px",
    paddingBottom: "15px"
  });

  // 分隔線
  if (data.projects && data.projects.length > 0) {
    contents.push({ type: "separator", color: "#dddddd" });
  }

  // --- 2. 各專案子預算區塊 ---
  if (data.projects) {
    data.projects.forEach(p => {
      const pStatusColor = p.pct >= 100 ? "#FF334B" : (p.pct >= 80 ? "#FF9900" : "#555555");
      contents.push({
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: p.name, size: "sm", weight: "bold", color: "#333333", flex: 1 },
              { type: "text", text: `${p.pct.toFixed(1)}%`, size: "sm", weight: "bold", color: pStatusColor, align: "end", flex: 1 }
            ]
          },
          buildProgressBarBox(p.pct),
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: `已用 ${formatCurrency(p.spent)}`, size: "xs", color: "#aaaaaa", flex: 1 },
              { type: "text", text: p.remaining >= 0 ? `剩 ${formatCurrency(p.remaining)}` : `超支 ${formatCurrency(Math.abs(p.remaining))}`, size: "xs", color: pStatusColor, align: "end", flex: 1 }
            ],
            margin: "sm"
          }
        ],
        paddingAll: "20px",
        paddingTop: "15px",
        paddingBottom: "15px"
      });
    });
  }

  // [v7.8] 分類預算區塊
  if (data.categories && data.categories.length > 0) {
    contents.push({ type: "separator", color: "#dddddd" });
    contents.push({
      type: "box", layout: "vertical",
      contents: [
        { type: "text", text: "🏷️ 分類預算", weight: "bold", size: "sm", color: "#8E44AD", margin: "md" }
      ],
      paddingAll: "20px", paddingBottom: "5px"
    });
    data.categories.forEach(c => {
      const cStatusColor = c.pct >= 100 ? "#FF334B" : (c.pct >= 80 ? "#FF9900" : "#555555");
      contents.push({
        type: "box", layout: "vertical",
        contents: [
          {
            type: "box", layout: "horizontal",
            contents: [
              { type: "text", text: c.name, size: "sm", weight: "bold", color: "#333333", flex: 1 },
              { type: "text", text: `${c.pct.toFixed(1)}%`, size: "sm", weight: "bold", color: cStatusColor, align: "end", flex: 1 }
            ]
          },
          buildProgressBarBox(c.pct),
          {
            type: "box", layout: "horizontal",
            contents: [
              { type: "text", text: `已用 ${formatCurrency(c.spent)}`, size: "xs", color: "#aaaaaa", flex: 1 },
              { type: "text", text: c.remaining >= 0 ? `剩 ${formatCurrency(c.remaining)}` : `超支 ${formatCurrency(Math.abs(c.remaining))}`, size: "xs", color: cStatusColor, align: "end", flex: 1 }
            ],
            margin: "sm"
          }
        ],
        paddingAll: "20px", paddingTop: "10px", paddingBottom: "10px"
      });
    });
  }

  // Footer
  contents.push({
    type: "box",
    layout: "vertical",
    contents: [
      { type: "text", text: "記帳小幫手 v7.8", size: "xxs", color: "#cccccc", align: "center" }
    ],
    paddingAll: "15px",
    backgroundColor: "#fafafa"
  });

  return {
    type: "bubble",
    size: "mega",
    body: {
      type: "box",
      layout: "vertical",
      contents: contents,
      paddingAll: "0px"
    }
  };
}

function formatMonthlyReport(month, year, stats) {
  const { total, totalIncome, count, categoryStats, paymentStats } = stats; // [v5.5 fix] 加入 totalIncome
  const topCategory = Object.entries(categoryStats).sort((a, b) => b[1] - a[1]).slice(0, 3);
  const now = new Date();
  const daysInMonth = (month === now.getMonth() + 1 && year === now.getFullYear())
    ? now.getDate()
    : new Date(year, month, 0).getDate();
  const avgDaily = Math.round(total / daysInMonth);
  const lineSep = '━━━━━━━━━━';
  const incomeStr = totalIncome > 0 ? `\n📥 本月收入：${formatCurrency(totalIncome)}` : ''; // [v5.5 fix] 顯示收入
  let message = `📊 ${month}月份消費報表\n${lineSep}${incomeStr}\n💰 總支出：${formatCurrency(total)}\n📝 總筆數：${count} 筆\n📊 日均支出：${formatCurrency(avgDaily)}\n${lineSep}\n\n🏆 支出排行 TOP 3：\n`;
  topCategory.forEach((item, index) => {
    const pct = total > 0 ? (item[1] / total * 100).toFixed(1) : '0.0'; // [v5.5 fix] 防止 total=0 時除以零
    message += `${index + 1}. ${item[0]}：${formatCurrency(item[1])} (${pct}%)\n`;
  });
  message += `\n💳 支付方式：\n`;
  Object.entries(paymentStats).forEach(([key, value]) => { message += `• ${key}：${formatCurrency(value)}\n`; });
  return message;
}

function buildReviewFlexMessage(record, uuid) {
  return {
    type: "bubble",
    size: "mega",
    body: {
      type: "box", layout: "vertical", spacing: "md",
      contents: [
        { type: "text", text: "⚠️ 待確認消費", weight: "bold", color: "#FF334B", size: "lg" },
        {
          type: "box", layout: "vertical", margin: "md", spacing: "sm",
          contents: [
            { type: "text", text: `項目：${record.item}`, size: "md", wrap: true },
            { type: "text", text: `分類：${record.category}`, size: "sm", color: "#666666" },
            { type: "text", text: `專案：${record.project}`, size: "sm", color: "#666666" },
            { type: "text", text: `金額：NT$ ${record.amount.toLocaleString()}`, size: "xl", weight: "bold", color: "#111111" },
            ...(record.installments && record.installments > 1 ? [{
              type: "box", layout: "horizontal", margin: "sm",
              contents: [
                { type: "text", text: `💳 分期`, size: "sm", color: "#FF9900", weight: "bold" },
                { type: "text", text: `${record.installments} 期 × NT$ ${Math.round(record.amount / record.installments).toLocaleString()}/月`, size: "sm", color: "#FF9900", align: "end", weight: "bold" }
              ]
            }] : [])
          ]
        },
        { type: "text", text: "這是一筆大額消費，請點擊確認寫入或是直接丟棄這筆紀錄。", size: "sm", color: "#999999", wrap: true }
      ]
    },
    footer: {
      type: "box", layout: "horizontal", spacing: "sm",
      contents: [
        { type: "button", style: "primary", color: "#28B463", action: { type: "postback", label: "✅ 正確, 寫入", data: `action=confirm&id=${uuid}` } },
        { type: "button", style: "secondary", action: { type: "postback", label: "❌ 手殘取消", data: `action=cancel&id=${uuid}` } }
      ]
    }
  };
}

function buildReceiptFlexMessage(record) {
  // 日期：支援 Date 物件或 ISO 字串（從 CacheService 還原時為字串）
  let dateStr = '';
  try { dateStr = Utilities.formatDate(new Date(record.date), CONFIG.TIMEZONE, 'MM/dd'); } catch (e) { dateStr = ''; }

  // 外幣原始金額：item 格式為「品名(¥3,000)」，擷取括號內容
  const foreignMatch = record.item.match(/\(([¥$€₩฿][\d,]+|[A-Z]{2,3}\$?[\d,]+)\)$/);
  const foreignStr = foreignMatch ? foreignMatch[1] : null;

  const detailRows = [
    { type: "box", layout: "horizontal", contents: [{ type: "text", text: "日期", size: "sm", color: "#555555" }, { type: "text", text: dateStr, size: "sm", color: "#111111", align: "end" }] },
    { type: "box", layout: "horizontal", contents: [{ type: "text", text: "分類", size: "sm", color: "#555555" }, { type: "text", text: record.category, size: "sm", color: "#111111", align: "end" }] },
    { type: "box", layout: "horizontal", contents: [{ type: "text", text: "專案", size: "sm", color: "#555555" }, { type: "text", text: record.project, size: "sm", color: "#111111", align: "end" }] },
    { type: "box", layout: "horizontal", contents: [{ type: "text", text: "支付", size: "sm", color: "#555555" }, { type: "text", text: record.payment, size: "sm", color: "#111111", align: "end" }] },
    { type: "box", layout: "horizontal", contents: [{ type: "text", text: "紀錄人", size: "sm", color: "#555555" }, { type: "text", text: record.userName, size: "sm", color: "#111111", align: "end" }] }
  ];
  // [v7.7] 卡別（有值才顯示，插入支付與紀錄人之間）
  if (record.cardId) {
    detailRows.splice(4, 0, {
      type: "box", layout: "horizontal",
      contents: [
        { type: "text", text: "卡別", size: "sm", color: "#555555" },
        { type: "text", text: record.cardId, size: "sm", color: "#0066CC", align: "end", weight: "bold" }
      ]
    });
  }
  if (foreignStr) {
    detailRows.push({ type: "box", layout: "horizontal", contents: [{ type: "text", text: "原始金額", size: "sm", color: "#555555" }, { type: "text", text: foreignStr, size: "sm", color: "#888888", align: "end" }] });
  }

  // [v7.3+] 分期資訊區塊（若有分期）
  if (record.installments && record.installments > 1) {
    const startDate = new Date(record.date);
    const monthlyAmt = Math.round(record.amount / record.installments);
    const endDate = new Date(startDate.getFullYear(), startDate.getMonth() + record.installments - 1, 1);
    const endMonthStr = Utilities.formatDate(endDate, CONFIG.TIMEZONE, 'yyyy-MM');
    const safeDay = Math.min(record.installment_day || startDate.getDate(), 28);
    detailRows.push({ type: "separator", margin: "sm" });
    detailRows.push({
      type: "box", layout: "horizontal",
      contents: [
        { type: "text", text: "💳 分期入帳", size: "sm", color: "#FF9900", weight: "bold" },
        { type: "text", text: `${record.installments} 期 × NT$ ${monthlyAmt.toLocaleString()}`, size: "sm", color: "#FF9900", align: "end", weight: "bold" }
      ]
    });
    detailRows.push({
      type: "box", layout: "horizontal",
      contents: [
        { type: "text", text: "每月扣款日", size: "xs", color: "#aaaaaa" },
        { type: "text", text: `每月 ${safeDay} 號，至 ${endMonthStr} 止`, size: "xs", color: "#aaaaaa", align: "end" }
      ]
    });
  }

  return {
    type: "bubble",
    size: "mega",
    body: {
      type: "box", layout: "vertical",
      contents: [
        { type: "text", text: "✅ 記帳成功", weight: "bold", color: "#1DB446", size: "sm" },
        { type: "text", text: `NT$ ${record.amount.toLocaleString()}`, weight: "bold", size: "xxl", margin: "md" },
        { type: "text", text: record.item, size: "xs", color: "#aaaaaa", wrap: true },
        { type: "separator", margin: "xxl" },
        { type: "box", layout: "vertical", margin: "xxl", spacing: "sm", contents: detailRows },
        { type: "separator", margin: "xxl" },
        { type: "box", layout: "horizontal", margin: "md", contents: [{ type: "text", text: "Source", size: "xs", color: "#aaaaaa" }, { type: "text", text: record.source || "LINE", color: "#aaaaaa", size: "xs", align: "end" }] }
      ]
    }
  };
}

function buildReportFlexMessage(month, stats, _imageUrl, aiAdvice = '') {
  const contents = [
    { type: "text", text: `${month}月份 財務報表`, weight: "bold", size: "xl", color: "#1DB446" },
    {
      type: "box", layout: "vertical", margin: "md", spacing: "sm", contents: [
        { type: "box", layout: "horizontal", contents: [{ type: "text", text: "本月總支出", color: "#555555" }, { type: "text", text: `NT$ ${stats.totalExpense.toLocaleString()}`, align: "end", weight: "bold", color: "#FF334B" }] },
        { type: "box", layout: "horizontal", contents: [{ type: "text", text: "本月總額外收入", color: "#555555" }, { type: "text", text: `NT$ ${stats.totalIncome.toLocaleString()}`, align: "end", weight: "bold", color: "#28B463" }] },
        { type: "box", layout: "horizontal", contents: [{ type: "text", text: "日常預算", color: "#555555" }, { type: "text", text: `NT$ ${CONFIG.MONTHLY_BUDGET.toLocaleString()}`, align: "end", color: "#111111" }] }
      ]
    },
    { type: "separator", margin: "xl" }
  ];

  const topCategories = Object.entries(stats.categoryStats)
    .filter(([k, v]) => k !== '工資' && v > 0)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  if (topCategories.length > 0) {
    const catBox = { type: "box", layout: "vertical", margin: "xl", spacing: "sm", contents: [{ type: "text", text: "消費分類排行 (Top 5)", weight: "bold", size: "sm", color: "#111111" }] };
    topCategories.forEach(([cat, amt]) => {
      catBox.contents.push({
        type: "box", layout: "horizontal",
        contents: [
          { type: "text", text: cat, size: "sm", color: "#555555" },
          { type: "text", text: `NT$ ${amt.toLocaleString()}`, size: "sm", color: "#111111", align: "end" }
        ]
      });
    });
    contents.push(catBox);
  }

  // [v7.7] 各卡消費排行（有卡別資料才顯示）
  if (stats.cardStats && Object.keys(stats.cardStats).length > 0) {
    const topCards = Object.entries(stats.cardStats)
      .filter(([, v]) => v > 0)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5);
    if (topCards.length > 0) {
      contents.push({ type: "separator", margin: "xl" });
      const cardBox = {
        type: "box", layout: "vertical", margin: "xl", spacing: "sm",
        contents: [{ type: "text", text: "💳 各卡消費排行", weight: "bold", size: "sm", color: "#0066CC" }]
      };
      topCards.forEach(([cardKey, amt]) => {
        cardBox.contents.push({
          type: "box", layout: "horizontal",
          contents: [
            { type: "text", text: cardKey, size: "sm", color: "#555555", flex: 3 },
            { type: "text", text: `NT$ ${amt.toLocaleString()}`, size: "sm", color: "#0066CC", align: "end", flex: 2 }
          ]
        });
      });
      contents.push(cardBox);
    }
  }

  // [v7.8] 兩人消費對比（有 2 人以上資料才顯示）
  if (stats.personStats && Object.keys(stats.personStats).length >= 2) {
    const totalPersonSpend = Object.values(stats.personStats).reduce((s, v) => s + v, 0);
    contents.push({ type: "separator", margin: "xl" });
    const personBox = {
      type: "box", layout: "vertical", margin: "xl", spacing: "sm",
      contents: [{ type: "text", text: "👫 兩人消費對比", weight: "bold", size: "sm", color: "#E74C3C" }]
    };
    Object.entries(stats.personStats).sort((a, b) => b[1] - a[1]).forEach(([person, amt]) => {
      const pct = totalPersonSpend > 0 ? (amt / totalPersonSpend * 100).toFixed(1) : '0.0';
      personBox.contents.push({
        type: "box", layout: "horizontal",
        contents: [
          { type: "text", text: person, size: "sm", color: "#555555", flex: 3 },
          { type: "text", text: `NT$ ${amt.toLocaleString()} (${pct}%)`, size: "sm", color: "#E74C3C", align: "end", flex: 2 }
        ]
      });
    });
    contents.push(personBox);
  }

  // [v7.0] 將 AI 顧問的精華點評加入最後
  if (aiAdvice) {
    contents.push({ type: "separator", margin: "md" });
    contents.push({
      type: "box", layout: "vertical", margin: "xl", spacing: "sm",
      contents: [
        { type: "text", text: "💡 專屬 AI 財務顧問總結", weight: "bold", size: "sm", color: "#8E44AD" },
        { type: "text", text: aiAdvice, size: "sm", color: "#333333", wrap: true }
      ]
    });
  }

  // Footer
  contents.push({ type: "separator", margin: "md" });
  contents.push({
    type: "box", layout: "vertical", margin: "sm",
    contents: [{ type: "text", text: "記帳小幫手 v7.8", size: "xxs", color: "#cccccc", align: "center" }]
  });

  const flex = {
    type: "bubble",
    size: "mega",
    body: { type: "box", layout: "vertical", contents: contents }
  };

  // 若想把圖片直接塞進 flex，可加在這裡；現階段 replyLine 會把它當第二則圖片訊息發送，這樣圖比較大張好看
  return flex;
}

// [v7.9] 關鍵字查詢結果 Flex 卡片
function buildSearchResultFlexMessage(keyword, result) {
  const { total, foldedCount, count, items } = result;
  const lineSep = { type: "separator", margin: "md" };

  const headerBox = {
    type: "box", layout: "vertical", spacing: "sm",
    contents: [
      { type: "text", text: `🔍 ${keyword}`, weight: "bold", size: "lg", color: "#333333" },
      { type: "text", text: `累計 ${formatCurrency(total)}・${foldedCount} 筆`, size: "sm", color: "#888888" }
    ]
  };

  const rowContents = items.map(item => ({
    type: "box", layout: "horizontal", spacing: "sm",
    contents: [
      { type: "text", text: item.date, size: "xs", color: "#888888", flex: 1 },
      { type: "text", text: item.item, size: "xs", color: "#333333", flex: 4, wrap: true },
      { type: "text", text: `$${item.amount.toLocaleString()}`, size: "xs", color: "#111111", flex: 2, align: "end" }
    ]
  }));

  const listBox = {
    type: "box", layout: "vertical", spacing: "sm",
    contents: rowContents.length > 0 ? rowContents : [{ type: "text", text: "（無符合記錄）", size: "sm", color: "#aaaaaa" }]
  };

  return {
    type: "bubble", size: "mega",
    body: {
      type: "box", layout: "vertical", spacing: "md",
      contents: [headerBox, lineSep, listBox]
    },
    footer: {
      type: "box", layout: "horizontal",
      contents: [
        { type: "text", text: `共 ${count} 項明細`, size: "xs", color: "#aaaaaa" },
        { type: "text", text: "v8.0", size: "xs", color: "#aaaaaa", align: "end" }
      ]
    }
  };
}

// [v8.0] 刪除確認 Flex 卡片
function buildDeleteConfirmFlexMessage(record, uuid) {
  let dateStr = '';
  try {
    dateStr = record.date instanceof Date
      ? Utilities.formatDate(record.date, CONFIG.TIMEZONE, 'MM/dd')
      : String(record.date).substring(5, 10).replace('-', '/');
  } catch (e) {}

  return {
    type: "bubble", size: "mega",
    body: {
      type: "box", layout: "vertical", spacing: "md",
      contents: [
        { type: "text", text: "🗑️ 確認刪除？", weight: "bold", color: "#FF334B", size: "lg" },
        { type: "separator", margin: "md" },
        {
          type: "box", layout: "vertical", margin: "md", spacing: "sm",
          contents: [
            { type: "box", layout: "horizontal", contents: [
              { type: "text", text: "項目", size: "sm", color: "#555555", flex: 2 },
              { type: "text", text: record.item, size: "sm", color: "#111111", flex: 5, wrap: true, align: "end" }
            ]},
            { type: "box", layout: "horizontal", contents: [
              { type: "text", text: "金額", size: "sm", color: "#555555", flex: 2 },
              { type: "text", text: `NT$ ${record.amount.toLocaleString()}`, size: "sm", color: "#FF334B", flex: 5, align: "end", weight: "bold" }
            ]},
            { type: "box", layout: "horizontal", contents: [
              { type: "text", text: "日期", size: "sm", color: "#555555", flex: 2 },
              { type: "text", text: dateStr, size: "sm", color: "#111111", flex: 5, align: "end" }
            ]},
            { type: "box", layout: "horizontal", contents: [
              { type: "text", text: "分類", size: "sm", color: "#555555", flex: 2 },
              { type: "text", text: record.category, size: "sm", color: "#111111", flex: 5, align: "end" }
            ]},
            { type: "box", layout: "horizontal", contents: [
              { type: "text", text: "記錄人", size: "sm", color: "#555555", flex: 2 },
              { type: "text", text: record.userName, size: "sm", color: "#111111", flex: 5, align: "end" }
            ]}
          ]
        }
      ]
    },
    footer: {
      type: "box", layout: "horizontal", spacing: "sm",
      contents: [
        { type: "button", style: "primary", color: "#FF334B",
          action: { type: "postback", label: "🗑️ 確認刪除", data: `action=delete_confirm&id=${uuid}` } },
        { type: "button", style: "secondary",
          action: { type: "postback", label: "✅ 取消", data: `action=delete_cancel&id=${uuid}` } }
      ]
    }
  };
}
