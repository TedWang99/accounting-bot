// ============================================
// Invoices - LINE 家庭記帳機器人 v7.7
// ============================================

function reclassifyAllFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    ui.alert('⚠️ 帳本目前沒有資料！');
    return;
  }

  const allData = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const targetRows = allData.reduce((acc, row, i) => {
    if (String(row[1] || '').trim() && String(row[5] || '').trim() === CONFIG.DEFAULT_CATEGORY) {
      acc.push(i);
    }
    return acc;
  }, []);

  if (targetRows.length === 0) {
    ui.alert('✅ 目前沒有分類為「其他」的項目，無需重新分類！');
    return;
  }

  const response = ui.alert(
    '🔄 重新分類【其他】項目',
    `找到 ${targetRows.length} 筆分類為「其他」的資料，將重新進行 AI 分類。\n每筆約 1-2 秒，共需約 ${Math.ceil(targetRows.length * 1.5 / 60)} 分鐘。\n\n確定開始？`,
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const BATCH_LOG_EVERY = 50;
  let changedCount = 0;
  let processedCount = 0;
  const changes = [];

  for (const i of targetRows) {
    const row = allData[i];
    const item = String(row[1] || '').trim();
    const oldCategory = String(row[5] || '').trim();
    const oldProject = String(row[6] || '').trim();
    const rowUser = String(row[3] || '').trim();

    const result = classifyWithGemini(item, rowUser);
    const newCategory = result.category;
    const newProject = result.project;

    if (newCategory !== oldCategory || newProject !== oldProject) {
      const rowNum = i + 2; // 1-indexed, +1 for header
      sheet.getRange(rowNum, 6).setValue(newCategory);
      sheet.getRange(rowNum, 7).setValue(newProject);
      changedCount++;
      const desc = [];
      if (newCategory !== oldCategory) desc.push(`分類 ${oldCategory}→${newCategory}`);
      if (newProject !== oldProject) desc.push(`專案 ${oldProject}→${newProject}`);
      changes.push(`• ${item}：${desc.join('、')}`);
    }

    processedCount++;
    if (processedCount % BATCH_LOG_EVERY === 0) {
      Logger.log(`[重新分類] 進度：${processedCount}/${targetRows.length} 筆`);
    }
    Utilities.sleep(300); // 避免 API rate limit
  }

  // 結果彈窗
  const MAX_SHOW = 20;
  let summary = `✅ 重新分類完成！\n━━━━━━━━━━\n掃描「其他」：${targetRows.length} 筆\n成功分類：${changedCount} 筆`;
  if (changes.length > 0) {
    summary += '\n━━━━━━━━━━\n' + changes.slice(0, MAX_SHOW).join('\n');
    if (changes.length > MAX_SHOW) summary += `\n…（還有 ${changes.length - MAX_SHOW} 筆，請查看 Logger）`;
  }
  ui.alert('🔄 重新分類結果', summary, ui.ButtonSet.OK);
}

function processInvoices(silent = false) {
  const notify = (msg) => {
    if (!silent) SpreadsheetApp.getUi().alert(msg);
    else Logger.log('[自動匯入] ' + msg);
  };
  const unimportedFolderId = CONFIG.UNIMPORTED_FOLDER_ID;
  const importedFolderId = CONFIG.IMPORTED_FOLDER_ID;
  if (!unimportedFolderId || !importedFolderId) {
    return notify('⚠️ 尚未設定資料夾 ID！\n請先在「專案設定 > 指令碼屬性」中新增 UNIMPORTED_FOLDER_ID 與 IMPORTED_FOLDER_ID。');
  }
  let unimportedFolder, importedFolder;
  try {
    unimportedFolder = DriveApp.getFolderById(unimportedFolderId);
    importedFolder = DriveApp.getFolderById(importedFolderId);
  } catch (e) {
    return notify('⚠️ 找不到指定的資料夾！請確認 ID 是否正確且具有權限。');
  }
  const files = unimportedFolder.getFilesByType('text/csv');
  let hasFiles = false;
  let successCount = 0;
  const mainSheet = getSheet();
  const existingInvoices = new Set();
  const mainLastRow = mainSheet.getLastRow();

  if (mainLastRow > 1) {
    mainSheet.getRange(2, 8, mainLastRow - 1, 1).getValues()
      .forEach(row => { if (row[0]) existingInvoices.add(String(row[0]).trim()); });
  }

  while (files.hasNext()) {
    hasFiles = true;
    const file = files.next();
    const fileName = file.getName();

    let fileOwner = 'Ted (老公)';
    if (fileName.includes('6652874') || fileName.includes('老婆') || fileName.toLowerCase().includes('wife')) {
      fileOwner = '老婆大人';
    } else if (fileName.includes('8602864') || fileName.includes('老公') || fileName.toLowerCase().includes('ted')) {
      fileOwner = 'Ted (老公)';
    }

    let csvText = file.getBlob().getDataAsString('UTF-8');
    let csvData = Utilities.parseCsv(csvText);

    if (csvData.length > 0 && String(csvData[0][13]).indexOf('品名') === -1) {
      try {
        csvText = file.getBlob().getDataAsString('Big5');
        csvData = Utilities.parseCsv(csvText);
      } catch (e) { Logger.log('Big5 解析失敗'); }
    }
    if (csvData.length <= 1) {
      file.moveTo(importedFolder);
      continue;
    }
    if (String(csvData[0][13]).indexOf('品名') === -1) {
      Logger.log(`⚠️ 檔案 ${fileName} 格式不符 (找不到 N 欄品名)，略過。`);
      file.moveTo(importedFolder);  // [v5.5 fix] 格式不符也需移走，避免下次重複處理
      continue;
    }

    const pendingItems = [];
    for (let i = 1; i < csvData.length; i++) {
      const row = csvData[i];
      if (row.length < 14) continue;
      const rawDate = row[1];
      const invoiceNum = row[2];
      const sellerName = row[7];
      const itemAmount = row[12];
      const itemName = row[13];
      if (!invoiceNum) continue;
      if (existingInvoices.has(String(invoiceNum).trim())) continue;
      if (!itemName || isNaN(itemAmount) || parseInt(itemAmount) === 0) continue;
      const combinedItemName = `${sellerName} - ${itemName}`.replace(/股份有限公司|有限公司/g, '');
      pendingItems.push({
        id: `${file.getId()}_${i}`,
        date: parseMoFDate(rawDate),
        invoiceNum: String(invoiceNum).trim(),
        item: combinedItemName,
        amount: parseInt(itemAmount),
        recordUser: fileOwner
      });
    }

    if (pendingItems.length > 0) {
      pendingItems.reverse();
      for (let i = 0; i < pendingItems.length; i += CONFIG.INVOICE_BATCH_SIZE) {
        const processBatch = pendingItems.slice(i, i + CONFIG.INVOICE_BATCH_SIZE);
        const geminiInput = processBatch.map(p => ({ id: p.id, item: p.item }));
        let aiResults = batchClassifyWithGemini(geminiInput);

        // [v5.5 fix] 原本 AI 失敗直接 return 中止整個匯入；改為使用關鍵字備援分類繼續處理
        if (!aiResults || aiResults.length === 0) {
          Logger.log(`⚠️ 第 ${Math.floor(i / CONFIG.INVOICE_BATCH_SIZE) + 1} 批 AI 分類失敗，改用關鍵字備援分類繼續處理...`);
          aiResults = processBatch.map(p => {
            const fb = fallbackClassify(p.item);
            return { id: p.id, category: fb.category, project: fb.project };
          });
        }

        // [v5.5 fix] getDataRange 移至 forEach 外，只讀取一次；用 mergedRowIndices 追蹤已合併列
        const currentData = mainSheet.getDataRange().getValues();
        const mergedRowIndices = new Set();

        aiResults.forEach(res => {
          const originalData = processBatch.find(p => p.id === res.id);
          if (originalData) {
            const isShopee = originalData.item.includes('蝦皮');
            const dateOnlyStr = Utilities.formatDate(originalData.date, CONFIG.TIMEZONE, 'yyyy/MM/dd');

            // === [v5.5] 智能合併邏輯 (CSV 補填發票號碼至 OCR 紀錄) ===
            // [v5.5 fix] 收集所有符合條件的列，只有唯一匹配才執行合併，避免同日同額誤合併
            const matchCandidates = [];
            for (let j = currentData.length - 1; j > 0; j--) {
              if (mergedRowIndices.has(j + 1)) continue; // 已被合併的列跳過
              const rowDateObj = new Date(currentData[j][0]);
              if (isNaN(rowDateObj)) continue;
              const rDate = Utilities.formatDate(rowDateObj, CONFIG.TIMEZONE, 'yyyy/MM/dd');
              const rAmount = parseInt(currentData[j][2]) || 0;
              const rInvoice = currentData[j][7];
              // 條件：日期相同、金額相同、且該筆紀錄沒有發票號碼 (代表是 OCR 或手動記帳)
              if (rDate === dateOnlyStr && rAmount === originalData.amount && !rInvoice) {
                matchCandidates.push(j + 1);
              }
            }
            // 僅有唯一匹配時才合併，有多筆相同條件則跳過以防誤合
            const matchIndex = matchCandidates.length === 1 ? matchCandidates[0] : -1;

            if (matchIndex !== -1) {
              // 找到唯一匹配的 OCR 孤兒紀錄，只補上發票號碼，保留精確品名
              mainSheet.getRange(matchIndex, 8).setValue(originalData.invoiceNum);
              mainSheet.getRange(matchIndex, 9).setValue('OCR+CSV智能合併');
              mergedRowIndices.add(matchIndex); // 標記已合併，防止同批次重複合併同一列
            } else {
              // 找不到唯一匹配紀錄，正常新增
              mainSheet.appendRow([
                dateOnlyStr, originalData.item, originalData.amount, originalData.recordUser,
                isShopee ? '現金' : '信用卡', res.category || CONFIG.DEFAULT_CATEGORY,
                res.project || CONFIG.DEFAULT_PROJECT, originalData.invoiceNum, 'CSV匯入'
              ]);
            }

            existingInvoices.add(originalData.invoiceNum);
            successCount++;
          }
        });
        if (i + CONFIG.INVOICE_BATCH_SIZE < pendingItems.length) Utilities.sleep(2000);
      }
    }
    file.moveTo(importedFolder);
  }

  if (!hasFiles) {
    return notify('✅ 雲端資料夾中目前沒有需要匯入的載具檔案！');
  }
  try {
    const lastRow = mainSheet.getLastRow();
    if (lastRow > 1) mainSheet.getRange(2, 1, lastRow - 1, 9).sort({ column: 1, ascending: true });
  } catch (sortError) { logError('processInvoices - sort', sortError); }
  notify(`✅ 全自動掃描完畢！共成功匯入並分類 ${successCount} 筆發票明細。\n(已處理的檔案皆已移至「已匯入」資料夾 ✨)`);
}

function parseMoFDate(rawDate) {
  const str = String(rawDate).trim();
  if (str.length === 8) return new Date(parseInt(str.substring(0, 4)), parseInt(str.substring(4, 6)) - 1, parseInt(str.substring(6, 8)));
  Logger.log(`⚠️ parseMoFDate: 無法解析日期格式「${str}」（長度=${str.length}），改用今日日期替代。`); // [v5.5 fix] 原本靜默回傳今日
  return new Date();
}
