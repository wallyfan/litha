function sheetname() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

function clearB16() {
  const sheet = SpreadsheetApp.getActiveSheet( );
  const contentB16 = sheet.getRange("B16").clearContent( );
}

function clearInputcontent() {
  const sheet = SpreadsheetApp.getActiveSheet( );
  const contentB5to14 = sheet.getRange("B5:B14").clearContent( );
  const contentE5to14 = sheet.getRange("E5:E14").clearContent( );
}

function submitCustomerUsageButtonA() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName("客戶使用key-in");
  var flowSheet = ss.getSheetByName("耗用量流水");
  
  if (!inputSheet || !flowSheet) {
    SpreadsheetApp.getUi().alert("找不到必要的輸入分頁或『耗用量流水』分頁！");
    return;
  }

  // 1. 讀取基本資訊
  var date = inputSheet.getRange("B1").getValue();
  var customerName = inputSheet.getRange("B2").getValue();
  var receiptNo = inputSheet.getRange("B3").getValue();

  if (!date || !customerName) {
    SpreadsheetApp.getUi().alert("請填寫【日期】與【客戶姓名】！");
    return;
  }

  // 2. 讀取多筆項目資料
var validItems = [];
  // 定義要讀取的列數區間：第 5~14 列 及 第 18~27 列
  var rowRanges = [
    { start: 5, end: 14 },
    { start: 18, end: 27 }
  ];

  // 逐一走訪每個區間
  for (var r = 0; r < rowRanges.length; r++) {
    var startRow = rowRanges[r].start;
    var endRow = rowRanges[r].end;
    var numRows = endRow - startRow + 1;
    
    // 從指定的 startRow 開始，讀取欄位 B (第 2 欄) 開始的 4 欄（品名、分類、細項、數量）
    var itemData = inputSheet.getRange(startRow, 2, numRows, 4).getValues();
    
    for (var i = 0; i < itemData.length; i++) {
      var itemName = itemData[i][0];
      var categorySheet = itemData[i][1];
      var detail = itemData[i][2];
      var qty = itemData[i][3];

      // 檢查該列是否有填寫品名且數量大於 0
      if (itemName && qty > 0) {
        validItems.push({ 
          itemName: itemName, 
          categorySheet: categorySheet, 
          detail: detail, 
          qty: qty 
        });
      }
    }
  }

  if (validItems.length === 0) {
    SpreadsheetApp.getUi().alert("請至少輸入一個品名與對應的使用單位量！");
    return;
  }

  // =============================================================
  // 任務一：寫入「耗用量流水」分頁
  // =============================================================

  var flowHeaderRow = 2; 
  var flowHeaders = flowSheet.getRange(flowHeaderRow, 1, 1, flowSheet.getLastColumn()).getValues()[0];

  // 嚴格只看 A 欄（流水號）來找最後一列，從下方往上找，忽略空白或標題列
  var flowAValues = flowSheet.getRange("A:A").getValues();
  var flowLastRow = 0;
  for (var r = flowAValues.length - 1; r >= 0; r--) {
    var val = flowAValues[r][0];
    // 確保抓到的是大於等於標題列號、且有填寫數字或有效流水號的列
    if (r + 1 > flowHeaderRow && val !== "" && !isNaN(val)) {
      flowLastRow = r + 1;
      break;
    }
  }
  
  // 如果找不到任何歷史資料流水號，就從標題列的下一列開始
  if (flowLastRow === 0) {
    flowLastRow = flowHeaderRow; 
  }

  var nextFlowSerial = 1;
  var lastFlowVal = flowSheet.getRange(flowLastRow, 1).getValue();
  if (!isNaN(lastFlowVal) && lastFlowVal !== "" && lastFlowVal !== "流水號") {
    nextFlowSerial = Number(lastFlowVal) + 1;
  }

  var targetFlowWriteRow = flowLastRow + 1;
  
  // 僅寫入基本欄位 (A~D)
  flowSheet.getRange(targetFlowWriteRow, 1).setValue(nextFlowSerial);
  flowSheet.getRange(targetFlowWriteRow, 2).setValue(date);
  flowSheet.getRange(targetFlowWriteRow, 3).setValue(customerName);
  flowSheet.getRange(targetFlowWriteRow, 4).setValue(receiptNo);

  // 僅針對有填寫的品名寫入對應欄位，不整行覆蓋
  validItems.forEach(function(item) {
    var colIdx = flowHeaders.indexOf(item.itemName);
    if (colIdx !== -1) {
      flowSheet.getRange(targetFlowWriteRow, colIdx + 1).setValue(item.qty);
    }
  });

  // =============================================================
  // 任務二：寫入大分類對應分頁（如：肉毒、玻尿酸、SUNMAX）
  // =============================================================
  var groupedByCategory = {};
  validItems.forEach(function(item) {
    if (item.categorySheet && item.categorySheet !== "#N/A") {
      if (!groupedByCategory[item.categorySheet]) {
        groupedByCategory[item.categorySheet] = [];
      }
      groupedByCategory[item.categorySheet].push(item);
    }
  });

  for (var catSheetName in groupedByCategory) {
    var targetSheet = ss.getSheetByName(catSheetName);
    if (!targetSheet) continue;

    var itemsInCat = groupedByCategory[catSheetName];
    var maxCols = targetSheet.getLastColumn();

    var row6 = targetSheet.getRange(6, 1, 1, maxCols).getValues()[0];
    var row7 = targetSheet.getRange(7, 1, 1, maxCols).getValues()[0];

    // 只檢查 A 欄找出真正最後一列
    var colAValues = targetSheet.getRange("A:A").getValues();
    var targetLastRow = 0;
    for (var r = colAValues.length - 1; r >= 0; r--) {
      if (colAValues[r][0] !== "" && !isNaN(colAValues[r][0])) {
        targetLastRow = r + 1;
        break;
      }
    }
    
    if (targetLastRow < 7) {
      targetLastRow = 7; 
    }

    var nextCatSerial = 1;
    var lastCatVal = targetSheet.getRange(targetLastRow, 1).getValue();
    if (!isNaN(lastCatVal) && lastCatVal !== "") {
      nextCatSerial = Number(lastCatVal) + 1;
    }

    var writeRowIndex = targetLastRow + 1;

    // **關鍵修改**：只寫入 A、B、C、D 欄，絕對不整行覆蓋，保護後方的公式！
    targetSheet.getRange(writeRowIndex, 1).setValue(nextCatSerial);
    targetSheet.getRange(writeRowIndex, 2).setValue(date);
    targetSheet.getRange(writeRowIndex, 3).setValue(customerName);
    targetSheet.getRange(writeRowIndex, 4).setValue(receiptNo);
    // 如果是「藥劑」分頁，將 B16 欄位的點滴名稱寫入三聯單號右側的 E 欄 (第 5 欄)
    if (catSheetName === "藥劑") {
      var dripName = inputSheet.getRange("B16").getValue();
      targetSheet.getRange(writeRowIndex, 5).setValue(dripName);
    }

    // 尋找對應細項的欄位並寫入
    itemsInCat.forEach(function(item) {
      var targetDetail = item.detail ? item.detail.toString().trim() : "";
      var matchedDetailCol = -1;

      // 關鍵修改：如果是「藥劑」分頁，改用 row7 搜尋細項；其他分頁維持用 row6
      var searchRow = (catSheetName === "藥劑") ? row7 : row6;

      for (var col = 4; col < searchRow.length; col++) { // 從 E 欄開始找
        var cellVal = searchRow[col] ? searchRow[col].toString().trim() : "";
        if (cellVal !== "" && cellVal.indexOf(targetDetail) !== -1) {
          matchedDetailCol = col;
          break;
        }
      }

      if (matchedDetailCol !== -1) {
        var targetColIndex = -1;
        var finalQty = item.qty; // 預設使用原本的正數數量

        // 判斷該大分類分頁是否為「藥劑」
        if (catSheetName === "藥劑") {
          // 如果是大分類「藥劑」，直接對應到該細項所在的欄位 (第7列)，且數量轉為負數
          targetColIndex = matchedDetailCol;
          finalQty = item.qty * -1;
        } else {
          // 其他分類維持原本邏輯：尋找細項右側的「耗用量」欄位
          var usedColIndex = -1;
          for (var c = matchedDetailCol; c < matchedDetailCol + 3 && c < maxCols; c++) {
            var cellVal7 = row7[c] ? row7[c].toString() : "";
            if (cellVal7.indexOf("耗用量") !== -1) {
              usedColIndex = c;
              break;
            }
          }
          if (usedColIndex === -1) {
            usedColIndex = matchedDetailCol + 1; // 預設中間欄 (耗用量)
          }
          targetColIndex = usedColIndex;
        }

        // 獨立寫入該品名的數量（藥劑為負數，其他為正數），不影響周圍公式
        targetSheet.getRange(writeRowIndex, targetColIndex + 1).setValue(finalQty);
      }
    });
  }

  // =============================================================
  // 清理與完成提示
  // =============================================================
  inputSheet.getRange("B2:B3").clearContent(); 
  inputSheet.getRange("B16").clearContent();
  inputSheet.getRange("B5:B14").clearContent();
  inputSheet.getRange("E5:E14").clearContent();
  inputSheet.getRange("E18:E27").clearContent();

  SpreadsheetApp.getUi().alert("資料已成功寫入！");
}

function submitStockInButtonB() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName("進貨key-in");
  
  if (!inputSheet) {
    SpreadsheetApp.getUi().alert("找不到『進貨key-in』分頁！");
    return;
  }

  // 1. 讀取「進貨key-in」基本資訊
  var date = inputSheet.getRange("B1").getValue();

  if (!date) {
    SpreadsheetApp.getUi().alert("請填寫【日期】！");
    return;
  }

  // 2. 讀取多筆進貨項目資料 (從第 5 列開始：B欄品名, C欄大分類, D欄細項, E欄進貨量, F欄進貨單位)
  var startRow = 5;
  var lastRow = inputSheet.getLastRow();
  
  if (lastRow < startRow) {
    SpreadsheetApp.getUi().alert("未找到進貨項目資料！");
    return;
  }

  // 讀取 B 到 F 欄共 5 個欄位
  var itemData = inputSheet.getRange(startRow, 2, lastRow - startRow + 1, 5).getValues(); 

  // 過濾出有效輸入（有填寫品名且進貨量 > 0）
  var validItems = [];
  for (var i = 0; i < itemData.length; i++) {
    var itemName = itemData[i][0];
    var categorySheet = itemData[i][1];
    var detail = itemData[i][2];      // 細項名稱 (例如：彩虹100U)
    var qty = itemData[i][3];         // 進貨單位量 (E欄)
    var unit = itemData[i][4];        // 進貨單位 (F欄)

    if (itemName && qty > 0) {
      validItems.push({
        itemName: itemName,
        categorySheet: categorySheet,
        detail: detail,
        qty: qty,
        unit: unit
      });
    }
  }

  if (validItems.length === 0) {
    SpreadsheetApp.getUi().alert("請至少輸入一個進貨品名與對應的進貨量！");
    return;
  }

  // =============================================================
  // 任務一：同步寫入「進貨流水」分頁
  // 欄位：[序號]、[日期]、[進貨]、[品名]、[大分類(分頁)]、[細項]、[進貨量]、[進貨單位]
  // =============================================================
  var logSheet = ss.getSheetByName("進貨流水");
  if (logSheet) {
    var logLastRow = logSheet.getLastRow();
    var nextLogSerial = 1;
    
    if (logLastRow >= 1) {
      var lastLogVal = logSheet.getRange(logLastRow, 1).getValue();
      if (!isNaN(lastLogVal) && lastLogVal !== "") {
        nextLogSerial = Number(lastLogVal) + 1;
      }
    }

    // 逐筆寫入進貨流水
    validItems.forEach(function(item, index) {
      var writeLogIndex = logLastRow + 1 + index;
      logSheet.getRange(writeLogIndex, 1).setValue(nextLogSerial + index); // 序號 (A欄)
      logSheet.getRange(writeLogIndex, 2).setValue(date);                 // 日期 (B欄)
      logSheet.getRange(writeLogIndex, 3).setValue("進貨");                 // 進貨固定字樣 (C欄)
      logSheet.getRange(writeLogIndex, 4).setValue(item.itemName);        // 品名 (D欄)
      logSheet.getRange(writeLogIndex, 5).setValue(item.categorySheet);   // 大分類(分頁) (E欄)
      logSheet.getRange(writeLogIndex, 6).setValue(item.detail);          // 細項 (F欄)
      logSheet.getRange(writeLogIndex, 7).setValue(item.qty);             // 進貨量 (G欄)
      logSheet.getRange(writeLogIndex, 8).setValue(item.unit);            // 進貨單位 (H欄)
    });
  }

  // =============================================================
  // 任務二：將進貨項目依「大分類分頁」分組處理
  // =============================================================
  var groupedByCategory = {};
  validItems.forEach(function(item) {
    if (item.categorySheet && item.categorySheet !== "#N/A") {
      if (!groupedByCategory[item.categorySheet]) {
        groupedByCategory[item.categorySheet] = [];
      }
      groupedByCategory[item.categorySheet].push(item);
    }
  });

  for (var catSheetName in groupedByCategory) {
    var targetSheet = ss.getSheetByName(catSheetName);
    if (!targetSheet) {
      Logger.log("找不到對應的大分類分頁：" + catSheetName);
      continue;
    }

    var itemsInCat = groupedByCategory[catSheetName];
    var maxCols = targetSheet.getLastColumn();

    var row6 = targetSheet.getRange(6, 1, 1, maxCols).getValues()[0];
    var row7 = targetSheet.getRange(7, 1, 1, maxCols).getValues()[0];

    // 只檢查 A 欄（序號），從最下方往上找，找出真正有填寫的最後一列
    var colAValues = targetSheet.getRange("A:A").getValues();
    var targetLastRow = 0;
    for (var r = colAValues.length - 1; r >= 0; r--) {
      if (colAValues[r][0] !== "" && !isNaN(colAValues[r][0])) {
        targetLastRow = r + 1;
        break;
      }
    }
    
    if (targetLastRow < 7) {
      targetLastRow = 7; 
    }

    var nextCatSerial = 1;
    var lastCatVal = targetSheet.getRange(targetLastRow, 1).getValue();
    if (!isNaN(lastCatVal) && lastCatVal !== "") {
      nextCatSerial = Number(lastCatVal) + 1;
    }

    var writeRowIndex = targetLastRow + 1;

    // 寫入固定欄位：A欄序號、B欄日期、C欄「進貨」、D欄空白
    targetSheet.getRange(writeRowIndex, 1).setValue(nextCatSerial);
    targetSheet.getRange(writeRowIndex, 2).setValue(date);
    targetSheet.getRange(writeRowIndex, 3).setValue("進貨");
    targetSheet.getRange(writeRowIndex, 4).setValue("");

    // 尋找對應細項的欄位並寫入
    itemsInCat.forEach(function(item) {
      var targetDetail = item.detail ? item.detail.toString().trim() : "";
      var matchedDetailCol = -1;

      var searchRow = (catSheetName === "藥劑") ? row7 : row6;

      for (var col = 4; col < searchRow.length; col++) { // 從 E 欄開始找
        var cellVal = searchRow[col] ? searchRow[col].toString().trim() : "";
        if (cellVal !== "" && cellVal.indexOf(targetDetail) !== -1) {
          matchedDetailCol = col;
          break;
        }
      }

      if (matchedDetailCol !== -1) {
        var stockInColIndex = -1;

        if (catSheetName === "藥劑") {
          stockInColIndex = matchedDetailCol;
        } else {
          for (var c = matchedDetailCol; c < matchedDetailCol + 3 && c < maxCols; c++) {
            var cellVal7 = row7[c] ? row7[c].toString() : "";
            if (cellVal7.indexOf("進貨量") !== -1) {
              stockInColIndex = c;
              break;
            }
          }
          if (stockInColIndex === -1) {
            stockInColIndex = matchedDetailCol; 
          }
        }

        targetSheet.getRange(writeRowIndex, stockInColIndex + 1).setValue(item.qty);
      } else {
        Logger.log("分頁【" + catSheetName + "】找不到對應細項：" + targetDetail);
      }
    });
  }

  // =============================================================
  // 清理與完成提示
  // =============================================================
  inputSheet.getRange("B1").clearContent(); // 清除日期
  for (var r = 0; r < itemData.length; r++) {
    inputSheet.getRange(startRow + r, 2).clearContent(); // 清除品名 (B欄)
    inputSheet.getRange(startRow + r, 5).clearContent(); // 清除進貨量 (E欄)
  }

  SpreadsheetApp.getUi().alert("進貨資料已成功寫入對應分頁與【進貨流水】！");
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('庫存與使用記錄管理系統')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 取得「全產品」分頁中，B欄為 TRUE（有使用）的產品詳細資訊
function getActiveProducts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('全產品');
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var products = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var isUsed = row[1]; // B欄
    
    if (isUsed === true || isUsed === "TRUE" || isUsed === "True") {
      products.push({
        sheetRow: i + 1,       // 記錄在「全產品」分頁的實際列號
        name: row[2],          // C欄: 品名
        category: row[3],      // D欄: 大分類(分頁)
        subCategory: row[4],   // E欄: 細項
        useUnit: row[5],       // F欄: 使用單位
        stockUnit: row[6]      // G欄: 盤點單位
      });
    }
  }
  return products;
}

// 輔助函式：只透過 A～D 欄判斷分頁的最後一列
function getLastRowAtoD(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow === 0) return 1;
  var range = sheet.getRange(1, 1, lastRow, 4).getValues();
  for (var r = range.length - 1; r >= 0; r--) {
    if (range[r][0] !== "" || range[r][1] !== "" || range[r][2] !== "" || range[r][3] !== "") {
      return r + 1; 
    }
  }
  return 1;
}

// 輔助函式：取得 A 欄流水號的下一個數字
function getNextId(sheet, lastRow) {
  if (lastRow <= 1) return 1;
  var lastIdVal = sheet.getRange(lastRow, 1).getValue();
  var parsedId = parseInt(lastIdVal, 10);
  return !isNaN(parsedId) ? parsedId + 1 : lastRow;
}

// 輔助函式：在分類分頁中尋找欄位
function findTargetColumn(sheet, subCategory, targetKeyword) {
  var maxCols = sheet.getLastColumn();
  if (maxCols === 0) return -1;
  
  var row6 = sheet.getRange(6, 1, 1, maxCols).getValues()[0];
  var row7 = sheet.getRange(7, 1, 1, maxCols).getValues()[0];
  var currentSubCat = "";
  
  for (var c = 0; c < maxCols; c++) {
    if (row6[c] !== "" && row6[c] != null) {
      currentSubCat = String(row6[c]).trim();
    }
    if (currentSubCat === String(subCategory).trim()) {
      var r7Val = String(row7[c]).trim();
      if (r7Val === targetKeyword) {
        return c + 1;
      }
    }
  }
  return -1;
}

// 【客戶使用記錄】寫入
function submitUsageData(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var flowSheet = ss.getSheetByName('耗用量流水');
    if (!flowSheet) {
      return { success: false, message: "找不到「耗用量流水」分頁！" };
    }
    
    var date = formData.date;
    var clientName = formData.clientName;
    var orderNo = formData.orderNo;
    var items = formData.items;
    
    if (!items || items.length === 0) {
      return { success: false, message: "沒有接收到任何有效的品項資料！" };
    }
    
    var flowLastRow = getLastRowAtoD(flowSheet);
    var nextId = getNextId(flowSheet, flowLastRow);
    var targetRow = flowLastRow + 1;
    
    flowSheet.getRange(targetRow, 1).setValue(nextId);
    flowSheet.getRange(targetRow, 2).setValue(date);
    flowSheet.getRange(targetRow, 3).setValue(clientName);
    flowSheet.getRange(targetRow, 4).setValue(orderNo);
    
    var activeProducts = getActiveProducts();
    
    items.forEach(function(item) {
      if (item.product && item.qty > 0) {
        var matchedProd = activeProducts.find(function(p) { return p.name === item.product; });
        if (matchedProd) {
          var colIndex = matchedProd.sheetRow + 3;
          flowSheet.getRange(targetRow, colIndex).setValue(item.qty);
        }
      }
    });
    
    var categoryMap = {};
    items.forEach(function(item) {
      var matchedProd = activeProducts.find(function(p) { return p.name === item.product; });
      if (matchedProd && matchedProd.category) {
        var catName = matchedProd.category;
        if (!categoryMap[catName]) {
          categoryMap[catName] = [];
        }
        categoryMap[catName].push({
          subCategory: matchedProd.subCategory,
          qty: item.qty
        });
      }
    });
    
    for (var catName in categoryMap) {
      var catSheet = ss.getSheetByName(catName);
      if (!catSheet) continue;
      
      var catLastRow = getLastRowAtoD(catSheet);
      var catNextId = getNextId(catSheet, catLastRow);
      var catTargetRow = catLastRow + 1;
      
      catSheet.getRange(catTargetRow, 1).setValue(catNextId);
      catSheet.getRange(catTargetRow, 2).setValue(date);
      catSheet.getRange(catTargetRow, 3).setValue(clientName);
      catSheet.getRange(catTargetRow, 4).setValue(orderNo);
      
      var catItems = categoryMap[catName];
      catItems.forEach(function(ci) {
        var col = findTargetColumn(catSheet, ci.subCategory, "耗用量");
        if (col > 0) {
          catSheet.getRange(catTargetRow, col).setValue(ci.qty);
        }
      });
    }
    
    return { success: true, message: "客戶使用記錄送出成功，流水號：" + nextId };
  } catch (error) {
    return { success: false, message: "發生錯誤: " + error.toString() };
  }
}

// 【進貨記錄】寫入
function submitPurchaseData(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var purchaseSheet = ss.getSheetByName('進貨流水');
    if (!purchaseSheet) {
      purchaseSheet = ss.insertSheet('進貨流水');
      purchaseSheet.getRange(1, 1, 1, 8).setValues([['流水號', '進貨日期', '備註', '品名', '大分類(分頁)', '細項', '進貨量', '單位']]);    }
    
    var date = formData.date;
    var items = formData.items;
    
    if (!items || items.length === 0) {
      return { success: false, message: "沒有接收到任何有效的進貨品項！" };
    }
    
    var purchaseLastRow = getLastRowAtoD(purchaseSheet);
    var nextId = getNextId(purchaseSheet, purchaseLastRow);
    
    items.forEach(function(item, index) {
      if (item.product && item.qty > 0) {
        var targetRow = purchaseSheet.getLastRow() + 1;
        var rowId = nextId + index;
        
        purchaseSheet.getRange(targetRow, 1).setValue(rowId);
        purchaseSheet.getRange(targetRow, 2).setValue(date);
        purchaseSheet.getRange(targetRow, 3).setValue("進貨"); // 這裡將 C 欄固定填入「進貨」
        purchaseSheet.getRange(targetRow, 4).setValue(item.product);
        purchaseSheet.getRange(targetRow, 5).setValue(item.category);
        purchaseSheet.getRange(targetRow, 6).setValue(item.subCategory);
        purchaseSheet.getRange(targetRow, 7).setValue(item.qty);
        purchaseSheet.getRange(targetRow, 8).setValue(item.unit);
      }
    });
    
    // 2. 寫入對應的「大分類(分頁)」
    var categoryMap = {};
    items.forEach(function(item) {
      if (item.category) {
        if (!categoryMap[item.category]) {
          categoryMap[item.category] = [];
        }
        categoryMap[item.category].push({
          subCategory: item.subCategory,
          qty: item.qty
        });
      }
    });
    
    for (var catName in categoryMap) {
      var catSheet = ss.getSheetByName(catName);
      if (!catSheet) continue;
      
      var catLastRow = getLastRowAtoD(catSheet);
      var catNextId = getNextId(catSheet, catLastRow);
      var catTargetRow = catLastRow + 1;
      
      catSheet.getRange(catTargetRow, 1).setValue(catNextId);
      catSheet.getRange(catTargetRow, 2).setValue(date);
      catSheet.getRange(catTargetRow, 3).setValue("進貨"); // 同步在分類分頁的 C 欄填入「進貨」
      
      var catItems = categoryMap[catName];
      catItems.forEach(function(ci) {
        var col = findTargetColumn(catSheet, ci.subCategory, "進貨量");
        if (col > 0) {
          catSheet.getRange(catTargetRow, col).setValue(ci.qty);
        }
      });
    }
    
    return { success: true, message: "進貨記錄送出成功！" };
  } catch (error) {
    return { success: false, message: "發生錯誤: " + error.toString() };
  }
}
