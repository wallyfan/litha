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
      purchaseSheet.getRange(1, 1, 1, 7).setValues([['流水號', '進貨日期', '品名', '大分類(分頁)', '細項', '進貨量', '單位']]);
    }
    
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
