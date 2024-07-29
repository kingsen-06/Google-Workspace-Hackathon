function checkInventory() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var productName = data[i][0];
    var stockLevel = data[i][1];
    var minStockLevel = data[i][2];
    var agentEmail = data[i][3];
    
    if (stockLevel < minStockLevel) {
      sendLowStockEmail(productName, stockLevel, minStockLevel, agentEmail);
    }
  }
}

function sendLowStockEmail(productName, stockLevel, minStockLevel, agentEmail) {
  var subject = "Low Stock for " + productName;
  var message = "Dear Purchasing Agent,\n\nOur inventory for " + productName + " is below the minimum threshold.\n\nCurrent stock level: " + stockLevel + "\nMinimum stock level: " + minStockLevel + "\n\nPlease restock as soon as possible.\n\nBest regards,\nYour Company";
  
  MailApp.sendEmail(agentEmail, subject, message);
}

function productOrder() {
  var formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Product Order');
  var formData = formSheet.getDataRange().getValues();

  var inventorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  var inventoryData = inventorySheet.getDataRange().getValues();
  
  for (var i = 1; i < formData.length; i++) {
    var productName = formData[i][1];
    var orderQuantity = parseInt(formData[i][2]);

    for (var inventoryRow = 1; inventoryRow < inventoryData.length; inventoryRow++) {
      if (inventoryData[inventoryRow][0] == productName) {
        var currentStock = parseInt(inventoryData[inventoryRow][1]);
        inventorySheet.getRange(inventoryRow + 1, 2).setValue(currentStock - orderQuantity);
        formSheet.getRange(i + 1, 2).setValue(productName + " (calculated)");

        break;
      }
    }
  }
  checkInventory();
}

function createOnSubmitTrigger() {
  var form = FormApp.openById('16-2v8bSl2hacL7j38DT6AUnYioZXZhl0mhDzGQCDpYs');
  ScriptApp.newTrigger('productOrder')
    .forForm(form)
    .onFormSubmit()
    .create();
}







