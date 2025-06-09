// ==== a single-entry undo stack ====
var DOC_PROPS = PropertiesService.getDocumentProperties();

/**
 * Add the custom menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Scanner')
    .addItem('Open Scanner Sidebar', 'openScannerSidebar')
    .addToUi();
}

/**
 * Show the sidebar.
 */
function openScannerSidebar() {
  var html = HtmlService
    .createHtmlOutputFromFile('ScannerSidebar')
    .setTitle('Parcel Scanner');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * First-time scan handler: marks Dispatched or signals confirmReturn.
 */
function processParcelScan(scannedValue) {
  scannedValue = scannedValue.trim().replace(/\s+/g, '');
  if (!scannedValue) return 'Empty';

  var ss      = SpreadsheetApp.getActiveSpreadsheet(),
      sheet   = ss.getSheetByName("Sheet1"),
      headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0],
      parcelCol  = headers.indexOf("Parcel number")+1,
      statusCol  = headers.indexOf("Shipping Status")+1,
      dateCol    = headers.indexOf("Dispatch Date")+1,
      productCol = headers.indexOf("Product name")+1,
      qtyCol     = headers.indexOf("Quantity")+1,
      amountCol  = headers.indexOf("Amount")+1;

  if (!parcelCol) return 'ParcelColNotFound';

  // find row
  var data = sheet.getDataRange().getValues(),
      foundRow = null;
  for (var i=1; i<data.length; i++) {
    var clean = String(data[i][parcelCol-1]).replace(/\s+/g,'');
    if (clean.toUpperCase() === scannedValue.toUpperCase()) {
      foundRow = i+1; break;
    }
  }
  if (!foundRow) return 'NotFound';

  // read old
  var rowData   = sheet.getRange(foundRow,1,1,sheet.getLastColumn()).getValues()[0],
      oldStatus = statusCol ? rowData[statusCol-1] : '',
      oldDate   = dateCol   ? rowData[dateCol-1] : null;

  // decide
  var actionType, newStatus;
  if (!oldStatus || (oldStatus!=='Dispatched' && oldStatus!=='Returned')) {
    actionType = 'dispatch';
    newStatus  = 'Dispatched';
  } else if (oldStatus==='Dispatched') {
    return 'confirmReturn';
  } else {
    return 'AlreadyReturned';
  }

  // write new
  sheet.getRange(foundRow,statusCol).setValue(newStatus);
  var now = new Date(),
      todayMid = new Date(now.getFullYear(),now.getMonth(),now.getDate());
  sheet.getRange(foundRow,dateCol).setValue(todayMid);

  // parse products & qty
  var products   = String(rowData[productCol-1]).split('\n').map(s=>s.trim()).filter(Boolean),
      quantities = String(rowData[qtyCol-1]).split('\n').map(s=>s.trim()).filter(Boolean),
      orderAmt   = amountCol ? Number(rowData[amountCol-1]||0) : 0;

  // update summaries
  if (actionType==='dispatch') {
    updateDispatchSummaries(products, quantities, orderAmt, todayMid);
  }

  // record undo
  DOC_PROPS.setProperty('lastAction', JSON.stringify({
    type:       actionType,
    row:        foundRow,
    oldStatus:  oldStatus,
    oldDate:    oldDate instanceof Date ? oldDate.toISOString() : null,
    newDate:    todayMid.toISOString(),
    products:   products,
    quantities: quantities,
    amount:     orderAmt
  }));

  return 'Dispatched';
}

/**
 * After confirm, mark Returned.
 */
function processParcelConfirmReturn(scannedValue) {
  scannedValue = scannedValue.trim().replace(/\s+/g,'');
  if (!scannedValue) return 'Empty';

  var ss      = SpreadsheetApp.getActiveSpreadsheet(),
      sheet   = ss.getSheetByName("Sheet1"),
      headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0],
      parcelCol  = headers.indexOf("Parcel number")+1,
      statusCol  = headers.indexOf("Shipping Status")+1,
      dateCol    = headers.indexOf("Dispatch Date")+1,
      productCol = headers.indexOf("Product name")+1,
      qtyCol     = headers.indexOf("Quantity")+1,
      amountCol  = headers.indexOf("Amount")+1;
      orderCol   = headers.indexOf("Order Number")+1;

  if (!parcelCol) return 'ParcelColNotFound';

  // find row
  var data = sheet.getDataRange().getValues(),
      foundRow = null;
  for (var i=1; i<data.length; i++) {
    var clean = String(data[i][parcelCol-1]).replace(/\s+/g,'');
    if (clean.toUpperCase() === scannedValue.toUpperCase()) {
      foundRow = i+1; break;
    }
  }
  if (!foundRow) return 'NotFound';

  // read old
  var rowData   = sheet.getRange(foundRow,1,1,sheet.getLastColumn()).getValues()[0],
      oldStatus = rowData[statusCol-1],
      oldDate   = rowData[dateCol-1];

  // write Returned
  sheet.getRange(foundRow,statusCol).setValue('Returned');
  var now = new Date(),
      todayMid = new Date(now.getFullYear(),now.getMonth(),now.getDate());
  sheet.getRange(foundRow,dateCol).setValue(todayMid);

  // parse
  var products   = String(rowData[productCol-1]).split('\n').map(s=>s.trim()).filter(Boolean),
      quantities = String(rowData[qtyCol-1]).split('\n').map(s=>s.trim()).filter(Boolean),
      orderAmt   = amountCol ? Number(rowData[amountCol-1]||0) : 0;

  // update
  updateReturnSummaries(products, quantities, orderAmt, todayMid);

    // ---- Shopify auto cancel by order number in column B (index 1) ----
  var orderNumber = orderCol ? rowData[orderCol - 1] : '';
  if (orderNumber) shopifyCancelByNumber(orderNumber);



  // record undo
  DOC_PROPS.setProperty('lastAction', JSON.stringify({
    type:       'return',
    row:        foundRow,
    oldStatus:  oldStatus,
    oldDate:    oldDate instanceof Date ? oldDate.toISOString() : null,
    newDate:    todayMid.toISOString(),
    products:   products,
    quantities: quantities,
    amount:     orderAmt
  }));

    // ---- Shopify auto cancel ----
  var orderName = orderCol ? rowData[orderCol - 1] : '';
  var orderId   = findOrderIdByName(orderName);
  if (orderId) {
    var ok = cancelOrderById(orderId);
    Logger.log('Shopify cancel ' + orderName + ' → ' + (ok ? 'OK' : 'FAILED'));
  } else {
    Logger.log('Shopify order not found for ' + orderName);
  }

  return 'Returned';
  
   
}

/**
 * Undo the last scan.
 */
function undoLastScan() {
  var raw = DOC_PROPS.getProperty('lastAction');
  if (!raw) return 'NoAction';
  var act = JSON.parse(raw);

  var ss      = SpreadsheetApp.getActiveSpreadsheet(),
      sheet   = ss.getSheetByName("Sheet1"),
      headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0],
      statusCol = headers.indexOf("Shipping Status")+1,
      dateCol   = headers.indexOf("Dispatch Date")+1;

  // revert Sheet1
  sheet.getRange(act.row, statusCol).setValue(act.oldStatus||'');
  if (act.oldDate) {
    sheet.getRange(act.row, dateCol).setValue(new Date(act.oldDate));
  } else {
    sheet.getRange(act.row, dateCol).clearContent();
  }

  // reverse summaries
  var success = false;
  if (act.type==='dispatch') {
    success = reverseDispatchSummaries(act.products, act.quantities, act.amount, new Date(act.newDate));
  } else {
    success = reverseReturnSummaries(act.products, act.quantities, act.amount, new Date(act.newDate));
  }

  DOC_PROPS.deleteProperty('lastAction');
  return success ? 'Undone' : 'UndoError';
}

/**
 * Fully implemented: add dispatched quantities.
 */
function updateDispatchSummaries(products, quantities, amount, dateObj) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      prodSh  = ss.getSheetByName("Product wise daily dispatch"),
      dailySh = ss.getSheetByName("Daily Dispatch Parcels"),
      today   = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());

  // Product‐wise
  if (prodSh) {
    var data = prodSh.getDataRange().getValues();
    for (var i=0; i<products.length; i++) {
      var name = products[i], qty = Number(quantities[i]||0);
      var found = false;
      for (var r=1; r<data.length; r++) {
        var rowDate = data[r][0], rowProd = data[r][1], rowQty = data[r][2];
        if (rowDate instanceof Date && rowDate.getTime()===today.getTime() && rowProd===name) {
          prodSh.getRange(r+1,3).setValue(Number(rowQty||0)+qty);
          found = true;
          break;
        }
      }
      if (!found) prodSh.appendRow([today, name, qty]);
    }
  }

  // Daily parcels
  if (dailySh) {
    var data2 = dailySh.getDataRange().getValues(), found2 = false;
    for (var k=1; k<data2.length; k++) {
      var d = data2[k][0];
      if (d instanceof Date && d.getTime()===today.getTime()) {
        dailySh.getRange(k+1,2).setValue(Number(data2[k][1]||0)+1);
        dailySh.getRange(k+1,3).setValue(Number(data2[k][2]||0)+amount);
        found2 = true;
        break;
      }
    }
    if (!found2) dailySh.appendRow([today, 1, amount]);
  }
}

/**
 * Fully implemented: add returned quantities.
 */
function updateReturnSummaries(products, quantities, amount, dateObj) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      prodSh  = ss.getSheetByName("Product wise daily return"),
      dailySh = ss.getSheetByName("Daily Return Parcels"),
      today   = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());

  // Product‐wise
  if (prodSh) {
    var data = prodSh.getDataRange().getValues();
    for (var i=0; i<products.length; i++) {
      var name = products[i], qty = Number(quantities[i]||0);
      var found = false;
      for (var r=1; r<data.length; r++) {
        var rowDate = data[r][0], rowProd = data[r][1], rowQty = data[r][2];
        if (rowDate instanceof Date && rowDate.getTime()===today.getTime() && rowProd===name) {
          prodSh.getRange(r+1,3).setValue(Number(rowQty||0)+qty);
          found = true;
          break;
        }
      }
      if (!found) prodSh.appendRow([today, name, qty]);
    }
  }

  // Daily returns
  if (dailySh) {
    var data2 = dailySh.getDataRange().getValues(), found2 = false;
    for (var k=1; k<data2.length; k++) {
      var d = data2[k][0];
      if (d instanceof Date && d.getTime()===today.getTime()) {
        dailySh.getRange(k+1,2).setValue(Number(data2[k][1]||0)+1);
        dailySh.getRange(k+1,3).setValue(Number(data2[k][2]||0)+amount);
        found2 = true;
        break;
      }
    }
    if (!found2) dailySh.appendRow([today, 1, amount]);
  }
}

/**
 * Subtract dispatched quantities and amount from your summary sheets.
 * Matches the date by comparing `toDateString()` so it’s more forgiving
 * of timezones or text‐formatted dates.
 *
 * @param {string[]} products   Array of product names.
 * @param {string[]} quantities Array of quantities corresponding to products.
 * @param {number}   amount     The total order amount.
 * @param {Date}     dateObj    The dispatch date (midnight).
 * @return {boolean}            True if a summary row was updated.
 */
function reverseDispatchSummaries(products, quantities, amount, dateObj) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet(),
      prodSh  = ss.getSheetByName("Product wise daily dispatch"),
      dailySh = ss.getSheetByName("Daily Dispatch Parcels"),
      target  = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
  // 1) Product‐wise sheet
  if (prodSh) {
    var rows = prodSh.getDataRange().getValues();
    for (var r = 1; r < rows.length; r++) {
      var cellDate = rows[r][0] instanceof Date
                     ? rows[r][0]
                     : new Date(rows[r][0]);
      if (cellDate.toDateString() === target.toDateString()) {
        for (var i = 0; i < products.length; i++) {
          if (rows[r][1] === products[i]) {
            var newQty = Number(rows[r][2]||0) - Number(quantities[i]||0);
            if (newQty > 0) {
              prodSh.getRange(r+1, 3).setValue(newQty);
            } else {
              prodSh.deleteRow(r+1);
              r--; // adjust for deleted row
            }
          }
        }
      }
    }
  }
  // 2) Daily parcels sheet
  if (dailySh) {
    var rows2 = dailySh.getDataRange().getValues();
    for (var k = 1; k < rows2.length; k++) {
      var cellDate2 = rows2[k][0] instanceof Date
                      ? rows2[k][0]
                      : new Date(rows2[k][0]);
      if (cellDate2.toDateString() === target.toDateString()) {
        var parcels = Number(rows2[k][1]||0) - 1;
        var amt     = Number(rows2[k][2]||0) - amount;
        if (parcels > 0) {
          dailySh.getRange(k+1, 2).setValue(parcels);
          dailySh.getRange(k+1, 3).setValue(amt);
        } else {
          dailySh.deleteRow(k+1);
        }
        return true;
      }
    }
  }
  return false;
}

/**
 * Subtract returned quantities and amount from your return summary sheets.
 * Uses the same robust date matching as above.
 *
 * @param {string[]} products   Array of product names.
 * @param {string[]} quantities Array of quantities corresponding to products.
 * @param {number}   amount     The total returned amount.
 * @param {Date}     dateObj    The return date (midnight).
 * @return {boolean}            True if a summary row was updated.
 */
function reverseReturnSummaries(products, quantities, amount, dateObj) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet(),
      prodSh  = ss.getSheetByName("Product wise daily return"),
      dailySh = ss.getSheetByName("Daily Return Parcels"),
      target  = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
  // 1) Product‐wise return sheet
  if (prodSh) {
    var rows = prodSh.getDataRange().getValues();
    for (var r = 1; r < rows.length; r++) {
      var cellDate = rows[r][0] instanceof Date
                     ? rows[r][0]
                     : new Date(rows[r][0]);
      if (cellDate.toDateString() === target.toDateString()) {
        for (var i = 0; i < products.length; i++) {
          if (rows[r][1] === products[i]) {
            var newQty = Number(rows[r][2]||0) - Number(quantities[i]||0);
            if (newQty > 0) {
              prodSh.getRange(r+1, 3).setValue(newQty);
            } else {
              prodSh.deleteRow(r+1);
              r--;
            }
          }
        }
      }
    }
  }
  // 2) Daily return parcels sheet
  if (dailySh) {
    var rows2 = dailySh.getDataRange().getValues();
    for (var k = 1; k < rows2.length; k++) {
      var cellDate2 = rows2[k][0] instanceof Date
                      ? rows2[k][0]
                      : new Date(rows2[k][0]);
      if (cellDate2.toDateString() === target.toDateString()) {
        var parcels = Number(rows2[k][1]||0) - 1;
        var amt     = Number(rows2[k][2]||0) - amount;
        if (parcels > 0) {
          dailySh.getRange(k+1, 2).setValue(parcels);
          dailySh.getRange(k+1, 3).setValue(amt);
        } else {
          dailySh.deleteRow(k+1);
        }
        return true;
      }
    }
  }
  return false;
}
/** Cancel an order by its Shopify numeric ID.  Returns true on 200 OK */
function shopifyCancelById(orderId) {
  const props  = PropertiesService.getScriptProperties();
  const token  = props.getProperty('SHOP_TOKEN');
  const domain = props.getProperty('SHOP_DOMAIN');
  if (!token || !domain) {
    Logger.log('Shopify creds missing');   // check Script properties
    return false;
  }

  const url  = 'https://' + domain +
               '/admin/api/2024-01/orders/' + orderId + '/cancel.json';
  const resp = UrlFetchApp.fetch(url, {
    method : 'post',
    muteHttpExceptions: true,
    headers: {
      'X-Shopify-Access-Token': token,
      'Content-Type'          : 'application/json'
    }
  });

  Logger.log('Cancel ' + orderId + ' → ' +
             resp.getResponseCode() + ' : ' +
             resp.getContentText().slice(0,120));
  return resp.getResponseCode() === 200;
}

/** Lookup the numeric ID from an order NAME/# (e.g. "#88768HN18"). */
function shopifyFindOrderId(orderNumber) {
  const props  = PropertiesService.getScriptProperties();
  const token  = props.getProperty('SHOP_TOKEN');
  const domain = props.getProperty('SHOP_DOMAIN');
  if (!token || !domain) return null;

  const name = orderNumber.replace('#','');          // API expects no hash
  const url  = 'https://' + domain +
               '/admin/api/2024-01/orders.json?name=' + encodeURIComponent(name);
  const resp = UrlFetchApp.fetch(url, {
    headers: { 'X-Shopify-Access-Token': token }
  });
  const list = JSON.parse(resp.getContentText()).orders || [];
  return list.length ? list[0].id : null;
}

/** Convenience: cancel directly from order name. */
function shopifyCancelByNumber(orderNumber) {
  const id = shopifyFindOrderId(orderNumber);
  return id ? shopifyCancelById(id) : false;
}

/**
 * Cancel a Shopify order by its numeric ID.
 * Returns true on 200 OK.
 */
function cancelOrderById(orderId) {
  var token  = PropertiesService.getScriptProperties().getProperty('SHOP_TOKEN');
  var domain = PropertiesService.getScriptProperties().getProperty('SHOP_DOMAIN');
  if (!token || !domain) throw new Error('Shopify creds missing in script properties.');

  var url  = 'https://' + domain + '/admin/api/2024-01/orders/' + orderId + '/cancel.json';
  var resp = UrlFetchApp.fetch(url, {
    method : 'post',
    muteHttpExceptions: true,
    headers: { 'X-Shopify-Access-Token': token,
               'Content-Type': 'application/json' }
  });
  return resp.getResponseCode() === 200;
}

/**
 * Look up the numeric order ID from an order name like "#88768HN18".
 * Returns ID or null.
 */
function findOrderIdByName(orderName) {
  var token  = PropertiesService.getScriptProperties().getProperty('SHOP_TOKEN');
  var domain = PropertiesService.getScriptProperties().getProperty('SHOP_DOMAIN');
  var url    = 'https://' + domain +
    '/admin/api/2024-01/orders.json?name=' + encodeURIComponent(orderName.replace('#',''));
  var resp = UrlFetchApp.fetch(url, {
    headers: { 'X-Shopify-Access-Token': token }
  });
  var list = JSON.parse(resp.getContentText()).orders || [];
  return list.length ? list[0].id : null;
}

/**
 * Cancels an order manually (customer request) if not yet dispatched or returned.
 * Updates column G to "Cancelled by Customer" and cancels on Shopify.
 */
function cancelOrderByCustomer(parcelNumberRaw) {
  var parcelNumber = parcelNumberRaw.trim().replace(/\s+/g, '');
  if (!parcelNumber) return 'Empty';

  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getSheetByName("Sheet1");
  var head   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var parcelCol = head.indexOf("Parcel number") + 1;
  var statusCol = head.indexOf("Shipping Status") + 1;
  var dateCol   = head.indexOf("Dispatch Date") + 1;
  var orderCol  = head.indexOf("Order Number") + 1;

  if (!parcelCol || !statusCol || !orderCol) return 'MissingHeaders';

  var data = sheet.getDataRange().getValues();
  var foundRow = -1;
  for (var i = 1; i < data.length; i++) {
    var val = String(data[i][parcelCol - 1]).replace(/\s+/g, '');
    if (val.toUpperCase() === parcelNumber.toUpperCase()) {
      foundRow = i + 1;
      break;
    }
  }
  if (foundRow === -1) return 'NotFound';

  var rowData = data[foundRow - 1];
  var currentStatus = rowData[statusCol - 1];
  if (currentStatus === "Dispatched" || currentStatus === "Returned") {
    return 'TooLate';
  }

  // Set "Cancelled by Customer"
  sheet.getRange(foundRow, statusCol).setValue("Cancelled by Customer");
  var todayMid = new Date(); todayMid.setHours(0, 0, 0, 0);
  sheet.getRange(foundRow, dateCol).setValue(todayMid);

  // Cancel on Shopify
  var orderName = rowData[orderCol - 1];
  var orderId = findOrderIdByName(orderName);
  if (orderId) {
    var ok = cancelOrderById(orderId);
    return ok ? 'Cancelled' : 'ShopifyFail';
  } else {
    return 'OrderNotFoundOnShopify';
  }
}




