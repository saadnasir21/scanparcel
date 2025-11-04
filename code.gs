// ==== a single-entry undo stack ====
var DOC_PROPS = PropertiesService.getDocumentProperties();

// caching for faster parcel lookups
var PARCEL_INDEX_KEY = 'parcelIndex';
// cache parcel lookups for a full day to avoid rebuilding the index
var PARCEL_CACHE_TTL = 24 * 60 * 60; // seconds

// disable row highlighting to speed up large batch scans
var HIGHLIGHT_ROWS = false;

var NEW_INVOICE_SHEET_NAME = 'TCS Invoice (New Format)';
var NEW_INVOICE_HEADERS = [
  'Consignment #',
  'Cust Ref #',
  'Shipper Name',
  'Booking Date',
  'Consignee',
  'Origin',
  'Destination',
  'Weight',
  'Payment Period',
  'Delivery Status',
  'COD Amount',
  'Shipping CHG.'
];
var NEW_INVOICE_HEADER_KEYS = NEW_INVOICE_HEADERS.map(function(h) {
  return h.trim().toLowerCase().replace(/\s+/g, '');
});

function normalizeInvoiceAmount(value) {
  if (typeof value === 'number') return isNaN(value) ? 0 : value;
  if (value === null || value === undefined) return 0;
  var str = String(value).replace(/[^0-9.\-]+/g, '');
  if (!str) return 0;
  var num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

function findHeaderIndex(headers, possibleKeys) {
  if (!Array.isArray(possibleKeys)) possibleKeys = [possibleKeys];
  for (var i = 0; i < possibleKeys.length; i++) {
    var idx = headers.indexOf(possibleKeys[i]);
    if (idx >= 0) return idx;
  }
  return -1;
}

function getParcelIndex(sheet, parcelCol) {
  var cache = CacheService.getDocumentCache();

  // attempt to load a single cached map first
  var raw = cache.get(PARCEL_INDEX_KEY);
  if (raw) return JSON.parse(raw);

  // check for split caches
  var prefixesRaw = cache.get(PARCEL_INDEX_KEY + '_prefixes');
  if (prefixesRaw) {
    try {
      var prefixes = JSON.parse(prefixesRaw);
      var merged = {};
      for (var i = 0; i < prefixes.length; i++) {
        var part = cache.get(PARCEL_INDEX_KEY + '_' + prefixes[i]);
        if (part) {
          var obj = JSON.parse(part);
          for (var k in obj) merged[k] = obj[k];
        }
      }
      if (Object.keys(merged).length) return merged;
    } catch (err) {
      Logger.log('Parcel index cache read failed: ' + err);
    }
  }

  var last = sheet.getLastRow();
  var values = sheet.getRange(2, parcelCol, Math.max(last - 1, 0), 1).getValues();
  var map = {};
  for (var i = 0; i < values.length; i++) {
    var key = String(values[i][0]).replace(/\s+/g, '').toUpperCase();
    if (key) map[key] = i + 2; // adjust for header row
  }

  var json = JSON.stringify(map);
  var MAX_BYTES = 90 * 1024; // ~90KB

  try {
    if (json.length <= MAX_BYTES) {
      cache.put(PARCEL_INDEX_KEY, json, PARCEL_CACHE_TTL);
    } else {
      var groups = {};
      for (var key in map) {
        var prefix = key.charAt(0);
        if (!groups[prefix]) groups[prefix] = {};
        groups[prefix][key] = map[key];
      }
      var stored = [];
      for (var p in groups) {
        var subJson = JSON.stringify(groups[p]);
        if (subJson.length <= MAX_BYTES) {
          cache.put(PARCEL_INDEX_KEY + '_' + p, subJson, PARCEL_CACHE_TTL);
          stored.push(p);
        } else {
          Logger.log('Parcel index segment too large for prefix ' + p + ': ' + subJson.length);
        }
      }
      if (stored.length) {
        cache.put(PARCEL_INDEX_KEY + '_prefixes', JSON.stringify(stored), PARCEL_CACHE_TTL);
      } else {
        Logger.log('Parcel index too large to cache; skipping. Size: ' + json.length);
      }
    }
  } catch (err) {
    Logger.log('Parcel index cache write failed: ' + err);
  }

  return map;
}

function invalidateParcelIndex() {
  var cache = CacheService.getDocumentCache();
  cache.remove(PARCEL_INDEX_KEY);
  var prefixesRaw = cache.get(PARCEL_INDEX_KEY + '_prefixes');
  if (prefixesRaw) {
    try {
      var prefixes = JSON.parse(prefixesRaw);
      for (var i = 0; i < prefixes.length; i++) {
        cache.remove(PARCEL_INDEX_KEY + '_' + prefixes[i]);
      }
    } catch (err) {
      Logger.log('Parcel index invalidation failed: ' + err);
    }
    cache.remove(PARCEL_INDEX_KEY + '_prefixes');
  }
}

/**
 * Add the custom menu.
 */
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureNewInvoiceSheet(ss);

  SpreadsheetApp.getUi()
    .createMenu('Scanner')
    .addItem('Open Scanner Sidebar', 'openScannerSidebar')
    .addItem('Reconcile COD Payments', 'reconcileCODPayments')
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Dispatch Summary')
        .addItem('Last 5 Days', 'showDispatchSummaryLast5')
        .addItem('Last Week', 'showDispatchSummaryWeek')
        .addItem('Last Month', 'showDispatchSummaryMonth'))
      .addToUi();
  }

/**
 * Show the sidebar.
 */
function openScannerSidebar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sheet1');
  if (sheet) {
    var head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var parcelCol = head.indexOf('Parcel number') + 1;
    if (parcelCol) {
      invalidateParcelIndex();
      getParcelIndex(sheet, parcelCol);
    }
  }
  var html = HtmlService
    .createHtmlOutputFromFile('ScannerSidebar')
    .setTitle('Parcel Scanner');
  SpreadsheetApp.getUi().showSidebar(html);
}

function ensureNewInvoiceSheet(ss) {
  var sheet = ss.getSheetByName(NEW_INVOICE_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(NEW_INVOICE_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, NEW_INVOICE_HEADERS.length).setValues([NEW_INVOICE_HEADERS]);
  }
  return sheet;
}

/**
 * Highlight and navigate to the given row on the active sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to activate.
 * @param {number} row Row number to highlight.
 */
function highlightRow(sheet, row) {
  if (!HIGHLIGHT_ROWS) return;
  sheet.setActiveRange(sheet.getRange(row, 1, 1, sheet.getLastColumn()));
  SpreadsheetApp.flush();
}

/**
 * Automatically handle manual edits to the Shipping Status column.
 * Adds or reverses inventory adjustments and sets dispatch dates.
 * Supported statuses:
 *   - "Dispatch through Local Rider" → adjust inventory only
 *   - "Dispatch through Bykea"      → normal dispatch summary
 *   - "Returned"                    → return adjustment
 * Clearing the status will reverse the previous adjustment.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The edit event.
 */
function onEdit(e) {
  var range = e.range;
  if (!range) return;
  var sheet = range.getSheet();
  if (!sheet || sheet.getName() !== 'Sheet1') return;
  invalidateParcelIndex();

  var head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var statusCol  = head.indexOf('Shipping Status') + 1;
  var dateCol    = head.indexOf('Dispatch Date') + 1;
  var productCol = head.indexOf('Product name') + 1;
  var qtyCol     = head.indexOf('Quantity') + 1;
  var amountCol  = head.indexOf('Amount') + 1;
  var orderCol   = head.indexOf('Order Number') + 1;

  if (!statusCol || !dateCol) return;
  if (range.getColumn() !== statusCol || range.getRow() === 1) return;
  if (range.getNumRows() > 1 || range.getNumColumns() > 1) return;

  var newStatus = String(range.getValue() || '').trim();
  var oldStatus = e.oldValue ? String(e.oldValue).trim() : '';
  var row = range.getRow();
  var dateCell = sheet.getRange(row, dateCol);
  var dateVal = dateCell.getValue();
  var oldDateObj = dateVal instanceof Date ? dateVal : (dateVal ? new Date(dateVal) : null);
  var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  var products   = productCol ? String(rowData[productCol - 1]).split('\n').map(function(s){return s.trim();}).filter(Boolean) : [];
  var quantities = qtyCol ? String(rowData[qtyCol - 1]).split('\n').map(function(s){return s.trim();}).filter(Boolean) : [];
  var orderAmt   = amountCol ? Number(rowData[amountCol - 1] || 0) : 0;

  var todayMid = new Date();
  todayMid.setHours(0, 0, 0, 0);

  function reverseOld() {
    if (!oldStatus) return;
    if (oldStatus === 'Dispatched' || oldStatus === 'Dispatch through Bykea') {
      reverseDispatchSummaries(products, quantities, orderAmt, oldDateObj || todayMid);
    } else if (oldStatus === 'Returned') {
      reverseReturnSummaries(products, quantities, orderAmt, oldDateObj || todayMid);
    } else if (oldStatus === 'Dispatch through Local Rider') {
      reverseDispatchInventoryOnly(products, quantities, orderAmt, oldDateObj || todayMid, rowData[orderCol - 1]);
    }
  }

  var statusLower = newStatus.toLowerCase();

  if (!newStatus) {
    reverseOld();
    dateCell.clearContent();
    return;
  }

  reverseOld();
  if (statusLower === 'dispatch through local rider') {
    dateCell.setValue(todayMid);
    updateDispatchInventoryOnly(products, quantities, orderAmt, todayMid, rowData[orderCol - 1]);
  } else if (statusLower === 'dispatch through bykea' || statusLower === 'dispatched') {
    dateCell.setValue(todayMid);
    updateDispatchSummaries(products, quantities, orderAmt, todayMid);
  } else if (statusLower === 'returned') {
    dateCell.setValue(todayMid);
    if (oldStatus === 'Dispatch through Local Rider') {
      updateReturnInventoryOnly(products, quantities, orderAmt, todayMid);
    } else {
      updateReturnSummaries(products, quantities, orderAmt, todayMid);
    }
    var orderNumber = orderCol ? String(rowData[orderCol - 1] || '').trim() : '';
    restockShopifyOrder(orderNumber, oldStatus, orderAmt);
  }
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

    var index = getParcelIndex(sheet, parcelCol);
    var foundRow = index[scannedValue.toUpperCase()] || null;
    if (!foundRow) {
      invalidateParcelIndex();
      index = getParcelIndex(sheet, parcelCol);
      foundRow = index[scannedValue.toUpperCase()] || null;
    }
    if (!foundRow) return 'NotFound';

    // read old
    var rowData   = sheet.getRange(foundRow,1,1,sheet.getLastColumn()).getValues()[0],
        oldStatus = statusCol ? rowData[statusCol-1] : '',
        oldDate   = dateCol   ? rowData[dateCol-1] : null;

  // prevent dispatching an order cancelled by the customer
  if (String(oldStatus).trim() === 'Cancelled by Customer') {
    return 'WasCancelled';
  }

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
  highlightRow(sheet, foundRow);
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

  var index = getParcelIndex(sheet, parcelCol);
  var foundRow = index[scannedValue.toUpperCase()] || null;
  if (!foundRow) {
    invalidateParcelIndex();
    index = getParcelIndex(sheet, parcelCol);
    foundRow = index[scannedValue.toUpperCase()] || null;
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

  // ---- Shopify auto cancel using the order column ----
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
  highlightRow(sheet, foundRow);
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
        if (!(rowDate instanceof Date)) rowDate = new Date(rowDate);
        if (rowDate.toDateString() === today.toDateString() && rowProd===name) {
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
        if (!(rowDate instanceof Date)) rowDate = new Date(rowDate);
        if (rowDate.toDateString() === today.toDateString() && rowProd===name) {
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
 * Add dispatched quantities to inventory only (no parcel summary).
 */
function updateDispatchInventoryOnly(products, quantities, amount, dateObj, orderNum) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet(),
      prodSh = ss.getSheetByName('Product wise daily dispatch'),
      today  = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
  if (prodSh) {
    var data = prodSh.getDataRange().getValues();
    for (var i = 0; i < products.length; i++) {
      var name = products[i], qty = Number(quantities[i] || 0);
      var found = false;
      for (var r = 1; r < data.length; r++) {
        var rowDate = data[r][0];
        if (!(rowDate instanceof Date)) rowDate = new Date(rowDate);
        if (rowDate.toDateString() === today.toDateString() && data[r][1] === name) {
          prodSh.getRange(r + 1, 3).setValue(Number(data[r][2] || 0) + qty);
          found = true;
          break;
        }
      }
      if (!found) prodSh.appendRow([today, name, qty]);
    }
  }
  if (orderNum) recordLocalRiderOrder(orderNum, amount, today);
}

/**
 * Reverse inventory-only dispatch quantities.
 */
function reverseDispatchInventoryOnly(products, quantities, amount, dateObj, orderNum) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet(),
      prodSh = ss.getSheetByName('Product wise daily dispatch'),
      target = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
  if (prodSh) {
    var rows = prodSh.getDataRange().getValues();
    for (var r = 1; r < rows.length; r++) {
      var cellDate = rows[r][0] instanceof Date ? rows[r][0] : new Date(rows[r][0]);
      if (cellDate.toDateString() === target.toDateString()) {
        for (var i = 0; i < products.length; i++) {
          if (rows[r][1] === products[i]) {
            var newQty = Number(rows[r][2] || 0) - Number(quantities[i] || 0);
            if (newQty > 0) {
              prodSh.getRange(r + 1, 3).setValue(newQty);
            } else {
              prodSh.deleteRow(r + 1);
              r--;
            }
          }
        }
      }
    }
  }
  if (orderNum) removeLocalRiderOrder(orderNum, target);
}

/**
 * Record an order dispatched through local rider.
 * Appends a row to the "Local Rider Orders" sheet with date, order ID and amount.
 */
function recordLocalRiderOrder(orderNum, amount, dateObj) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Local Rider Orders');
  if (!sheet) sheet = ss.insertSheet('Local Rider Orders');
  var day = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
  sheet.appendRow([day, orderNum, amount]);
}

/**
 * Remove a previously recorded local rider order entry.
 */
function removeLocalRiderOrder(orderNum, dateObj) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Local Rider Orders');
  if (!sheet) return;
  var rows = sheet.getDataRange().getValues();
  for (var r = 1; r < rows.length; r++) {
    var d = rows[r][0] instanceof Date ? rows[r][0] : new Date(rows[r][0]);
    if (d.toDateString() === dateObj.toDateString() && String(rows[r][1]).trim() === String(orderNum).trim()) {
      sheet.deleteRow(r + 1);
      break;
    }
  }
}

/**
 * Add returned quantities back to inventory only.
 */
function updateReturnInventoryOnly(products, quantities, amount, dateObj) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet(),
      prodSh = ss.getSheetByName('Product wise daily return'),
      today  = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
  if (prodSh) {
    var data = prodSh.getDataRange().getValues();
    for (var i = 0; i < products.length; i++) {
      var name = products[i], qty = Number(quantities[i] || 0);
      var found = false;
      for (var r = 1; r < data.length; r++) {
        var rowDate = data[r][0];
        if (!(rowDate instanceof Date)) rowDate = new Date(rowDate);
        if (rowDate.toDateString() === today.toDateString() && data[r][1] === name) {
          prodSh.getRange(r + 1, 3).setValue(Number(data[r][2] || 0) + qty);
          found = true;
          break;
        }
      }
      if (!found) prodSh.appendRow([today, name, qty]);
    }
  }
}

/**
 * Reverse inventory-only return quantities.
 */
function reverseReturnInventoryOnly(products, quantities, amount, dateObj) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet(),
      prodSh = ss.getSheetByName('Product wise daily return'),
      target = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
  if (prodSh) {
    var rows = prodSh.getDataRange().getValues();
    for (var r = 1; r < rows.length; r++) {
      var cellDate = rows[r][0] instanceof Date ? rows[r][0] : new Date(rows[r][0]);
      if (cellDate.toDateString() === target.toDateString()) {
        for (var i = 0; i < products.length; i++) {
          if (rows[r][1] === products[i]) {
            var newQty = Number(rows[r][2] || 0) - Number(quantities[i] || 0);
            if (newQty > 0) {
              prodSh.getRange(r + 1, 3).setValue(newQty);
            } else {
              prodSh.deleteRow(r + 1);
              r--;
            }
          }
        }
      }
    }
  }
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
    },
    payload: JSON.stringify({ restock: true })
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
 * Cancel the Shopify order for a returned parcel when inventory should restock.
 * Skips orders already marked returned or cancelled earlier.
 */
function restockShopifyOrder(orderNumber, previousStatus, orderAmount) {
  if (!orderNumber) return { status: 'missing' };
  var prev = String(previousStatus || '').trim().toLowerCase();
  if (prev === 'returned' || prev === 'cancelled by customer') {
    return { status: 'skipped' };
  }

  var orderId = shopifyFindOrderId(orderNumber);
  if (!orderId) {
    Logger.log('Shopify order not found for ' + orderNumber);
    return { status: 'notFound' };
  }

  var refunded = shopifyRefundReturnById(orderId, orderAmount);
  if (refunded) {
    Logger.log('Shopify refund succeeded for order ' + orderNumber);
    return { status: 'refunded' };
  }

  var ok = shopifyCancelById(orderId);
  if (!ok) {
    Logger.log('Shopify restock failed for order ' + orderNumber);
    return { status: 'failed' };
  }

  return { status: 'cancelled' };
}

function shopifyRefundReturnById(orderId, orderAmount) {
  try {
    var order = shopifyGetOrder(orderId);
    if (!order) return false;

    var lineItems = order.line_items || [];
    if (!lineItems.length) return false;

    var defaultLocationId = order.location_id || null;
    var refundLineItems = [];
    for (var i = 0; i < lineItems.length; i++) {
      var item = lineItems[i];
      var locId = item.location_id || (item.origin_location && item.origin_location.id) || defaultLocationId;
      var refundLine = {
        line_item_id: item.id,
        quantity: item.quantity,
        restock_type: locId ? 'return' : 'no_restock'
      };
      if (locId) refundLine.location_id = locId;
      refundLineItems.push(refundLine);
    }

    if (!refundLineItems.length) return false;

    var amountNumber = Number(orderAmount || 0);
    if (!amountNumber && order.total_price) amountNumber = Number(order.total_price);
    if (!amountNumber && order.current_total_price) amountNumber = Number(order.current_total_price);
    if (!amountNumber && order.total_price_set && order.total_price_set.shop_money) {
      amountNumber = Number(order.total_price_set.shop_money.amount);
    }
    if (!amountNumber) return false;
    var amountStr = amountNumber.toFixed(2);

    var transactions = shopifyGetOrderTransactions(orderId);
    if (!transactions || !transactions.length) return false;
    var parentTx = null;
    for (var t = 0; t < transactions.length; t++) {
      var kind = String(transactions[t].kind || '').toLowerCase();
      if (kind === 'sale' || kind === 'capture') {
        parentTx = transactions[t];
        break;
      }
      if (!parentTx && (kind === 'authorization' || kind === 'pending')) {
        parentTx = transactions[t];
      }
    }
    if (!parentTx) parentTx = transactions[0];
    if (!parentTx || !parentTx.id) return false;

    var refundKind = String(parentTx.kind || '').toLowerCase() === 'authorization' ? 'void' : 'refund';
    var payload = {
      refund: {
        notify: false,
        note: 'Auto return recorded from sheet',
        shipping: { full_refund: true },
        refund_line_items: refundLineItems,
        transactions: [
          {
            parent_id: parentTx.id,
            amount: amountStr,
            kind: refundKind,
            gateway: parentTx.gateway
          }
        ]
      }
    };

    var props = PropertiesService.getScriptProperties();
    var token = props.getProperty('SHOP_TOKEN');
    var domain = props.getProperty('SHOP_DOMAIN');
    if (!token || !domain) return false;

    var url = 'https://' + domain + '/admin/api/2024-01/orders/' + orderId + '/refunds.json';
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      muteHttpExceptions: true,
      headers: {
        'X-Shopify-Access-Token': token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload)
    });

    var code = resp.getResponseCode();
    if (code >= 200 && code < 300) {
      return true;
    }

    Logger.log('Shopify refund failed for order ' + orderId + ': ' + code + ' ' + resp.getContentText().slice(0, 120));
  } catch (err) {
    Logger.log('Shopify refund exception for order ' + orderId + ': ' + err);
  }
  return false;
}

function shopifyGetOrder(orderId) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('SHOP_TOKEN');
  var domain = props.getProperty('SHOP_DOMAIN');
  if (!token || !domain) return null;

  var url = 'https://' + domain + '/admin/api/2024-01/orders/' + orderId + '.json';
  var resp = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: { 'X-Shopify-Access-Token': token }
  });
  if (resp.getResponseCode() !== 200) return null;
  var data = JSON.parse(resp.getContentText());
  return data.order || null;
}

function shopifyGetOrderTransactions(orderId) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('SHOP_TOKEN');
  var domain = props.getProperty('SHOP_DOMAIN');
  if (!token || !domain) return [];

  var url = 'https://' + domain + '/admin/api/2024-01/orders/' + orderId + '/transactions.json';
  var resp = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: { 'X-Shopify-Access-Token': token }
  });
  if (resp.getResponseCode() !== 200) return [];
  var data = JSON.parse(resp.getContentText());
  return data.transactions || [];
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
               'Content-Type': 'application/json' },
    payload: JSON.stringify({ restock: true })
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
  var amountCol = head.indexOf("Amount") + 1;

  if (!parcelCol || !statusCol || !orderCol) return 'MissingHeaders';

  var index = getParcelIndex(sheet, parcelCol);
  var foundRow = index[parcelNumber.toUpperCase()] || null;
  if (!foundRow) {
    invalidateParcelIndex();
    index = getParcelIndex(sheet, parcelCol);
    foundRow = index[parcelNumber.toUpperCase()] || null;
  }
  if (!foundRow) return 'NotFound';
  var data = sheet.getDataRange().getValues();

  var rowData = data[foundRow - 1];
  var currentStatus = rowData[statusCol - 1];
  if (currentStatus === "Dispatched" || currentStatus === "Returned") {
    return 'TooLate';
  }

  // Set "Cancelled by Customer"
  sheet.getRange(foundRow, statusCol).setValue("Cancelled by Customer");
  var todayMid = new Date(); todayMid.setHours(0, 0, 0, 0);
  sheet.getRange(foundRow, dateCol).setValue(todayMid);

  // Sync to Shopify (refund/cancel so net sales reflect a return)
  var orderName   = rowData[orderCol - 1];
  var orderAmount = amountCol ? Number(rowData[amountCol - 1] || 0) : 0;
  var syncResult  = restockShopifyOrder(String(orderName || ''), currentStatus, orderAmount);

  if (!orderName || !syncResult || syncResult.status === 'missing' || syncResult.status === 'notFound') {
    return 'OrderNotFoundOnShopify';
  }
  if (syncResult.status === 'failed') {
    return 'ShopifyFail';
  }
  return 'Cancelled';
}

/**
 * Cancel by entering the order number instead of the parcel.
 * The input can be the full order string or just the numeric portion.
 */
function cancelOrderByNumber(orderNumRaw) {
  var orderNum = String(orderNumRaw).trim().replace(/^#/, '');
  if (!orderNum) return 'Empty';

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sheet1");
  var head  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var orderCol  = head.indexOf("Order Number") + 1;
  var statusCol = head.indexOf("Shipping Status") + 1;
  var dateCol   = head.indexOf("Dispatch Date") + 1;
  var amountCol = head.indexOf("Amount") + 1;

  if (!orderCol || !statusCol || !dateCol) return 'MissingHeaders';

  var data     = sheet.getDataRange().getValues();
  var foundRow = -1;
  for (var r = 1; r < data.length; r++) {
    var val = String(data[r][orderCol - 1]).trim().replace(/^#/, '');
    if (val.toUpperCase().indexOf(orderNum.toUpperCase()) === 0) {
      foundRow = r + 1;
      break;
    }
  }
  if (foundRow === -1) return 'NotFound';

  var rowData = data[foundRow - 1];
  var status  = rowData[statusCol - 1];
  if (status === "Dispatched" || status === "Returned") {
    return 'TooLate';
  }

  sheet.getRange(foundRow, statusCol).setValue("Cancelled by Customer");
  var todayMid = new Date(); todayMid.setHours(0, 0, 0, 0);
  sheet.getRange(foundRow, dateCol).setValue(todayMid);

  var orderName   = rowData[orderCol - 1];
  var orderAmount = amountCol ? Number(rowData[amountCol - 1] || 0) : 0;
  var syncResult  = restockShopifyOrder(String(orderName || ''), status, orderAmount);

  if (!orderName || !syncResult || syncResult.status === 'missing' || syncResult.status === 'notFound') {
    return 'OrderNotFoundOnShopify';
  }
  if (syncResult.status === 'failed') {
    return 'ShopifyFail';
  }
  return 'Cancelled';
}




/**
 * Manually set the shipping status and optional date for a parcel.
 * @param {string} parcelRaw Parcel number.
 * @param {string} newStatus Status text to set.
 * @param {string} dateStr   Optional date string YYYY-MM-DD.
 * @return {string} result code.
 */
function manualSetStatus(parcelRaw, newStatus, dateStr) {
  var parcel = String(parcelRaw).trim().replace(/\s+/g, '');
  if (!parcel) return 'Empty';

  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getSheetByName("Sheet1");
  var head   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var parcelCol  = head.indexOf("Parcel number") + 1;
  var statusCol  = head.indexOf("Shipping Status") + 1;
  var dateCol    = head.indexOf("Dispatch Date") + 1;
  var productCol = head.indexOf("Product name") + 1;
  var qtyCol     = head.indexOf("Quantity") + 1;
  var amountCol  = head.indexOf("Amount") + 1;
  var orderCol   = head.indexOf("Order Number") + 1;

  if (!parcelCol || !statusCol || !dateCol) return 'MissingHeaders';

  var index = getParcelIndex(sheet, parcelCol);
  var foundRow = index[parcel.toUpperCase()] || null;
  if (!foundRow) {
    invalidateParcelIndex();
    index = getParcelIndex(sheet, parcelCol);
    foundRow = index[parcel.toUpperCase()] || null;
  }
  if (!foundRow) return 'NotFound';
  var data = sheet.getDataRange().getValues();

  var rowData   = data[foundRow - 1];
  var oldStatus = rowData[statusCol - 1];
  var oldStatusTrim = String(oldStatus || '').trim();
  var oldDate   = rowData[dateCol - 1];
  var orderNum  = orderCol ? rowData[orderCol - 1] : '';
  var newStatusLower = String(newStatus || '').trim().toLowerCase();

  if (oldStatusTrim === 'Cancelled by Customer' && newStatus === 'Dispatched') {
    return 'WasCancelled';
  }

  var dateObj = dateStr ? new Date(dateStr) : new Date();
  dateObj.setHours(0, 0, 0, 0);

  var products   = productCol ? String(rowData[productCol - 1]).split('\n').map(function(s){return s.trim();}).filter(Boolean) : [];
  var quantities = qtyCol ? String(rowData[qtyCol - 1]).split('\n').map(function(s){return s.trim();}).filter(Boolean) : [];
  var orderAmt   = amountCol ? Number(rowData[amountCol - 1] || 0) : 0;

  var oldDateObj = oldDate ? (oldDate instanceof Date ? oldDate : new Date(oldDate)) : null;

  // remove previous summary data if needed
  if (oldDateObj) {
    if (oldStatus === 'Dispatched') {
      reverseDispatchSummaries(products, quantities, orderAmt, oldDateObj);
    } else if (oldStatus === 'Returned') {
      reverseReturnSummaries(products, quantities, orderAmt, oldDateObj);
    } else if (oldStatus === 'Dispatch through Local Rider') {
      reverseDispatchInventoryOnly(products, quantities, orderAmt, oldDateObj, orderNum);
    }
  }

  // write new status/date
  sheet.getRange(foundRow, statusCol).setValue(newStatus);
  sheet.getRange(foundRow, dateCol).setValue(dateObj);

  // add new summary data if needed
  if (newStatus === 'Dispatched' || newStatus === 'Dispatch through Bykea') {
    updateDispatchSummaries(products, quantities, orderAmt, dateObj);
  } else if (newStatus === 'Dispatch through Local Rider') {
    updateDispatchInventoryOnly(products, quantities, orderAmt, dateObj, orderNum);
  } else if (newStatus === 'Returned') {
    updateReturnSummaries(products, quantities, orderAmt, dateObj);
  }

  if (newStatusLower === 'returned') {
    var orderNumber = String(orderNum || '').trim();
    restockShopifyOrder(orderNumber, oldStatusTrim, orderAmt);
  }

  return 'Updated';
}


/**
 * Generate a summary sheet of dispatched products for the last `days` days.
 * Results are written to a sheet named "Dispatch Summary".
 * @param {number} days Number of days to include.
 */
function updateDispatchSummarySheet(days) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheetByName('Product wise daily dispatch');
  if (!source) return 'MissingSheet';
  var out = ss.getSheetByName('Dispatch Summary');
  if (!out) out = ss.insertSheet('Dispatch Summary');
  out.clearContents();
  out.appendRow(['Product name', 'Quantity']);

  var today = new Date();
  today.setHours(0,0,0,0);
  var start = new Date(today.getTime() - (days-1)*24*60*60*1000);

  var rows = source.getDataRange().getValues();
  var totals = {};
  for (var i=1; i<rows.length; i++) {
    var d = rows[i][0];
    var prod = rows[i][1];
    var qty = Number(rows[i][2]||0);
    if (!(d instanceof Date)) d = new Date(d);
    if (d >= start && d <= today) {
      totals[prod] = (totals[prod]||0) + qty;
    }
  }
  var keys = Object.keys(totals).sort();
  for (var j=0; j<keys.length; j++) {
    out.appendRow([keys[j], totals[keys[j]]]);
  }
  return 'Updated';
}

function showDispatchSummaryLast5()  { updateDispatchSummarySheet(5); }
function showDispatchSummaryWeek()   { updateDispatchSummarySheet(7); }
function showDispatchSummaryMonth()  { updateDispatchSummarySheet(30); }

function updateDispatchSummaryRange(start, end) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheetByName('Product wise daily dispatch');
  if (!source) return 'MissingSheet';
  var out = ss.getSheetByName('Dispatch Summary');
  if (!out) out = ss.insertSheet('Dispatch Summary');
  out.clearContents();
  out.appendRow(['Product name', 'Quantity']);

  var startDate = new Date(start); startDate.setHours(0,0,0,0);
  var endDate   = new Date(end);   endDate.setHours(0,0,0,0);

  var rows = source.getDataRange().getValues();
  var totals = {};
  for (var i=1; i<rows.length; i++) {
    var d = rows[i][0];
    var prod = rows[i][1];
    var qty = Number(rows[i][2]||0);
    if (!(d instanceof Date)) d = new Date(d);
    if (d >= startDate && d <= endDate) {
      totals[prod] = (totals[prod]||0) + qty;
    }
  }
  var keys = Object.keys(totals).sort();
  for (var j=0; j<keys.length; j++) {
    out.appendRow([keys[j], totals[keys[j]]]);
  }
  return 'Updated';
}

/**
 * Show a dialog to upload the COD invoice CSV.
 */
function openCodUploadDialog() {
  var html = HtmlService.createHtmlOutputFromFile('CodUploadDialog')
    .setWidth(300).setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload COD Invoice');
}

/**
 * Receive uploaded invoice file (base64 or blob), store it and reconcile.
 *
 * @param {string|Blob|Object} fileData Base64 string, blob or {data,name}.
 * @return {string} Confirmation message.
 */
function uploadCodInvoice(fileData) {
  var blob;
  if (fileData instanceof Blob) {
    blob = fileData;
  } else if (fileData && typeof fileData === 'object' && fileData.data) {
    var base = String(fileData.data).replace(/^data:.*;base64,/, '');
    var type = fileData.type || 'application/octet-stream';
    blob = Utilities.newBlob(Utilities.base64Decode(base), type, fileData.name || 'upload');
  } else if (typeof fileData === 'string') {
    var baseStr = fileData.replace(/^data:.*;base64,/, '');
    blob = Utilities.newBlob(Utilities.base64Decode(baseStr));
  } else {
    throw new Error('Unsupported file');
  }

  var extMatch = blob.getName().match(/\.([^.]+)$/);
  var ext = extMatch ? extMatch[1].toLowerCase() : 'csv';
  var data;

  if (ext === 'csv') {
    data = Utilities.parseCsv(blob.getDataAsString());
  } else if (ext === 'xls' || ext === 'xlsx') {
    var tmp = Drive.Files.insert({title: blob.getName(), mimeType: blob.getContentType()}, blob, {convert: true});
    try {
      var tempSs = SpreadsheetApp.openById(tmp.id);
      data = tempSs.getSheets()[0].getDataRange().getValues();
    } finally {
      Drive.Files.remove(tmp.id);
    }
  } else {
    throw new Error('Unsupported file extension: ' + ext);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureNewInvoiceSheet(ss);

  var headerKeys = data[0].map(function(h) {
    return String(h).trim().toLowerCase().replace(/\s+/g, '');
  });
  var targetSheetName = headerKeys.join('\u0001') === NEW_INVOICE_HEADER_KEYS.join('\u0001')
    ? NEW_INVOICE_SHEET_NAME
    : 'TCS Invoice';

  var sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) sheet = ss.insertSheet(targetSheetName);

  var hasHeader = sheet.getLastRow() > 0;
  if (hasHeader) {
    if (data.length > 1) {
      sheet.getRange(sheet.getLastRow() + 1, 1, data.length - 1, data[0].length)
           .setValues(data.slice(1));
    }
  } else {
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }
  reconcileCODPayments();
  return 'Invoice uploaded and reconciled.';
}

/**
 * Reconcile COD payments from invoice data and mark orders.
 */
function reconcileCODPayments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName('Sheet1');
  if (!orderSheet) return;

  ensureNewInvoiceSheet(ss);

  const orderData = orderSheet.getDataRange().getValues();
  const invoiceSources = [];

  const invoiceSheet = ss.getSheetByName('TCS Invoice');
  if (invoiceSheet) {
    const data = invoiceSheet.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, ''));
      const parcelIdx = findHeaderIndex(headers, ['parcelno', 'consignment#']);
      const codIdx = findHeaderIndex(headers, 'codamount');
      const statusIdx = findHeaderIndex(headers, ['status', 'deliverystatus']);
      const specialIdx = findHeaderIndex(headers, 'specialinstruction');
      if (parcelIdx >= 0 && codIdx >= 0 && statusIdx >= 0) {
        invoiceSources.push({
          sheet: invoiceSheet,
          data: data,
          parcelIdx: parcelIdx,
          codIdx: codIdx,
          statusIdx: statusIdx,
          specialIdx: specialIdx
        });
      }
    }
  }

  const newSheet = ss.getSheetByName(NEW_INVOICE_SHEET_NAME);
  if (newSheet) {
    const data = newSheet.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, ''));
      const parcelIdx = findHeaderIndex(headers, ['parcelno', 'consignment#']);
      const codIdx = findHeaderIndex(headers, 'codamount');
      const statusIdx = findHeaderIndex(headers, ['status', 'deliverystatus']);
      const specialIdx = findHeaderIndex(headers, 'specialinstruction');
      if (parcelIdx >= 0 && codIdx >= 0 && statusIdx >= 0) {
        invoiceSources.push({
          sheet: newSheet,
          data: data,
          parcelIdx: parcelIdx,
          codIdx: codIdx,
          statusIdx: statusIdx,
          specialIdx: specialIdx
        });
      }
    }
  }

  if (!invoiceSources.length) return;

  const invoiceMap = {};
  invoiceSources.forEach((source, sourceIndex) => {
    for (let i = 1; i < source.data.length; i++) {
      const rawParcel = source.data[i][source.parcelIdx];
      const cleaned = String(rawParcel).replace(/\s+/g, '').trim();
      if (!cleaned) continue;
      const status = String(source.data[i][source.statusIdx]).toLowerCase();
      const entry = invoiceMap[cleaned];
      const shouldReplace =
        !entry ||
        (status === 'delivered' && entry.status !== 'delivered') ||
        (entry && entry.sourceIndex === sourceIndex && i > entry.row);
      if (shouldReplace) {
        invoiceMap[cleaned] = {
          cod: normalizeInvoiceAmount(source.data[i][source.codIdx]),
          status: status,
          row: i,
          sourceIndex: sourceIndex
        };
      }
    }
  });

  const matchedParcels = new Set();
  const paidRowsPerSource = invoiceSources.map(() => new Set());

  const headers = orderData[0];
  const parcelCol = headers.indexOf('Parcel number');
  if (parcelCol < 0) return;
  let statusCol = headers.indexOf('Shipping Status');
  if (statusCol < 0) statusCol = headers.indexOf('Status');
  const deliveryCol = headers.indexOf('Delivery Status');
  const orderCol = headers.indexOf('Order Number');
  const amountCol = headers.indexOf('Amount');
  const resultCol = 13; // Column N (0-based)

  const orderMap = {};
  if (orderCol >= 0) {
    for (let r = 1; r < orderData.length; r++) {
      const num = String(orderData[r][orderCol] || '').trim().replace(/^#/, '');
      if (num) orderMap[num] = r;
    }
  }

  for (let r = 1; r < orderData.length; r++) {
    const shippingStatus = statusCol >= 0 ? String(orderData[r][statusCol]).toLowerCase() : '';
    const rawParcel = String(orderData[r][parcelCol] || '').replace(/\s+/g, '').trim();
    const rec = invoiceMap[rawParcel];
    if (rec) matchedParcels.add(rawParcel);
    const deliveryCell = deliveryCol >= 0 ? orderSheet.getRange(r + 1, deliveryCol + 1) : null;
    const currentResult = String(orderData[r][resultCol] || '').trim();
    if (currentResult === 'Paid ✅' && deliveryCell) {
      const currentDelivery = String(orderData[r][deliveryCol] || '').toLowerCase();
      if (rec && rec.status === 'delivered') {
        if (currentDelivery !== 'delivered') deliveryCell.setValue('Delivered');
        paidRowsPerSource[rec.sourceIndex].add(rec.row);
      }
    }
    if (shippingStatus === 'dispatched') {
      let result = 'Dispatched – No COD ❌';
      if (rec && rec.status === 'delivered' && rec.cod && rec.cod > 0) {
        result = 'Paid ✅';
        if (deliveryCell) deliveryCell.setValue('Delivered');
        paidRowsPerSource[rec.sourceIndex].add(rec.row);
      }
      orderSheet.getRange(r + 1, resultCol + 1).setValue(result);
    } else if (!shippingStatus && rec && rec.status === 'delivered' && rec.cod && rec.cod > 0) {
      if (deliveryCell) deliveryCell.setValue('Delivered');
      orderSheet.getRange(r + 1, resultCol + 1).setValue('Paid ✅');
      paidRowsPerSource[rec.sourceIndex].add(rec.row);
    }
  }

  if (orderCol >= 0 && amountCol >= 0) {
    invoiceSources.forEach((source, sourceIndex) => {
      if (source.specialIdx < 0) return;
      for (let i = 1; i < source.data.length; i++) {
        const status = String(source.data[i][source.statusIdx]).toLowerCase();
        const codVal = normalizeInvoiceAmount(source.data[i][source.codIdx]);
        if (status === 'delivered' && (!codVal || codVal === 0)) {
          const instr = String(source.data[i][source.specialIdx] || '');
          const match = instr.match(/(\d+)/);
          if (match) {
            const num = match[1];
            const row = orderMap[num];
            if (row !== undefined) {
              const amount = orderData[row][amountCol];
              if (!amount || parseFloat(amount) === 0) {
                if (deliveryCol >= 0) orderSheet.getRange(row + 1, deliveryCol + 1).setValue('Delivered');
                orderSheet.getRange(row + 1, resultCol + 1).setValue('Paid \u2013 Bank Transfer \u2705');
                matchedParcels.add(String(source.data[i][source.parcelIdx]).replace(/\s+/g, '').trim());
                paidRowsPerSource[sourceIndex].add(i);
              }
            }
          }
        }
      }
    });
  }

  invoiceSources.forEach((source, sourceIndex) => {
    const lastCol = source.sheet.getLastColumn();
    if (source.data.length > 1) {
      source.sheet.getRange(2, 1, source.data.length - 1, lastCol).setBackground(null);
      for (let i = 1; i < source.data.length; i++) {
        const cleaned = String(source.data[i][source.parcelIdx]).replace(/\s+/g, '').trim();
        const status = String(source.data[i][source.statusIdx]).toLowerCase();
        if (status === 'delivered' && !matchedParcels.has(cleaned)) {
          source.sheet.getRange(i + 1, 1, 1, lastCol).setBackground('#fff2cc');
        } else if (paidRowsPerSource[sourceIndex].has(i)) {
          source.sheet.getRange(i + 1, 1, 1, lastCol).setBackground('#ccffcc');
        }
      }
    }
  });

  SpreadsheetApp.flush();
}
