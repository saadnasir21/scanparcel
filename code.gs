// ==== a single-entry undo stack ====
var DOC_PROPS = PropertiesService.getDocumentProperties();

// caching for faster parcel lookups
var PARCEL_INDEX_KEY = 'parcelIndex';
// cache parcel lookups for a full day to avoid rebuilding the index
var PARCEL_CACHE_TTL = 24 * 60 * 60; // seconds

// maximum size of a cache entry in bytes
var CACHE_MAX_BYTES = 100 * 1024;

// disable row highlighting to speed up large batch scans
var HIGHLIGHT_ROWS = false;

// cache customer info for duplicate checks
var CUSTOMER_INDEX_KEY = 'customerIndex';

// caches for summary sheets
var DISPATCH_PROD_INDEX_KEY  = 'dispatchProdIndex';
var DISPATCH_DAY_INDEX_KEY   = 'dispatchDayIndex';
var RETURN_PROD_INDEX_KEY    = 'returnProdIndex';
var RETURN_DAY_INDEX_KEY     = 'returnDayIndex';
var SUMMARY_CACHE_TTL        = 24 * 60 * 60; // seconds

function getParcelIndex(sheet, parcelCol) {
  if (!sheet) return {};
  var cache = CacheService.getDocumentCache();
  var raw = cache.get(PARCEL_INDEX_KEY);
  if (raw) return JSON.parse(raw);

  // check for chunked cache pieces
  var combined = {};
  var letters = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  for (var i = 0; i < letters.length; i++) {
    var part = cache.get(PARCEL_INDEX_KEY + ':' + letters[i]);
    if (part) {
      Object.assign(combined, JSON.parse(part));
    }
  }
  if (Object.keys(combined).length) return combined;

  var last = sheet.getLastRow();
  var values = sheet.getRange(2, parcelCol, Math.max(last - 1, 0), 1).getValues();
  var map = {};
  for (var j = 0; j < values.length; j++) {
    var key = String(values[j][0]).replace(/\s+/g, '').toUpperCase();
    if (key) map[key] = j + 2; // adjust for header row
  }

  var json = JSON.stringify(map);
  var byteLen = Utilities.newBlob(json).getBytes().length;
  if (byteLen <= CACHE_MAX_BYTES) {
    cache.put(PARCEL_INDEX_KEY, json, PARCEL_CACHE_TTL);
  } else {
    // split map into chunks by first character
    var buckets = {};
    Object.keys(map).forEach(function(k) {
      var ch = k.charAt(0);
      if (!buckets[ch]) buckets[ch] = {};
      buckets[ch][k] = map[k];
    });
    var skipped = false;
    for (var ch in buckets) {
      var chunkJson = JSON.stringify(buckets[ch]);
      var chunkLen = Utilities.newBlob(chunkJson).getBytes().length;
      if (chunkLen <= CACHE_MAX_BYTES) {
        cache.put(PARCEL_INDEX_KEY + ':' + ch, chunkJson, PARCEL_CACHE_TTL);
      } else {
        Logger.log('Parcel index chunk for ' + ch + ' exceeded cache size; skipping cache for this chunk.');
        skipped = true;
      }
    }
    if (skipped) {
      Logger.log('Parcel index exceeded cache size; some chunks were not cached.');
    }
  }
  return map;
}

function invalidateParcelIndex() {
  var cache = CacheService.getDocumentCache();
  cache.remove(PARCEL_INDEX_KEY);
  var letters = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  for (var i = 0; i < letters.length; i++) {
    cache.remove(PARCEL_INDEX_KEY + ':' + letters[i]);
  }
}

function getCustomerIndex(sheet, nameCol, phoneCol, statusCol) {
  var cache = CacheService.getDocumentCache();
  var raw = cache.get(CUSTOMER_INDEX_KEY);
  if (raw) return JSON.parse(raw);

  var last = sheet.getLastRow();
  var values = sheet.getRange(2, 1, Math.max(last - 1, 0), sheet.getLastColumn()).getValues();
  var names = {}, phones = {};
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var status = statusCol ? row[statusCol - 1] : '';
    if (status === 'Dispatched' || status === 'Returned') continue;
    if (nameCol) {
      var name = String(row[nameCol - 1]).trim().toUpperCase();
      if (name) {
        if (!names[name]) names[name] = [];
        names[name].push(i + 2);
      }
    }
    if (phoneCol) {
      var phone = String(row[phoneCol - 1]).trim();
      if (phone) {
        if (!phones[phone]) phones[phone] = [];
        phones[phone].push(i + 2);
      }
    }
  }
  var idx = { names: names, phones: phones };
  var json = JSON.stringify(idx);
  var byteLen = Utilities.newBlob(json).getBytes().length;
  if (byteLen <= CACHE_MAX_BYTES) {
    cache.put(CUSTOMER_INDEX_KEY, json, PARCEL_CACHE_TTL);
  } else {
    Logger.log('Customer index size ' + byteLen + ' exceeds cache limit; caching skipped.');
  }
  return idx;
}

function invalidateCustomerIndex() {
  CacheService.getDocumentCache().remove(CUSTOMER_INDEX_KEY);
}

/**
 * Remove a row's references from the cached customer index so we
 * don't rebuild the entire index on every scan.
 * @param {number} row The 1-based row number that was updated.
 * @param {string} name Customer name value from that row.
 * @param {string} phone Phone number value from that row.
 */
function removeFromCustomerIndex(row, name, phone) {
  var cache = CacheService.getDocumentCache();
  var raw = cache.get(CUSTOMER_INDEX_KEY);
  if (!raw) return;
  var idx = JSON.parse(raw);
  var changed = false;
  if (name) {
    var key = String(name).trim().toUpperCase();
    var arr = idx.names[key];
    if (arr) {
      idx.names[key] = arr.filter(function(r){ return r !== row; });
      if (!idx.names[key].length) delete idx.names[key];
      changed = true;
    }
  }
  if (phone) {
    var key2 = String(phone).trim();
    var arr2 = idx.phones[key2];
    if (arr2) {
      idx.phones[key2] = arr2.filter(function(r){ return r !== row; });
      if (!idx.phones[key2].length) delete idx.phones[key2];
      changed = true;
    }
  }
  if (changed) {
    cache.put(CUSTOMER_INDEX_KEY, JSON.stringify(idx), PARCEL_CACHE_TTL);
  }
}

// ----- summary sheet caches -----
function dateKey(d) {
  var t = d instanceof Date ? d : new Date(d);
  t = new Date(t.getFullYear(), t.getMonth(), t.getDate());
  return String(t.getTime());
}

function getDispatchProdIndex(sheet) {
  var cache = CacheService.getDocumentCache();
  var raw = cache.get(DISPATCH_PROD_INDEX_KEY);
  if (raw) return JSON.parse(raw);
  var last = sheet.getLastRow();
  var values = sheet.getRange(2, 1, Math.max(last - 1, 0), 3).getValues();
  var map = {};
  for (var i = 0; i < values.length; i++) {
    var k = dateKey(values[i][0]) + '|' + values[i][1];
    map[k] = i + 2;
  }
  cache.put(DISPATCH_PROD_INDEX_KEY, JSON.stringify(map), SUMMARY_CACHE_TTL);
  return map;
}

function getDispatchDayIndex(sheet) {
  var cache = CacheService.getDocumentCache();
  var raw = cache.get(DISPATCH_DAY_INDEX_KEY);
  if (raw) return JSON.parse(raw);
  var last = sheet.getLastRow();
  var values = sheet.getRange(2, 1, Math.max(last - 1, 0), 1).getValues();
  var map = {};
  for (var i = 0; i < values.length; i++) {
    map[dateKey(values[i][0])] = i + 2;
  }
  cache.put(DISPATCH_DAY_INDEX_KEY, JSON.stringify(map), SUMMARY_CACHE_TTL);
  return map;
}

function getReturnProdIndex(sheet) {
  var cache = CacheService.getDocumentCache();
  var raw = cache.get(RETURN_PROD_INDEX_KEY);
  if (raw) return JSON.parse(raw);
  var last = sheet.getLastRow();
  var values = sheet.getRange(2, 1, Math.max(last - 1, 0), 3).getValues();
  var map = {};
  for (var i = 0; i < values.length; i++) {
    var k = dateKey(values[i][0]) + '|' + values[i][1];
    map[k] = i + 2;
  }
  cache.put(RETURN_PROD_INDEX_KEY, JSON.stringify(map), SUMMARY_CACHE_TTL);
  return map;
}

function getReturnDayIndex(sheet) {
  var cache = CacheService.getDocumentCache();
  var raw = cache.get(RETURN_DAY_INDEX_KEY);
  if (raw) return JSON.parse(raw);
  var last = sheet.getLastRow();
  var values = sheet.getRange(2, 1, Math.max(last - 1, 0), 1).getValues();
  var map = {};
  for (var i = 0; i < values.length; i++) {
    map[dateKey(values[i][0])] = i + 2;
  }
  cache.put(RETURN_DAY_INDEX_KEY, JSON.stringify(map), SUMMARY_CACHE_TTL);
  return map;
}

function invalidateDispatchProdIndex() {
  CacheService.getDocumentCache().remove(DISPATCH_PROD_INDEX_KEY);
}
function invalidateDispatchDayIndex() {
  CacheService.getDocumentCache().remove(DISPATCH_DAY_INDEX_KEY);
}
function invalidateReturnProdIndex() {
  CacheService.getDocumentCache().remove(RETURN_PROD_INDEX_KEY);
}
function invalidateReturnDayIndex() {
  CacheService.getDocumentCache().remove(RETURN_DAY_INDEX_KEY);
}

function invalidateSummaryIndexes() {
  invalidateDispatchProdIndex();
  invalidateDispatchDayIndex();
  invalidateReturnProdIndex();
  invalidateReturnDayIndex();
}

function buildSummaryIndexes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh;
  sh = ss.getSheetByName('Product wise daily dispatch');
  if (sh) {
    invalidateDispatchProdIndex();
    getDispatchProdIndex(sh);
  }
  sh = ss.getSheetByName('Daily Dispatch Parcels');
  if (sh) {
    invalidateDispatchDayIndex();
    getDispatchDayIndex(sh);
  }
  sh = ss.getSheetByName('Product wise daily return');
  if (sh) {
    invalidateReturnProdIndex();
    getReturnProdIndex(sh);
  }
  sh = ss.getSheetByName('Daily Return Parcels');
  if (sh) {
    invalidateReturnDayIndex();
    getReturnDayIndex(sh);
  }
}

/**
 * Add the custom menu.
 */
function onOpen() {
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
    var nameCol   = head.indexOf('Customer Name') + 1;
    var phoneCol  = head.indexOf('Phone Number') + 1;
    var statusCol = head.indexOf('Shipping Status') + 1;
    if (parcelCol) {
      invalidateParcelIndex();
      getParcelIndex(sheet, parcelCol);
    }
    if (nameCol || phoneCol) {
      invalidateCustomerIndex();
      getCustomerIndex(sheet, nameCol, phoneCol, statusCol);
    }
  }
  buildSummaryIndexes();
  var html = HtmlService
    .createHtmlOutputFromFile('ScannerSidebar')
    .setTitle('Parcel Scanner');
  SpreadsheetApp.getUi().showSidebar(html);
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
  if (!sheet) return;
  var shName = sheet.getName();
  if (shName !== 'Sheet1') {
    if (shName === 'Product wise daily dispatch' ||
        shName === 'Daily Dispatch Parcels' ||
        shName === 'Product wise daily return' ||
        shName === 'Daily Return Parcels') {
      invalidateSummaryIndexes();
    }
    return;
  }
  invalidateParcelIndex();
  invalidateCustomerIndex();

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
  }
}

/**
 * First-time scan handler: marks Dispatched or signals confirmReturn.
 */
function processParcelScan(scannedValue) {
  scannedValue = String(scannedValue || '').trim().replace(/\s+/g, '');
  if (!scannedValue) return 'Empty';

  var ss      = SpreadsheetApp.getActiveSpreadsheet(),
      sheet   = ss.getSheetByName("Sheet1");
  if (!sheet) return 'SheetNotFound';
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0],
      parcelCol  = headers.indexOf("Parcel number")+1,
      statusCol  = headers.indexOf("Shipping Status")+1,
      dateCol    = headers.indexOf("Dispatch Date")+1,
      productCol = headers.indexOf("Product name")+1,
      qtyCol     = headers.indexOf("Quantity")+1,
      amountCol  = headers.indexOf("Amount")+1,
      nameCol    = headers.indexOf("Customer Name")+1,
      phoneCol   = headers.indexOf("Phone Number")+1;

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

  // check for other orders by same customer using cached index
  var dupFound = false;
  if (nameCol || phoneCol) {
    var custName = nameCol ? rowData[nameCol-1] : '';
    var phone    = phoneCol ? rowData[phoneCol-1] : '';
    if (custName || phone) {
      var idx = getCustomerIndex(sheet, nameCol, phoneCol, statusCol);
      var keyName = custName ? custName.trim().toUpperCase() : '';
      var keyPhone = phone ? phone.trim() : '';
      if (keyName && idx.names[keyName]) {
        var arr = idx.names[keyName];
        for (var i = 0; i < arr.length; i++) {
          if (arr[i] !== foundRow) { dupFound = true; break; }
        }
      }
      if (!dupFound && keyPhone && idx.phones[keyPhone]) {
        var arr2 = idx.phones[keyPhone];
        for (var j = 0; j < arr2.length; j++) {
          if (arr2[j] !== foundRow) { dupFound = true; break; }
        }
      }
    }
  }
  if (dupFound) return 'confirmDuplicate';

  // write new
  sheet.getRange(foundRow,statusCol).setValue(newStatus);
  removeFromCustomerIndex(foundRow, nameCol ? rowData[nameCol-1] : '',
                         phoneCol ? rowData[phoneCol-1] : '');
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
  scannedValue = String(scannedValue || '').trim().replace(/\s+/g,'');
  if (!scannedValue) return 'Empty';

  var ss      = SpreadsheetApp.getActiveSpreadsheet(),
      sheet   = ss.getSheetByName("Sheet1");
  if (!sheet) return 'SheetNotFound';
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0],
      parcelCol  = headers.indexOf("Parcel number")+1,
      statusCol  = headers.indexOf("Shipping Status")+1,
      dateCol    = headers.indexOf("Dispatch Date")+1,
      productCol = headers.indexOf("Product name")+1,
      qtyCol     = headers.indexOf("Quantity")+1,
      amountCol  = headers.indexOf("Amount")+1,
      nameCol    = headers.indexOf("Customer Name")+1,
      phoneCol   = headers.indexOf("Phone Number")+1,
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
  removeFromCustomerIndex(foundRow, nameCol ? rowData[nameCol-1] : '',
                         phoneCol ? rowData[phoneCol-1] : '');
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
 * After customer duplicate warning, mark Dispatched.
 */
function processParcelConfirmDuplicate(scannedValue) {
  scannedValue = String(scannedValue || '').trim().replace(/\s+/g,'');
  if (!scannedValue) return 'Empty';

  var ss      = SpreadsheetApp.getActiveSpreadsheet(),
      sheet   = ss.getSheetByName("Sheet1");
  if (!sheet) return 'SheetNotFound';
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0],
      parcelCol  = headers.indexOf("Parcel number")+1,
      statusCol  = headers.indexOf("Shipping Status")+1,
      dateCol    = headers.indexOf("Dispatch Date")+1,
      productCol = headers.indexOf("Product name")+1,
      qtyCol     = headers.indexOf("Quantity")+1,
      amountCol  = headers.indexOf("Amount")+1,
      nameCol    = headers.indexOf("Customer Name")+1,
      phoneCol   = headers.indexOf("Phone Number")+1;

  if (!parcelCol) return 'ParcelColNotFound';

  var index = getParcelIndex(sheet, parcelCol);
  var foundRow = index[scannedValue.toUpperCase()] || null;
  if (!foundRow) {
    invalidateParcelIndex();
    index = getParcelIndex(sheet, parcelCol);
    foundRow = index[scannedValue.toUpperCase()] || null;
  }
  if (!foundRow) return 'NotFound';

  var rowData   = sheet.getRange(foundRow,1,1,sheet.getLastColumn()).getValues()[0],
      oldStatus = statusCol ? rowData[statusCol-1] : '',
      oldDate   = dateCol   ? rowData[dateCol-1] : null;
    if (String(oldStatus).trim() === "Cancelled by Customer") {
        return "WasCancelled";
    }

  if (oldStatus==='Dispatched' || oldStatus==='Returned') {
    return oldStatus==='Dispatched' ? 'confirmReturn' : 'AlreadyReturned';
  }

  sheet.getRange(foundRow,statusCol).setValue('Dispatched');
  removeFromCustomerIndex(foundRow, nameCol ? rowData[nameCol-1] : '',
                         phoneCol ? rowData[phoneCol-1] : '');
  var now = new Date(), todayMid = new Date(now.getFullYear(),now.getMonth(),now.getDate());
  sheet.getRange(foundRow,dateCol).setValue(todayMid);

  var products   = String(rowData[productCol-1]).split('\n').map(s=>s.trim()).filter(Boolean),
      quantities = String(rowData[qtyCol-1]).split('\n').map(s=>s.trim()).filter(Boolean),
      orderAmt   = amountCol ? Number(rowData[amountCol-1]||0) : 0;

  updateDispatchSummaries(products, quantities, orderAmt, todayMid);

  DOC_PROPS.setProperty('lastAction', JSON.stringify({
    type:       'dispatch',
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
 * Undo the last scan.
 */
function undoLastScan() {
  var raw = DOC_PROPS.getProperty('lastAction');
  if (!raw) return 'NoAction';
  var act = JSON.parse(raw);

  var ss      = SpreadsheetApp.getActiveSpreadsheet(),
      sheet   = ss.getSheetByName("Sheet1");
  if (!sheet) return 'SheetNotFound';
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0],
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
    var idx = getDispatchProdIndex(prodSh);
    for (var i=0; i<products.length; i++) {
      var name = products[i];
      var qty  = Number(quantities[i]||0);
      var key  = dateKey(today)+'|'+name;
      var row  = idx[key];
      if (row) {
        var cur = Number(prodSh.getRange(row,3).getValue()||0);
        prodSh.getRange(row,3).setValue(cur+qty);
      } else {
        prodSh.appendRow([today, name, qty]);
        idx[key] = prodSh.getLastRow();
      }
    }
    CacheService.getDocumentCache()
      .put(DISPATCH_PROD_INDEX_KEY, JSON.stringify(idx), SUMMARY_CACHE_TTL);
  }

  // Daily parcels
  if (dailySh) {
    var idx2 = getDispatchDayIndex(dailySh);
    var key2 = dateKey(today);
    var r2 = idx2[key2];
    if (r2) {
      var parcels = Number(dailySh.getRange(r2,2).getValue()||0)+1;
      var amt     = Number(dailySh.getRange(r2,3).getValue()||0)+amount;
      dailySh.getRange(r2,2).setValue(parcels);
      dailySh.getRange(r2,3).setValue(amt);
    } else {
      dailySh.appendRow([today, 1, amount]);
      idx2[key2] = dailySh.getLastRow();
    }
    CacheService.getDocumentCache()
      .put(DISPATCH_DAY_INDEX_KEY, JSON.stringify(idx2), SUMMARY_CACHE_TTL);
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
    var idx = getReturnProdIndex(prodSh);
    for (var i=0; i<products.length; i++) {
      var name = products[i];
      var qty  = Number(quantities[i]||0);
      var key  = dateKey(today)+'|'+name;
      var row  = idx[key];
      if (row) {
        var cur = Number(prodSh.getRange(row,3).getValue()||0);
        prodSh.getRange(row,3).setValue(cur+qty);
      } else {
        prodSh.appendRow([today, name, qty]);
        idx[key] = prodSh.getLastRow();
      }
    }
    CacheService.getDocumentCache()
      .put(RETURN_PROD_INDEX_KEY, JSON.stringify(idx), SUMMARY_CACHE_TTL);
  }

  // Daily returns
  if (dailySh) {
    var idx2 = getReturnDayIndex(dailySh);
    var key2 = dateKey(today);
    var r2 = idx2[key2];
    if (r2) {
      var parcels = Number(dailySh.getRange(r2,2).getValue()||0)+1;
      var amt     = Number(dailySh.getRange(r2,3).getValue()||0)+amount;
      dailySh.getRange(r2,2).setValue(parcels);
      dailySh.getRange(r2,3).setValue(amt);
    } else {
      dailySh.appendRow([today, 1, amount]);
      idx2[key2] = dailySh.getLastRow();
    }
    CacheService.getDocumentCache()
      .put(RETURN_DAY_INDEX_KEY, JSON.stringify(idx2), SUMMARY_CACHE_TTL);
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
    var idx = getDispatchProdIndex(prodSh);
    for (var i=0; i<products.length; i++) {
      var key = dateKey(target)+'|'+products[i];
      var row = idx[key];
      if (row) {
        var newQty = Number(prodSh.getRange(row,3).getValue()||0) - Number(quantities[i]||0);
        if (newQty > 0) {
          prodSh.getRange(row,3).setValue(newQty);
        } else {
          prodSh.deleteRow(row);
          delete idx[key];
          for (var k in idx) if (idx[k] > row) idx[k]--;
        }
      }
    }
    CacheService.getDocumentCache()
      .put(DISPATCH_PROD_INDEX_KEY, JSON.stringify(idx), SUMMARY_CACHE_TTL);
  }
  // 2) Daily parcels sheet
  if (dailySh) {
    var idx2 = getDispatchDayIndex(dailySh);
    var key2 = dateKey(target);
    var r2 = idx2[key2];
    if (r2) {
      var parcels = Number(dailySh.getRange(r2,2).getValue()||0) - 1;
      var amt     = Number(dailySh.getRange(r2,3).getValue()||0) - amount;
      if (parcels > 0) {
        dailySh.getRange(r2,2).setValue(parcels);
        dailySh.getRange(r2,3).setValue(amt);
      } else {
        dailySh.deleteRow(r2);
        delete idx2[key2];
        for (var j in idx2) if (idx2[j] > r2) idx2[j]--;
      }
      CacheService.getDocumentCache()
        .put(DISPATCH_DAY_INDEX_KEY, JSON.stringify(idx2), SUMMARY_CACHE_TTL);
      return true;
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
    var idx = getDispatchProdIndex(prodSh);
    for (var i=0;i<products.length;i++) {
      var name = products[i];
      var qty  = Number(quantities[i]||0);
      var key  = dateKey(today)+'|'+name;
      var row  = idx[key];
      if (row) {
        var cur = Number(prodSh.getRange(row,3).getValue()||0);
        prodSh.getRange(row,3).setValue(cur+qty);
      } else {
        prodSh.appendRow([today, name, qty]);
        idx[key] = prodSh.getLastRow();
      }
    }
    CacheService.getDocumentCache()
      .put(DISPATCH_PROD_INDEX_KEY, JSON.stringify(idx), SUMMARY_CACHE_TTL);
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
    var idx = getDispatchProdIndex(prodSh);
    for (var i=0;i<products.length;i++) {
      var key = dateKey(target)+'|'+products[i];
      var row = idx[key];
      if (row) {
        var newQty = Number(prodSh.getRange(row,3).getValue()||0) - Number(quantities[i]||0);
        if (newQty > 0) {
          prodSh.getRange(row,3).setValue(newQty);
        } else {
          prodSh.deleteRow(row);
          delete idx[key];
          for (var k in idx) if (idx[k] > row) idx[k]--;
        }
      }
    }
    CacheService.getDocumentCache()
      .put(DISPATCH_PROD_INDEX_KEY, JSON.stringify(idx), SUMMARY_CACHE_TTL);
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
    var idx = getReturnProdIndex(prodSh);
    for (var i=0;i<products.length;i++) {
      var name = products[i];
      var qty  = Number(quantities[i]||0);
      var key  = dateKey(today)+'|'+name;
      var row  = idx[key];
      if (row) {
        var cur = Number(prodSh.getRange(row,3).getValue()||0);
        prodSh.getRange(row,3).setValue(cur+qty);
      } else {
        prodSh.appendRow([today, name, qty]);
        idx[key] = prodSh.getLastRow();
      }
    }
    CacheService.getDocumentCache()
      .put(RETURN_PROD_INDEX_KEY, JSON.stringify(idx), SUMMARY_CACHE_TTL);
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
    var idx = getReturnProdIndex(prodSh);
    for (var i=0;i<products.length;i++) {
      var key = dateKey(target)+'|'+products[i];
      var row = idx[key];
      if (row) {
        var newQty = Number(prodSh.getRange(row,3).getValue()||0) - Number(quantities[i]||0);
        if (newQty > 0) {
          prodSh.getRange(row,3).setValue(newQty);
        } else {
          prodSh.deleteRow(row);
          delete idx[key];
          for (var k in idx) if (idx[k] > row) idx[k]--;
        }
      }
    }
    CacheService.getDocumentCache()
      .put(RETURN_PROD_INDEX_KEY, JSON.stringify(idx), SUMMARY_CACHE_TTL);
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
    var idx = getReturnProdIndex(prodSh);
    for (var i=0; i<products.length; i++) {
      var key = dateKey(target)+'|'+products[i];
      var row = idx[key];
      if (row) {
        var newQty = Number(prodSh.getRange(row,3).getValue()||0) - Number(quantities[i]||0);
        if (newQty > 0) {
          prodSh.getRange(row,3).setValue(newQty);
        } else {
          prodSh.deleteRow(row);
          delete idx[key];
          for (var k in idx) if (idx[k] > row) idx[k]--;
        }
      }
    }
    CacheService.getDocumentCache()
      .put(RETURN_PROD_INDEX_KEY, JSON.stringify(idx), SUMMARY_CACHE_TTL);
  }
  // 2) Daily return parcels sheet
  if (dailySh) {
    var idx2 = getReturnDayIndex(dailySh);
    var key2 = dateKey(target);
    var r2 = idx2[key2];
    if (r2) {
      var parcels = Number(dailySh.getRange(r2,2).getValue()||0) - 1;
      var amt     = Number(dailySh.getRange(r2,3).getValue()||0) - amount;
      if (parcels > 0) {
        dailySh.getRange(r2,2).setValue(parcels);
        dailySh.getRange(r2,3).setValue(amt);
      } else {
        dailySh.deleteRow(r2);
        delete idx2[key2];
        for (var j in idx2) if (idx2[j] > r2) idx2[j]--;
      }
      CacheService.getDocumentCache()
        .put(RETURN_DAY_INDEX_KEY, JSON.stringify(idx2), SUMMARY_CACHE_TTL);
      return true;
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
  var parcelNumber = String(parcelNumberRaw || '').trim().replace(/\s+/g, '');
  if (!parcelNumber) return 'Empty';

  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getSheetByName("Sheet1");
  if (!sheet) return 'SheetNotFound';
  var head   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var parcelCol = head.indexOf("Parcel number") + 1;
  var statusCol = head.indexOf("Shipping Status") + 1;
  var dateCol   = head.indexOf("Dispatch Date") + 1;
  var orderCol  = head.indexOf("Order Number") + 1;
  var nameCol   = head.indexOf("Customer Name") + 1;
  var phoneCol  = head.indexOf("Phone Number") + 1;

  if (!parcelCol || !statusCol || !orderCol) return 'MissingHeaders';

  var index = getParcelIndex(sheet, parcelCol);
  var foundRow = index[parcelNumber.toUpperCase()] || null;
  if (!foundRow) {
    invalidateParcelIndex();
    index = getParcelIndex(sheet, parcelCol);
    foundRow = index[parcelNumber.toUpperCase()] || null;
  }
  if (!foundRow) return 'NotFound';

  var rowData = sheet.getRange(foundRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var currentStatus = rowData[statusCol - 1];
  if (currentStatus === "Dispatched" || currentStatus === "Returned") {
    return 'TooLate';
  }

  // Set "Cancelled by Customer"
  sheet.getRange(foundRow, statusCol).setValue("Cancelled by Customer");
  removeFromCustomerIndex(foundRow, nameCol ? rowData[nameCol - 1] : '',
                         phoneCol ? rowData[phoneCol - 1] : '');
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

/**
 * Cancel by entering the order number instead of the parcel.
 * The input can be the full order string or just the numeric portion.
 */
function cancelOrderByNumber(orderNumRaw) {
  var orderNum = String(orderNumRaw).trim().replace(/^#/, '');
  if (!orderNum) return 'Empty';

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sheet1");
  if (!sheet) return 'SheetNotFound';
  var head  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var orderCol  = head.indexOf("Order Number") + 1;
  var statusCol = head.indexOf("Shipping Status") + 1;
  var dateCol   = head.indexOf("Dispatch Date") + 1;
  var nameCol   = head.indexOf("Customer Name") + 1;
  var phoneCol  = head.indexOf("Phone Number") + 1;

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
  removeFromCustomerIndex(foundRow, nameCol ? rowData[nameCol - 1] : '',
                         phoneCol ? rowData[phoneCol - 1] : '');
  var todayMid = new Date(); todayMid.setHours(0, 0, 0, 0);
  sheet.getRange(foundRow, dateCol).setValue(todayMid);

  var orderName = rowData[orderCol - 1];
  var orderId   = findOrderIdByName(orderName);
  if (orderId) {
    var ok = cancelOrderById(orderId);
    return ok ? 'Cancelled' : 'ShopifyFail';
  } else {
    return 'OrderNotFoundOnShopify';
  }
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
  if (!sheet) return 'SheetNotFound';
  var head   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var parcelCol  = head.indexOf("Parcel number") + 1;
  var statusCol  = head.indexOf("Shipping Status") + 1;
  var dateCol    = head.indexOf("Dispatch Date") + 1;
  var productCol = head.indexOf("Product name") + 1;
  var qtyCol     = head.indexOf("Quantity") + 1;
  var amountCol  = head.indexOf("Amount") + 1;
  var orderCol   = head.indexOf("Order Number") + 1;
  var nameCol    = head.indexOf("Customer Name") + 1;
  var phoneCol   = head.indexOf("Phone Number") + 1;

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
  var oldDate   = rowData[dateCol - 1];
  var orderNum  = orderCol ? rowData[orderCol - 1] : '';

  if (String(oldStatus).trim() === 'Cancelled by Customer' && newStatus === 'Dispatched') {
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
  removeFromCustomerIndex(foundRow, nameCol ? rowData[nameCol - 1] : '',
                         phoneCol ? rowData[phoneCol - 1] : '');

  // add new summary data if needed
  if (newStatus === 'Dispatched' || newStatus === 'Dispatch through Bykea') {
    updateDispatchSummaries(products, quantities, orderAmt, dateObj);
  } else if (newStatus === 'Dispatch through Local Rider') {
    updateDispatchInventoryOnly(products, quantities, orderAmt, dateObj, orderNum);
  } else if (newStatus === 'Returned') {
    updateReturnSummaries(products, quantities, orderAmt, dateObj);
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

/**
 * Reconcile COD payments from invoice data and mark orders.
 */
function reconcileCODPayments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName('Sheet1');
  const invoiceSheet = ss.getSheetByName('TCS Invoice');
  if (!orderSheet || !invoiceSheet) return;

  const orderData = orderSheet.getDataRange().getValues();
  const invoiceData = invoiceSheet.getDataRange().getValues();
  if (invoiceData.length < 2) return;

  // map invoice headers
  const invHeaders = invoiceData[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, ''));
  const parcelIdx = invHeaders.indexOf('parcelno');
  const codIdx = invHeaders.indexOf('codamount');
  const statusIdx = invHeaders.indexOf('status');
  const specialIdx = invHeaders.indexOf('specialinstruction');
  if (parcelIdx < 0 || codIdx < 0 || statusIdx < 0) return;

  // build lookup of parcel → {status, cod, row}
  const invoiceMap = {};
  for (let i = 1; i < invoiceData.length; i++) {
    const rawParcel = invoiceData[i][parcelIdx];
    const cleaned = String(rawParcel).replace(/\s+/g, '').trim();
    if (!cleaned) continue;
    const status = String(invoiceData[i][statusIdx]).toLowerCase();
    const entry = invoiceMap[cleaned];
    if (!entry || status === 'delivered' || (entry.status !== 'delivered' && i > entry.row)) {
      invoiceMap[cleaned] = {
        cod: invoiceData[i][codIdx],
        status: status,
        row: i
      };
    }
  }

  const matchedParcels = new Set();
  const paidRows = new Set();

  const headers = orderData[0];
  const parcelCol = headers.indexOf('Parcel number');
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
      if (currentDelivery !== 'delivered') deliveryCell.setValue('Delivered');
      if (rec) paidRows.add(rec.row);
    }
    if (shippingStatus === 'dispatched') {
      let result = 'Dispatched – No COD ❌';
      if (rec && rec.status === 'delivered' && rec.cod && parseFloat(rec.cod) > 0) {
        result = 'Paid ✅';
        if (deliveryCell) deliveryCell.setValue('Delivered');
        paidRows.add(rec.row);
      }
      orderSheet.getRange(r + 1, resultCol + 1).setValue(result);
    } else if (!shippingStatus && rec && rec.status === 'delivered' && rec.cod && parseFloat(rec.cod) > 0) {
      if (deliveryCell) deliveryCell.setValue('Delivered');
      orderSheet.getRange(r + 1, resultCol + 1).setValue('Paid ✅');
      paidRows.add(rec.row);
    }
  }

  // handle delivered parcels with 0 COD linked by order number
  if (specialIdx >= 0 && orderCol >= 0 && amountCol >= 0) {
    for (let i = 1; i < invoiceData.length; i++) {
      const status = String(invoiceData[i][statusIdx]).toLowerCase();
      const codVal = invoiceData[i][codIdx];
      if (status === 'delivered' && (!codVal || parseFloat(codVal) === 0)) {
        const instr = String(invoiceData[i][specialIdx] || '');
        const match = instr.match(/(\d+)/);
        if (match) {
          const num = match[1];
          const row = orderMap[num];
          if (row !== undefined) {
            const amount = orderData[row][amountCol];
            if (!amount || parseFloat(amount) === 0) {
              if (deliveryCol >= 0) orderSheet.getRange(row + 1, deliveryCol + 1).setValue('Delivered');
              orderSheet.getRange(row + 1, resultCol + 1).setValue('Paid \u2013 Bank Transfer \u2705');
              matchedParcels.add(String(invoiceData[i][parcelIdx]).replace(/\s+/g, '').trim());
              paidRows.add(i);
            }
          }
        }
      }
    }
  }

  // highlight invoice rows that were not matched or were paid
  const lastCol = invoiceSheet.getLastColumn();
  if (invoiceData.length > 1) {
    invoiceSheet.getRange(2, 1, invoiceData.length - 1, lastCol).setBackground(null);
    for (let i = 1; i < invoiceData.length; i++) {
      const cleaned = String(invoiceData[i][parcelIdx]).replace(/\s+/g, '').trim();
      const status = String(invoiceData[i][statusIdx]).toLowerCase();
      if (status === 'delivered' && !matchedParcels.has(cleaned)) {
        invoiceSheet.getRange(i + 1, 1, 1, lastCol).setBackground('#fff2cc');
      } else if (paidRows.has(i)) {
        invoiceSheet.getRange(i + 1, 1, 1, lastCol).setBackground('#ccffcc');
      }
    }
  }

  SpreadsheetApp.flush();
}
