// ============================================================
// FOOD DISTRIBUTION COUPON SYSTEM - Google Apps Script Backend
// ============================================================
//
// ARCHITECTURE:
//   - One sheet tab per event date (e.g. "2026-10-01", "2026-10-02")
//   - Each tab: Timestamp | Name | Phone | Plates | CouponCode | Redeemed | RedeemedAt | Source
//   - "DateSettings" tab: Date | Enabled | MaxTokens
//   - Config loaded from config.gs via getConfig()
//
// SETUP INSTRUCTIONS:
//   1. Go to https://script.google.com and create a new project
//   2. Create TWO script files:
//      - config.gs  (paste config.gs content — all settings live here)
//      - Code.gs    (paste this file — all logic lives here)
//   3. Update config.gs with your values (NON_PROD for testing, PROD for live)
//   4. Run setupSheet() once to initialize DateSettings and date tabs
//   5. Deploy as Web App (Execute as: Me, Access: Anyone)
//   6. Copy the Web App URL and paste it in your HTML files
//
// ============================================================

// ============================================================
// CONSTANTS
// ============================================================

/** Column indices for date-tab rows (1-based for Sheets API) */
var COL = {
  TIMESTAMP:   1,
  NAME:        2,
  PHONE:       3,
  PLATES:      4,
  COUPON_CODE: 5,
  REDEEMED:    6,
  REDEEMED_AT: 7,
  SOURCE:      8
};

/** Headers used in every date-specific tab */
var DATE_TAB_HEADERS = [
  'Timestamp', 'Name', 'Phone', 'Plates',
  'CouponCode', 'Redeemed', 'RedeemedAt', 'Source'
];

// ============================================================
// WEB APP ENTRY POINTS
// ============================================================

/**
 * Handles GET requests — all actions route through here.
 *
 * Supports two response modes:
 *   - JSONP: if ?callback=functionName is present, wraps response in callback(...)
 *     This avoids CORS issues when calling from external HTML pages.
 *   - JSON: standard JSON response (works for same-origin or server-to-server calls)
 */
function doGet(e) {
  var params = e.parameter;
  var action = params.action;
  var callback = params.callback; // JSONP callback function name

  var result;

  try {
    switch (action) {
      case 'register':
        result = registerUser(params);
        break;
      case 'lookup':
        result = lookupCoupon(params.code, params.date);
        break;
      case 'redeem':
        result = redeemCoupon(params.code, params.date);
        break;
      case 'walkin':
        result = registerWalkin(params);
        break;
      case 'list':
        result = listEntries(params.date || 'all');
        break;
      case 'summary':
        result = getSummary(params.date);
        break;
      case 'getDateSettings':
        result = getDateSettings();
        break;
      case 'toggleDate':
        result = toggleDate(params.date, params.enabled);
        break;
      case 'setDateLimit':
        result = setDateLimit(params.date, params.maxTokens);
        break;
      case 'getMembersStatus':
        result = getMembersStatus(params.date);
        break;
      case 'setMemberRemark':
        result = setMemberRemark(params.row, params.phone, params.remark);
        break;
      case 'quickSearch':
        result = quickSearch(params.query, params.date, params.includeRedeemed);
        break;
      default:
        result = { status: 'error', message: 'Unknown action: ' + (action || '(none)') };
    }
  } catch (err) {
    result = { status: 'error', message: err.toString() };
    var config = getConfig();
    if (config.ENABLE_LOGGING) {
      Logger.log('doGet error [action=' + action + ']: ' + err.toString());
    }
  }

  var jsonString = JSON.stringify(result);

  // If a callback parameter is provided, return JSONP (wrapped in function call)
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + jsonString + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // Otherwise, return plain JSON
  return ContentService
    .createTextOutput(jsonString)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handles POST requests (registration).
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var result = registerUser(data);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// SHEET HELPERS — DATE-WISE TABS
// ============================================================

var _ssCache_ = null;
/**
 * Opens the spreadsheet by config ID. Cached for the rest of THIS execution
 * so helper functions that each call getSpreadsheet_() don't each pay for a
 * fresh openById() round-trip (that was a big part of the slowness).
 */
function getSpreadsheet_() {
  if (!_ssCache_) {
    _ssCache_ = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  }
  return _ssCache_;
}

function getSheetForDate(date) {
  if (!date) throw new Error('getSheetForDate: date is required');

  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(date);

  if (!sheet) {
    sheet = ss.insertSheet(date);
    sheet.appendRow(DATE_TAB_HEADERS);
    sheet.getRange(1, 1, 1, DATE_TAB_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    // Keep Timestamp / Phone / RedeemedAt as plain text (no auto number/date conversion)
    var rows = Math.max(sheet.getMaxRows() - 1, 1);
    sheet.getRange(2, COL.TIMESTAMP, rows, 1).setNumberFormat('@');
    sheet.getRange(2, COL.PHONE, rows, 1).setNumberFormat('@');
    sheet.getRange(2, COL.REDEEMED_AT, rows, 1).setNumberFormat('@');
  }

  return sheet;
}

/**
 * Return all date-tab sheets that currently exist in the spreadsheet.
 * A tab is considered a date tab if its name matches YYYY-MM-DD format.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet[]}
 */
function getAllDateSheets_() {
  var ss = getSpreadsheet_();
  var allSheets = ss.getSheets();
  var datePattern = /^\d{4}-\d{2}-\d{2}$/;
  var result = [];
  for (var i = 0; i < allSheets.length; i++) {
    if (datePattern.test(allSheets[i].getName())) {
      result.push(allSheets[i]);
    }
  }
  return result;
}

/**
 * Get today's date as YYYY-MM-DD in the script's timezone.
 * @returns {string}
 */
function getTodayDate_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Safely turn a DateSettings "Date" cell into a 'yyyy-MM-dd' string.
 * Handles all three shapes Google Sheets can hand back:
 *   - plain string  "2026-10-16"
 *   - Date object   (Sheets auto-converted the text)
 *   - number        (serial date — happens when a date cell is re-formatted
 *                    as Plain text; this is what silently broke matching)
 * Dates are formatted in the SPREADSHEET's timezone (where the cell was
 * typed), not the script's, so a date can never shift by a day.
 */
function normalizeDateStr_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, getSpreadsheet_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }
  if (typeof value === 'number' && value > 30000 && value < 80000) {
    var ms = Math.round((value - 25569) * 86400 * 1000);
    return Utilities.formatDate(new Date(ms), 'UTC', 'yyyy-MM-dd');
  }
  return String(value).trim();
}

// ------------------------------------------------------------
// CELL SANITISERS — the single place raw sheet values become clean
// JSON-safe values. Google Sheets silently turns phone numbers into
// NUMBERS and timestamps into Date objects, which crashed the admin
// page (str.replace / phone.indexOf "is not a function").
// Every raw-read spot below goes through these.
// ------------------------------------------------------------
function cellToStr_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toISOString();
  return String(v);
}
function phoneToStr_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return String(Math.round(v));
  return String(v).trim();
}
function digitsOnly_(v) { return phoneToStr_(v).replace(/\D/g, ''); }
/** Last 10 digits — so "+91 98765 43210", "09876543210" and "9876543210" all match. */
function last10_(v) {
  var d = digitsOnly_(v);
  return d.length > 10 ? d.slice(-10) : d;
}
/** Turn one raw date-tab row into a clean entry object. */
function rowToEntry_(row, date, sheetRow) {
  return {
    row: sheetRow,
    timestamp:  cellToStr_(row[COL.TIMESTAMP - 1]),
    name:       cellToStr_(row[COL.NAME - 1]).trim(),
    phone:      phoneToStr_(row[COL.PHONE - 1]),
    date:       date,
    plates:     parseInt(row[COL.PLATES - 1], 10) || 0,
    couponCode: cellToStr_(row[COL.COUPON_CODE - 1]).trim(),
    redeemed:   String(row[COL.REDEEMED - 1]).toLowerCase() === 'yes',
    redeemedAt: cellToStr_(row[COL.REDEEMED_AT - 1]),
    source:     cellToStr_(row[COL.SOURCE - 1]) || 'Online'
  };
}
/** Read data rows (no header) of a date tab. Returns [] for an empty tab. */
function readDataRows_(sheet, numCols) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, numCols || 8).getValues();
}
/** Append a row while keeping Timestamp/Phone/RedeemedAt as TEXT so Sheets can't mangle them. */
function appendEntryRow_(sheet, row) {
  var r = sheet.getLastRow() + 1;
  sheet.getRange(r, COL.TIMESTAMP).setNumberFormat('@');
  sheet.getRange(r, COL.PHONE).setNumberFormat('@');
  sheet.getRange(r, COL.REDEEMED_AT).setNumberFormat('@');
  sheet.getRange(r, 1, 1, row.length).setValues([row]);
  return r;
}

// ============================================================
// CORE FUNCTIONS
// ============================================================

/**
 * Register a user for one or more event dates.
 *
 * Input formats:
 *   1. Multi-date (preferred): data.dates = JSON string of [{date, coupons}, ...]
 *   2. Single-date (legacy):   data.date + data.plates
 *
 * Behaviour changes from v1:
 *   - Writes to per-date tabs instead of a single sheet
 *   - Duplicate phone+date now UPDATES plate count (returns status 'updated')
 *   - Checks daily token limits before writing
 *   - Batch-writes rows per date tab
 */
function registerUser(data) {
  var config = getConfig();
  var name = (data.name || '').toString().trim();
  var phone = (data.phone || '').toString().trim();

  // Validate required fields
  if (!name) name = 'Walk-in';   // name is optional for walk-ins
  if (!phone) {
    return { status: 'error', message: 'Phone is required' };
  }

  // Parse date selections
  var selections = [];
  if (data.dates) {
    try {
      selections = (typeof data.dates === 'string') ? JSON.parse(data.dates) : data.dates;
    } catch (e) {
      return { status: 'error', message: 'Invalid dates format' };
    }
  } else if (data.date && data.plates) {
    selections = [{ date: data.date, coupons: parseInt(data.plates, 10) }];
  }

  if (selections.length === 0) {
    return { status: 'error', message: 'At least one date is required' };
  }

  var registeredCoupons = []; // successfully registered
  var updatedCoupons = [];    // duplicates that were updated
  var fullDates = [];         // dates that exceeded token limit or disabled
  var timestamp = new Date().toISOString();

  // Load date settings ONCE (one sheet read) — used for every date in the loop
  var dateSettingsMap = getDateSettingsMap_();

  // Process each date selection
  for (var i = 0; i < selections.length; i++) {
    var sel = selections[i];
    var date = sel.date;
    var requestedPlates = parseInt(sel.coupons, 10) || 0;

    if (!date || requestedPlates <= 0) continue;

    // Cap plates to max allowed
    if (requestedPlates > config.MAX_PLATES_PER_REGISTRATION) {
      requestedPlates = config.MAX_PLATES_PER_REGISTRATION;
    }

    // --- Check if date is admin-disabled ---
    var settings = dateSettingsMap[date];
    if (settings && !settings.enabled) {
      fullDates.push({ date: date, reason: 'disabled' });
      continue;
    }

    // --- Read the sheet ONCE and reuse for both duplicate check AND row append ---
    var sheet = getSheetForDate(date);
    var sheetData = sheet.getDataRange().getValues(); // single read

    // --- Duplicate check (phone already registered for this date?) ---
    var existing = findByPhoneInData_(sheetData, phone);
    if (existing) {
      // UPDATE: just change plate count on the existing row — no new sheet reads needed
      sheet.getRange(existing.row, COL.PLATES).setValue(requestedPlates);
      updatedCoupons.push({
        date: date,
        coupons: requestedPlates,
        couponCode: existing.couponCode,
        status: 'updated'
      });
      continue;
    }

    // --- Token limit check (new registrations only, uses already-loaded sheetData) ---
    if (settings && settings.maxTokens > 0) {
      var issuedCount = sumPlatesFromData_(sheetData);
      if (issuedCount + requestedPlates > settings.maxTokens) {
        fullDates.push({ date: date, reason: 'full' });
        continue;
      }
    }

    // --- New registration: generate coupon & append row ---
    var couponCode = generateCouponCode(sheetData);
    var newRow = [timestamp, name, phone, requestedPlates, couponCode, 'No', '', 'Online'];
    appendEntryRow_(sheet, newRow);

    registeredCoupons.push({
      date: date,
      coupons: requestedPlates,
      couponCode: couponCode
    });
  }

  // Determine overall status
  var allCoupons = registeredCoupons.concat(updatedCoupons);

  if (allCoupons.length === 0 && fullDates.length > 0) {
    return {
      status: 'full',
      message: 'All selected dates have reached their coupon limit',
      full: fullDates
    };
  }

  if (allCoupons.length === 0) {
    return { status: 'error', message: 'No valid dates selected' };
  }

  invalidateSummaryCache_();
  invalidateDateSettingsCache_();

  // Send notifications for new + updated coupons (all in one message)
  try {
    sendMultiDateNotifications(name, phone, allCoupons);
  } catch (msgErr) {
    if (config.ENABLE_LOGGING) {
      Logger.log('Message send error: ' + msgErr.toString());
    }
  }

  // Determine granular status
  var status = 'success';
  if (updatedCoupons.length > 0 && registeredCoupons.length === 0) {
    status = 'updated';
  } else if (updatedCoupons.length > 0) {
    status = 'partial';
  }

  var response = {
    status: status,
    name: name,
    phone: phone,
    coupons: allCoupons
  };
  if (updatedCoupons.length > 0) response.updated = updatedCoupons;
  if (fullDates.length > 0)      response.full = fullDates;

  return response;
}

/**
 * Register a walk-in guest. Coupons are marked as redeemed immediately.
 * Runs under a script lock so two volunteers entering the same person at the
 * same instant can't create two coupons. Uses ONE sheet read.
 *
 * @param {Object} data - { name, phone, plates }
 */
function registerWalkin(data) {
  var config = getConfig();
  var name = (data.name || '').toString().trim();
  var phone = (data.phone || '').toString().trim();
  var plates = parseInt(data.plates, 10) || 1;

  if (!name) name = 'Walk-in';   // name is optional for walk-ins
  if (!phone) {
    return { status: 'error', message: 'Phone is required' };
  }
  if (plates > config.MAX_PLATES_PER_REGISTRATION) {
    plates = config.MAX_PLATES_PER_REGISTRATION;
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(8000); }
  catch (e) { return { status: 'error', message: 'System busy — please try again in a moment' }; }

  try {
    var date = getTodayDate_();
    var now = new Date().toISOString();

    var sheet = getSheetForDate(date);
    var sheetData = sheet.getDataRange().getValues(); // ONE read

    var existing = findByPhoneInData_(sheetData, phone);
    var existingPlates = existing ? (parseInt(existing.plates, 10) || 0) : 0;

    // Token limit (an existing row is REPLACED, so only count the difference)
    var settings = getDateSettingsMap_()[date];
    if (settings && settings.maxTokens > 0) {
      var issuedCount = sumPlatesFromData_(sheetData) - existingPlates;
      if (issuedCount + plates > settings.maxTokens) {
        return { status: 'error', message: 'Today\'s coupon limit has been reached', date: date };
      }
    }

    var couponCode;
    if (existing) {
      couponCode = existing.couponCode;
      // Columns: Plates | CouponCode | Redeemed | RedeemedAt  (keep the code intact!)
      sheet.getRange(existing.row, COL.PLATES, 1, 4).setValues([[plates, couponCode, 'Yes', now]]);
    } else {
      couponCode = generateCouponCode(sheetData);
      appendEntryRow_(sheet, [now, name, phone, plates, couponCode, 'Yes', now, 'Walk-in']);
    }

    invalidateSummaryCache_();
    invalidateDateSettingsCache_();
    return { status: 'success', couponCode: couponCode, name: name, phone: phone,
             date: date, plates: plates, updated: !!existing, timestamp: now };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Look up a coupon by code.
 *
 * @param {string} code   - The coupon code (e.g. "BHOG-A3X7")
 * @param {string} [date] - Optional date hint for faster lookup
 * @returns {Object} result
 */
function lookupCoupon(code, date) {
  if (!code) return { status: 'error', message: 'Code is required' };

  var entry = findByCode(code.toUpperCase(), date);

  if (!entry) {
    return { status: 'notfound', message: 'Coupon not found' };
  }

  return {
    status: 'found',
    couponCode: entry.couponCode,
    name: entry.name,
    phone: entry.phone,
    date: entry.date,
    plates: entry.plates,
    redeemed: entry.redeemed,
    redeemedAt: entry.redeemedAt,
    source: entry.source
  };
}

/**
 * Mark a coupon as redeemed. Protected by LockService to prevent
 * double-redemption under concurrent requests.
 *
 * @param {string} code   - The coupon code
 * @param {string} [date] - Optional date hint for faster lookup
 * @returns {Object} result
 */
function redeemCoupon(code, date) {
  if (!code) return { status: 'error', message: 'Code is required' };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000); // wait up to 5 seconds
  } catch (lockErr) {
    return { status: 'error', message: 'System busy — please try again in a moment' };
  }

  try {
    var entry = findByCode(code.toUpperCase(), date);

    if (!entry) {
      return { status: 'notfound', message: 'Coupon not found' };
    }

    if (entry.redeemed) {
      return {
        status: 'already_redeemed',
        message: 'Coupon already redeemed',
        redeemedAt: entry.redeemedAt
      };
    }

    // Mark as redeemed on the correct date tab
    var sheet = getSheetForDate(entry.date);
    sheet.getRange(entry.row, COL.REDEEMED).setValue('Yes');
    sheet.getRange(entry.row, COL.REDEEMED_AT).setValue(new Date().toISOString());
    invalidateSummaryCache_();

    return {
      status: 'success',
      message: 'Coupon redeemed successfully',
      couponCode: code.toUpperCase(),
      name: entry.name,
      phone: entry.phone,
      date: entry.date,
      plates: entry.plates
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * List entries, optionally filtered by date.
 *
 * @param {string} dateFilter - A specific date ("2026-10-01") or "all"
 * @returns {Object} result with data array
 */
function listEntries(dateFilter) {
  var entries = [];

  if (dateFilter && dateFilter !== 'all') {
    // Read only the specific date's tab
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName(dateFilter);
    if (sheet) {
      entries = readEntriesFromSheet_(sheet, dateFilter);
    }
  } else {
    // Read all date tabs and merge
    var dateSheets = getAllDateSheets_();
    for (var i = 0; i < dateSheets.length; i++) {
      var tabDate = dateSheets[i].getName();
      var tabEntries = readEntriesFromSheet_(dateSheets[i], tabDate);
      entries = entries.concat(tabEntries);
    }
  }

  return {
    status: 'success',
    data: entries,
    total: entries.length
  };
}

/**
 * Read all data rows from a date-tab sheet and return CLEAN entry objects
 * (phone always a string, timestamps always ISO strings, plates a number).
 * Used by listEntries and quickSearch.
 */
function readEntriesFromSheet_(sheet, date) {
  var rows = readDataRows_(sheet, 8);
  var entries = [];
  for (var i = 0; i < rows.length; i++) {
    if (!rows[i][COL.COUPON_CODE - 1] && !rows[i][COL.NAME - 1]) continue; // skip blank rows
    entries.push(rowToEntry_(rows[i], date, i + 2));
  }
  return entries;
}

/**
 * Fast search used by the Scanner tab's "find by name / phone" box.
 * Reads ONE date tab only (defaults to today). Matches name, coupon code,
 * or phone — phone matching is digits-only, so "98765 43210", "+91 98765…"
 * or just the last 4 digits all work.
 */
function quickSearch(query, date, includeRedeemed) {
  query = (query || '').toString().trim().toLowerCase();
  if (!query) return { status: 'error', message: 'Search text is required' };

  var queryDigits = query.replace(/\D/g, '');
  var searchDate = date || getTodayDate_();
  var wantRedeemed = includeRedeemed === 'true' || includeRedeemed === true;

  var sheet = getSpreadsheet_().getSheetByName(searchDate);
  if (!sheet) return { status: 'success', date: searchDate, data: [] };

  var results = readEntriesFromSheet_(sheet, searchDate).filter(function(e) {
    if (!wantRedeemed && e.redeemed) return false;
    if (e.name.toLowerCase().indexOf(query) !== -1) return true;
    if (e.couponCode.toLowerCase().indexOf(query) !== -1) return true;
    if (queryDigits && digitsOnly_(e.phone).indexOf(queryDigits) !== -1) return true;
    return false;
  });

  return { status: 'success', date: searchDate, data: results };
}

/**
 * Summary of registrations / walk-ins / coupons / redeemed / pending.
 * Pass a date to summarise ONE tab (fast); omit or 'all' to read every tab.
 * Cached 60s per date and cleared on every write.
 */
function getSummary(date) {
  var scope = (date && date !== 'all') ? date : 'all';
  var cache = CacheService.getScriptCache();
  var key = 'summary_v2_' + scope;
  var cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  var result = computeSummary_(scope);
  try { cache.put(key, JSON.stringify(result), 60); } catch (e) { /* ignore */ }
  return result;
}

/** Clear every cached summary after any write that would change its numbers. */
function invalidateSummaryCache_() {
  try {
    var cfg = getConfig();
    var keys = ['summary_v2_all', 'summary_v2_' + getTodayDate_()];
    (cfg.EVENT_DATES || []).forEach(function(d) { keys.push('summary_v2_' + d); });
    CacheService.getScriptCache().removeAll(keys);
  } catch (e) { /* ignore */ }
}

function computeSummary_(scope) {
  var today = getTodayDate_();
  var single = scope && scope !== 'all';
  var focusDate = single ? scope : today;
  var settingsMap = getDateSettingsMap_();
  var ss = getSpreadsheet_();

  var dateSheets;
  if (single) {
    var one = ss.getSheetByName(scope);
    dateSheets = one ? [one] : [];
  } else {
    dateSheets = getAllDateSheets_();
  }

  var days = [];
  var focus = null;

  for (var i = 0; i < dateSheets.length; i++) {
    var tabDate = dateSheets[i].getName();
    var last = dateSheets[i].getLastRow();
    // Only read the 5 columns we need: Plates..Source
    var rows = last >= 2 ? dateSheets[i].getRange(2, COL.PLATES, last - 1, 5).getValues() : [];

    var registered = 0, walkins = 0, totalCoupons = 0, redeemed = 0;
    for (var r = 0; r < rows.length; r++) {
      var source = rows[r][COL.SOURCE - COL.PLATES] || 'Online';
      if (source === 'Walk-in') walkins++; else registered++;
      totalCoupons += parseInt(rows[r][0], 10) || 0;
      if (String(rows[r][COL.REDEEMED - COL.PLATES]).toLowerCase() === 'yes') redeemed++;
    }

    var maxTokens = (settingsMap[tabDate] && settingsMap[tabDate].maxTokens) || 0;
    var obj = {
      date: tabDate,
      registered: registered,
      walkins: walkins,
      totalRegistrations: registered + walkins,
      totalCoupons: totalCoupons,
      redeemed: redeemed,
      pending: (registered + walkins) - redeemed,
      maxTokens: maxTokens
    };
    days.push(obj);

    if (tabDate === focusDate) {
      focus = Object.assign({}, obj, {
        remaining: (maxTokens > 0) ? Math.max(0, maxTokens - totalCoupons) : undefined
      });
    }
  }

  if (!focus && single) {
    var mt = (settingsMap[focusDate] && settingsMap[focusDate].maxTokens) || 0;
    focus = { date: focusDate, registered: 0, walkins: 0, totalRegistrations: 0, totalCoupons: 0,
              redeemed: 0, pending: 0, maxTokens: mt, remaining: mt > 0 ? mt : undefined };
  }

  return {
    status: 'success',
    days: days,
    today: focus,            // "focus" day: the selected date, or real today
    focusDate: focusDate,
    isToday: focusDate === today,
    scope: single ? scope : 'all'
  };
}

// ============================================================
// FIND / SEARCH HELPERS
// ============================================================

/**
 * Find an entry by coupon code across date tabs.
 *
 * Strategy: if a date hint is provided, search only that tab.
 * Otherwise search today's tab first (most common scan scenario),
 * then all other tabs.
 *
 * @param {string} code   - Coupon code (already uppercased by caller)
 * @param {string} [date] - Optional date hint
 * @returns {Object|null}
 */
function findByCode(code, date) {
  var ss = getSpreadsheet_();

  // If date hint is provided, search only that tab
  if (date) {
    var sheet = ss.getSheetByName(date);
    if (sheet) {
      var entry = findCodeInSheet_(sheet, code, date);
      if (entry) return entry;
    }
    return null;
  }

  // No date hint — search today's tab first
  var today = getTodayDate_();
  var todaySheet = ss.getSheetByName(today);
  if (todaySheet) {
    var todayEntry = findCodeInSheet_(todaySheet, code, today);
    if (todayEntry) return todayEntry;
  }

  // Search all other date tabs
  var allSheets = getAllDateSheets_();
  for (var i = 0; i < allSheets.length; i++) {
    var tabDate = allSheets[i].getName();
    if (tabDate === today) continue; // already searched
    var entry = findCodeInSheet_(allSheets[i], code, tabDate);
    if (entry) return entry;
  }

  return null;
}

/**
 * Search a single sheet for a coupon code. Used by lookupCoupon and redeemCoupon.
 * Returns a clean entry (plus 1-based sheet `row`) or null.
 */
function findCodeInSheet_(sheet, code, date) {
  var rows = readDataRows_(sheet, 8);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][COL.COUPON_CODE - 1]).trim().toUpperCase() === code) {
      return rowToEntry_(rows[i], date, i + 2);
    }
  }
  return null;
}

/**
 * Find a registration by phone within an already-loaded 2D data array
 * (row 0 = headers). Compares the LAST 10 DIGITS so formatting differences
 * (+91, spaces, leading 0, numeric cells) never cause a missed duplicate.
 */
function findByPhoneInData_(data, phone) {
  var target = last10_(phone);
  if (!target) return null;
  for (var i = 1; i < data.length; i++) {
    if (last10_(data[i][COL.PHONE - 1]) === target) {
      return {
        row: i + 1, // 1-based
        couponCode: String(data[i][COL.COUPON_CODE - 1]),
        plates: data[i][COL.PLATES - 1],
        redeemed: String(data[i][COL.REDEEMED - 1]).toLowerCase() === 'yes'
      };
    }
  }
  return null;
}

/**
 * Sum Plates column from already-loaded 2D data array.
 * Avoids an extra sheet read for token limit checks.
 *
 * @param {Array[][]} data - Sheet data (row 0 = headers)
 * @returns {number}
 */
function sumPlatesFromData_(data) {
  var total = 0;
  for (var i = 1; i < data.length; i++) {
    total += parseInt(data[i][COL.PLATES - 1], 10) || 0;
  }
  return total;
}

/**
 * Find a registration by phone number within a specific date's sheet.
 * Used when only the sheet object is available (not preloaded data).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} phone
 * @returns {Object|null}
 */
function findByPhoneAndDateInSheet_(sheet, phone) {
  return findByPhoneInData_(sheet.getDataRange().getValues(), phone);
}



/**
 * Generate a coupon code like BHOG-A3X7. If the caller already has the tab's
 * data loaded (2D array incl. header) pass it in — no extra sheet read.
 */
function generateCouponCode(existingData) {
  var config = getConfig();
  var prefix = config.COUPON_PREFIX || 'FOOD';
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // No 0,O,1,I (ambiguous)

  var data = existingData;
  if (!data) {
    var todaySheet = getSpreadsheet_().getSheetByName(getTodayDate_());
    data = todaySheet ? todaySheet.getDataRange().getValues() : [];
  }
  var existingCodes = {};
  for (var r = 1; r < data.length; r++) {
    var c = String(data[r][COL.COUPON_CODE - 1]);
    if (c) existingCodes[c] = true;
  }

  var code, attempts = 0;
  do {
    code = prefix + '-';
    for (var j = 0; j < 4; j++) code += chars.charAt(Math.floor(Math.random() * chars.length));
    attempts++;
  } while (existingCodes[code] && attempts < 50);
  return code;
}

// ============================================================
// DAILY TOKEN LIMIT HELPERS
// ============================================================

/** Total issued plates for a date — reads ONLY the Plates column. */
function getIssuedCountForDate_(date) {
  var sheet = getSpreadsheet_().getSheetByName(date);
  if (!sheet) return 0;
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var vals = sheet.getRange(2, COL.PLATES, last - 1, 1).getValues();
  var total = 0;
  for (var i = 0; i < vals.length; i++) total += parseInt(vals[i][0], 10) || 0;
  return total;
}

/** date -> { enabled, maxTokens }. Last row wins if a date is duplicated. */
function getDateSettingsMap_() {
  var data = getDateSettingsSheet().getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var dateStr = normalizeDateStr_(data[i][0]);
    if (!dateStr) continue;
    map[dateStr] = {
      enabled: String(data[i][1]).toLowerCase() === 'yes',
      maxTokens: parseInt(data[i][2], 10) || 0
    };
  }
  return map;
}

// ============================================================
// DATE SETTINGS (stored in "DateSettings" tab)
// Columns: Date | Enabled | MaxTokens
// ============================================================

/**
 * Get or create the DateSettings sheet (Date | Enabled | MaxTokens).
 * NOTE: the plain-text formatting is only applied when the sheet is CREATED
 * (or by fixDateSettingsSheet). It used to be re-applied on every call, which
 * was a slow write on every request AND converted existing date cells to
 * serial numbers — the root cause of "max shows 0" / toggles not sticking.
 */
function getDateSettingsSheet() {
  var config = getConfig();
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('DateSettings');

  if (!sheet) {
    sheet = ss.insertSheet('DateSettings');
    sheet.appendRow(['Date', 'Enabled', 'MaxTokens']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');

    if (config.EVENT_DATES && config.EVENT_DATES.length > 0) {
      var rows = config.EVENT_DATES.map(function(dateStr) { return [dateStr, 'Yes', 0]; });
      sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    }
  }
  return sheet;
}

/**
 * Get date settings for all configured dates.
 * Returns each date's enabled status, max tokens, issued count, and whether full.
 *
 * @returns {Object}
 */
function getDateSettings() {
  // index.html calls this on every page load to decide which dates are
  // open, and it reads every date tab (for issued counts) on top of the
  // DateSettings tab itself. Cache briefly so a burst of visitors hitting
  // the registration page at once doesn't each trigger a full re-read.
  var cache = CacheService.getScriptCache();
  var cached = cache.get('dateSettings_v1');
  if (cached) return JSON.parse(cached);

  var result = computeDateSettings_();
  try {
    cache.put('dateSettings_v1', JSON.stringify(result), 15); // seconds
  } catch (e) { /* ignore */ }
  return result;
}

/** Clear the cached date settings after any write that would change them. */
function invalidateDateSettingsCache_() {
  try { CacheService.getScriptCache().remove('dateSettings_v1'); } catch (e) { /* ignore */ }
}

function computeDateSettings_() {
  var data = getDateSettingsSheet().getDataRange().getValues();
  var order = [];
  var byDate = {};

  // If duplicate rows exist for a date, the LAST one wins (same rule as getDateSettingsMap_)
  for (var i = 1; i < data.length; i++) {
    var dateStr = normalizeDateStr_(data[i][0]);
    if (!dateStr) continue;
    if (!(dateStr in byDate)) order.push(dateStr);
    byDate[dateStr] = {
      enabled: String(data[i][1]).toLowerCase() === 'yes',
      maxTokens: parseInt(data[i][2], 10) || 0
    };
  }

  var settings = order.map(function(dateStr) {
    var s = byDate[dateStr];
    var issuedCount = getIssuedCountForDate_(dateStr);
    return {
      date: dateStr,
      enabled: s.enabled,
      maxTokens: s.maxTokens,
      issuedCount: issuedCount,
      isFull: s.maxTokens > 0 && issuedCount >= s.maxTokens
    };
  });

  return { status: 'success', dates: settings };
}

/** Toggle a date's enabled/disabled status (updates EVERY row for that date). */
function toggleDate(date, enabled) {
  if (!date) return { status: 'error', message: 'Date is required' };

  var sheet = getDateSettingsSheet();
  var data = sheet.getDataRange().getValues();
  var newValue = (enabled === 'true' || enabled === true) ? 'Yes' : 'No';
  var found = false;

  for (var i = 1; i < data.length; i++) {
    if (normalizeDateStr_(data[i][0]) === date) {
      sheet.getRange(i + 1, 2).setValue(newValue);
      found = true;
    }
  }
  if (!found) appendSettingsRow_(sheet, date, newValue, 0);

  invalidateSummaryCache_();
  invalidateDateSettingsCache_();
  return { status: 'success', date: date, enabled: newValue === 'Yes' };
}

/** Append a DateSettings row with the date stored as TEXT. */
function appendSettingsRow_(sheet, date, enabled, limit) {
  var r = sheet.getLastRow() + 1;
  sheet.getRange(r, 1).setNumberFormat('@');
  sheet.getRange(r, 1, 1, 3).setValues([[date, enabled, limit]]);
}

/**
 * Set the max coupon limit for a date (0 = unlimited).
 * Updates EVERY row for that date. Previously it only updated the FIRST
 * matching row while reads used the LAST one — so with duplicate rows the
 * new 300 was saved but the old 0 was what got displayed.
 */
function setDateLimit(date, maxTokens) {
  if (!date) return { status: 'error', message: 'Date is required' };

  var limit = parseInt(maxTokens, 10);
  if (isNaN(limit) || limit < 0) {
    return { status: 'error', message: 'maxTokens must be a non-negative integer' };
  }

  var sheet = getDateSettingsSheet();
  var data = sheet.getDataRange().getValues();
  var found = false;

  for (var i = 1; i < data.length; i++) {
    if (normalizeDateStr_(data[i][0]) === date) {
      sheet.getRange(i + 1, 3).setValue(limit);
      found = true;
    }
  }
  if (!found) appendSettingsRow_(sheet, date, 'Yes', limit);

  invalidateSummaryCache_();
  invalidateDateSettingsCache_();
  var issued = getIssuedCountForDate_(date);
  return { status: 'success', date: date, maxTokens: limit, issuedCount: issued,
           isFull: limit > 0 && issued >= limit };
}

/**
 * ONE-TIME REPAIR — run from the Apps Script editor (Run > fixDateSettingsSheet).
 * Collapses duplicate DateSettings rows to one per date (keeping the LAST
 * row's values) and rewrites the Date column as plain text.
 */
function fixDateSettingsSheet() {
  var sheet = getDateSettingsSheet();
  var data = sheet.getDataRange().getValues();

  var order = [];
  var latest = {};
  for (var i = 1; i < data.length; i++) {
    var dateStr = normalizeDateStr_(data[i][0]);
    if (!dateStr) continue;
    if (!(dateStr in latest)) order.push(dateStr);
    latest[dateStr] = { enabled: data[i][1], maxTokens: data[i][2] };
  }

  if (data.length > 1) sheet.getRange(2, 1, data.length - 1, 3).clearContent();
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');

  var rows = order.map(function(dateStr) {
    return [dateStr, latest[dateStr].enabled, latest[dateStr].maxTokens];
  });
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  invalidateDateSettingsCache_();
  invalidateSummaryCache_();
  Logger.log('fixDateSettingsSheet: consolidated to ' + rows.length + ' date row(s): ' + order.join(', '));
  return { status: 'success', dates: order };
}

/**
 * ONE-TIME REPAIR — run from the editor (Run > fixDateTabFormats).
 * For every existing date tab: forces Timestamp / Phone / RedeemedAt to plain
 * text and rewrites numeric phone cells as text so nothing gets converted again.
 */
function fixDateTabFormats() {
  getAllDateSheets_().forEach(function(sheet) {
    var last = sheet.getLastRow();
    if (last < 2) return;
    var n = last - 1;
    [COL.TIMESTAMP, COL.PHONE, COL.REDEEMED_AT].forEach(function(c) {
      var rng = sheet.getRange(2, c, n, 1);
      var vals = rng.getValues().map(function(r) {
        var v = r[0];
        return [c === COL.PHONE ? phoneToStr_(v) : cellToStr_(v)];
      });
      rng.setNumberFormat('@');
      rng.setValues(vals);
    });
    Logger.log('Fixed formats on tab ' + sheet.getName());
  });
  invalidateSummaryCache_();
}

// ============================================================
// MESSAGING FUNCTIONS (Twilio WhatsApp + SMS)
// ============================================================

/**
 * Send WhatsApp and/or SMS notification for multi-date registration.
 * Sends one consolidated message with per-coupon links.
 *
 * @param {string} name      - Registrant name
 * @param {string} phone     - Registrant phone
 * @param {Object[]} coupons - Array of { date, coupons, couponCode }
 */
function sendMultiDateNotifications(name, phone, coupons) {
  var config = getConfig();

  // Build per-coupon lines with individual coupon links
  var couponLines = coupons.map(function(c) {
    // Append T12:00:00 so date parsing is midday IST and avoids UTC-midnight off-by-one
    var formattedDate = Utilities.formatDate(
      new Date(c.date + 'T12:00:00'), Session.getScriptTimeZone(), 'EEE, MMM d'
    );
    var couponUrl = config.APP_BASE_URL + '/coupon.html?code=' + encodeURIComponent(c.couponCode);
    return (
      '\uD83D\uDCC5 ' + formattedDate + ' \u2014 ' + c.coupons + ' coupon(s) \u2014 *' + c.couponCode + '*\n' +
      '\uD83D\uDC49 ' + couponUrl
    );
  }).join('\n\n');

  var messageBody =
    '\uD83E\uDEB7 *Agomoni Durga Puja \u2014 Bhog Confirmed!*\n\n' +
    '\u2705 Name: ' + name + '\n\n' +
    'Your Coupons:\n' + couponLines + '\n\n' +
    'Joy Ma Durga! \uD83D\uDE4F';

  // Try WhatsApp first
  if (config.ENABLE_WHATSAPP) {
    try {
      sendTwilioMessage(
        'whatsapp:' + normalizePhone(phone),
        config.TWILIO_WHATSAPP_NUMBER,
        messageBody
      );
      if (config.ENABLE_LOGGING) Logger.log('WhatsApp sent to ' + phone);
      return;
    } catch (waErr) {
      if (config.ENABLE_LOGGING) Logger.log('WhatsApp failed, falling back to SMS: ' + waErr.toString());
    }
  }

  // Fallback to SMS (shorter message due to character limits)
  if (config.ENABLE_SMS) {
    try {
      var smsLines = coupons.map(function(c) {
        var d = Utilities.formatDate(new Date(c.date + 'T12:00:00'), Session.getScriptTimeZone(), 'MMM d');
        var url = config.APP_BASE_URL + '/coupon.html?code=' + encodeURIComponent(c.couponCode);
        return d + ': ' + c.couponCode + ' (' + c.coupons + ')\n' + url;
      }).join('\n');

      var smsBody =
        'Bhog Coupons Confirmed!\n' +
        'Name: ' + name + '\n' +
        smsLines;

      sendTwilioMessage(
        normalizePhone(phone),
        config.TWILIO_PHONE_NUMBER,
        smsBody
      );
      if (config.ENABLE_LOGGING) Logger.log('SMS sent to ' + phone);
    } catch (smsErr) {
      if (config.ENABLE_LOGGING) Logger.log('SMS failed: ' + smsErr.toString());
    }
  }
}

/**
 * Send a message via Twilio API.
 *
 * @param {string} to   - Recipient (e.g. "+1234567890" or "whatsapp:+1234567890")
 * @param {string} from - Sender number
 * @param {string} body - Message text
 * @returns {Object} Twilio API response
 */
function sendTwilioMessage(to, from, body) {
  var config = getConfig();
  var url = 'https://api.twilio.com/2010-04-01/Accounts/' +
    config.TWILIO_ACCOUNT_SID + '/Messages.json';

  var payload = {
    To: to,
    From: from,
    Body: body
  };

  var options = {
    method: 'post',
    payload: payload,
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(
        config.TWILIO_ACCOUNT_SID + ':' + config.TWILIO_AUTH_TOKEN
      )
    },
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var responseCode = response.getResponseCode();

  if (responseCode !== 201) {
    throw new Error('Twilio API error (' + responseCode + '): ' + response.getContentText());
  }

  return JSON.parse(response.getContentText());
}

/**
 * Normalize phone number — ensure it starts with the configured country code.
 *
 * @param {string} phone
 * @returns {string}
 */
function normalizePhone(phone) {
  var config = getConfig();
  var cleaned = phone.replace(/[\s\-()]/g, '');
  if (!cleaned.startsWith('+')) {
    cleaned = config.DEFAULT_COUNTRY_CODE + cleaned;
  }
  return cleaned;
}

// ============================================================
// MEMBERS LOOKUP (Members tab lives in the SAME spreadsheet)
// ============================================================

/**
 * Members list for a date, cross-referenced with that date's registrations.
 *
 * Members tab columns: Name | Phone | Email | Notes | Remarks | RemarksUpdated
 * (Remarks + RemarksUpdated are added automatically the first time a remark is saved.)
 *
 * Each member gets status: 'not_registered' | 'registered' (has coupon, not collected) | 'redeemed'
 */
function getMembersStatus(date) {
  var config = getConfig();
  if (!date) date = getTodayDate_();

  var ss = getSpreadsheet_();
  var memberSheet = ss.getSheetByName(config.MEMBERS_SHEET_NAME || 'Members');
  if (!memberSheet) {
    return { status: 'error', message: 'No "Members" tab found in the spreadsheet. Add one with columns: Name | Phone | Email | Notes.' };
  }

  var nameCol = (config.MEMBERS_NAME_COLUMN || 1) - 1;
  var phoneCol = (config.MEMBERS_PHONE_COLUMN || 2) - 1;
  var last = memberSheet.getLastRow();
  var width = Math.max(6, memberSheet.getLastColumn());
  var memberRows = last >= 2 ? memberSheet.getRange(2, 1, last - 1, width).getValues() : [];

  // Registrations for the date, keyed by last-10-digits of phone
  var regMap = {};
  var dateSheet = ss.getSheetByName(date);
  if (dateSheet) {
    var regRows = readDataRows_(dateSheet, 8);
    for (var j = 0; j < regRows.length; j++) {
      var e = rowToEntry_(regRows[j], date, j + 2);
      var k = last10_(e.phone);
      if (k) regMap[k] = e;
    }
  }

  var members = [];
  var counts = { total: 0, notRegistered: 0, pendingPickup: 0, redeemed: 0, noRemark: 0 };

  for (var i = 0; i < memberRows.length; i++) {
    var row = memberRows[i];
    var phone = phoneToStr_(row[phoneCol]);
    var name = cellToStr_(row[nameCol]).trim();
    if (!phone && !name) continue;

    var reg = regMap[last10_(phone)];
    var status = !reg ? 'not_registered' : (reg.redeemed ? 'redeemed' : 'registered');
    var remarks = cellToStr_(row[4]).trim();

    counts.total++;
    if (status === 'not_registered') counts.notRegistered++;
    else if (status === 'registered') counts.pendingPickup++;
    else counts.redeemed++;
    if (!remarks) counts.noRemark++;

    members.push({
      row: i + 2,                       // row in Members tab (used to save remarks)
      name: name || (reg ? reg.name : ''),
      phone: phone,
      email: cellToStr_(row[2]),
      notes: cellToStr_(row[3]),
      remarks: remarks,
      remarksUpdated: cellToStr_(row[5]),
      status: status,
      plates: reg ? reg.plates : 0,
      couponCode: reg ? reg.couponCode : '',
      redeemedAt: reg ? reg.redeemedAt : ''
    });
  }

  return { status: 'success', date: date, members: members, counts: counts };
}

/**
 * Save a remark (e.g. "Notified", "No answer") against a member so every
 * admin can see it. Members tab: column E = Remarks, column F = RemarksUpdated.
 */
function setMemberRemark(row, phone, remark) {
  var config = getConfig();
  var sheet = getSpreadsheet_().getSheetByName(config.MEMBERS_SHEET_NAME || 'Members');
  if (!sheet) return { status: 'error', message: 'No "Members" tab found' };

  var lock = LockService.getScriptLock();
  try { lock.waitLock(5000); }
  catch (e) { return { status: 'error', message: 'System busy — please try again' }; }

  try {
    var phoneCol = (config.MEMBERS_PHONE_COLUMN || 2);
    var target = last10_(phone);
    var r = parseInt(row, 10);

    // Trust the row number only if that row still has the same phone; otherwise search.
    var ok = r >= 2 && r <= sheet.getLastRow() && last10_(sheet.getRange(r, phoneCol).getValue()) === target;
    if (!ok) {
      r = 0;
      var vals = sheet.getRange(2, phoneCol, Math.max(sheet.getLastRow() - 1, 1), 1).getValues();
      for (var i = 0; i < vals.length; i++) {
        if (last10_(vals[i][0]) === target) { r = i + 2; break; }
      }
    }
    if (!r) return { status: 'error', message: 'Member not found in the Members tab' };

    // Ensure headers exist for the two remark columns
    if (String(sheet.getRange(1, 5).getValue()).trim() !== 'Remarks') {
      sheet.getRange(1, 5, 1, 2).setValues([['Remarks', 'RemarksUpdated']]).setFontWeight('bold');
    }
    var text = String(remark || '').trim().slice(0, 300);
    var stamp = text ? new Date().toISOString() : '';
    sheet.getRange(r, 5, 1, 2).setNumberFormat('@').setValues([[text, stamp]]);

    return { status: 'success', row: r, remarks: text, remarksUpdated: stamp };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// INITIAL SETUP FUNCTION (Run once manually)
// ============================================================

/**
 * Run this function once to set up the spreadsheet.
 * Creates the DateSettings sheet (with MaxTokens column) and
 * one tab per event date.
 *
 * In the Apps Script editor: Run > setupSheet
 */
function setupSheet() {
  var config = getConfig();
  Logger.log('Active Environment: ' + ACTIVE_ENV);

  // Create DateSettings sheet (with Date, Enabled, MaxTokens)
  var dateSheet = getDateSettingsSheet();
  Logger.log('DateSettings sheet is ready!');

  // Create a tab for each event date
  if (config.EVENT_DATES && config.EVENT_DATES.length > 0) {
    for (var i = 0; i < config.EVENT_DATES.length; i++) {
      var sheet = getSheetForDate(config.EVENT_DATES[i]);
      Logger.log('Date tab "' + config.EVENT_DATES[i] + '" is ready.');
    }
  }

  Logger.log('Spreadsheet URL: https://docs.google.com/spreadsheets/d/' + config.SPREADSHEET_ID);

  // Also run validation
  validateConfig();
}