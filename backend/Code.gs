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
        result = getSummary();
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

/**
 * Opens the spreadsheet by config ID. Cached per execution via CacheService
 * is not needed since SpreadsheetApp handles its own per-execution cache.
 */
function getSpreadsheet_() {
  var config = getConfig();
  return SpreadsheetApp.openById(config.SPREADSHEET_ID);
}

/**
 * Get or create a sheet tab for a specific event date.
 * Tab name = date string (e.g. "2026-10-01").
 *
 * @param {string} date - Date string in YYYY-MM-DD format
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheetForDate(date) {
  if (!date) throw new Error('getSheetForDate: date is required');

  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(date);

  if (!sheet) {
    sheet = ss.insertSheet(date);
    sheet.appendRow(DATE_TAB_HEADERS);
    sheet.getRange(1, 1, 1, DATE_TAB_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
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
  if (!name || !phone) {
    return { status: 'error', message: 'Name and phone are required' };
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
  var fullDates = [];         // dates that exceeded token limit
  var timestamp = new Date().toISOString();

  // Load date settings once for limit checks
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

    // --- Token limit check ---
    if (settings && settings.maxTokens > 0) {
      var issuedCount = getIssuedCountForDate_(date);
      if (issuedCount + requestedPlates > settings.maxTokens) {
        fullDates.push({ date: date, reason: 'full' });
        continue;
      }
    }

    // --- Check for existing registration FIRST (same phone + date) ---
    // Must be before token limit check so we can correctly calculate net new plates.
    var sheet = getSheetForDate(date);
    var existing = findByPhoneAndDateInSheet_(sheet, phone);

    if (existing) {
      // UPDATE: existing plates already counted in issuedCount; net change is the delta.
      // We allow the update even if limit would otherwise block it (it's the same person adjusting).
      sheet.getRange(existing.row, COL.PLATES).setValue(requestedPlates);
      updatedCoupons.push({
        date: date,
        coupons: requestedPlates,
        couponCode: existing.couponCode,
        status: 'updated'
      });
      continue;
    }

    // --- Token limit check (new registrations only) ---
    if (settings && settings.maxTokens > 0) {
      var issuedCount = getIssuedCountForDate_(date);
      if (issuedCount + requestedPlates > settings.maxTokens) {
        fullDates.push({ date: date, reason: 'full' });
        continue;
      }
    }

    // --- New registration: generate coupon & write row ---
    var couponCode = generateCouponCode();
    var newRow = [timestamp, name, phone, requestedPlates, couponCode, 'No', '', 'Online'];
    sheet.appendRow(newRow);

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
 *
 * @param {Object} data - { name, phone, plates }
 * @returns {Object} result
 */
function registerWalkin(data) {
  var config = getConfig();
  var name = (data.name || '').toString().trim();
  var phone = (data.phone || '').toString().trim();
  var plates = parseInt(data.plates, 10) || 1;

  if (!name || !phone) {
    return { status: 'error', message: 'Name and phone are required' };
  }

  if (plates > config.MAX_PLATES_PER_REGISTRATION) {
    plates = config.MAX_PLATES_PER_REGISTRATION;
  }

  var date = getTodayDate_();
  var now = new Date().toISOString();

  // --- Token limit check ---
  var dateSettingsMap = getDateSettingsMap_();
  var settings = dateSettingsMap[date];
  if (settings && settings.maxTokens > 0) {
    var issuedCount = getIssuedCountForDate_(date);
    if (issuedCount + plates > settings.maxTokens) {
      return {
        status: 'error',
        message: 'Today\'s coupon limit has been reached',
        date: date
      };
    }
  }

  var sheet = getSheetForDate(date);
  var existing = findByPhoneAndDateInSheet_(sheet, phone);

  var couponCode;

  if (existing) {
    // Update existing walk-in: change plate count, keep code, ensure redeemed
    couponCode = existing.couponCode;
    sheet.getRange(existing.row, COL.PLATES).setValue(plates);
    sheet.getRange(existing.row, COL.REDEEMED).setValue('Yes');
    sheet.getRange(existing.row, COL.REDEEMED_AT).setValue(now);
  } else {
    // New walk-in registration — immediately redeemed
    couponCode = generateCouponCode();
    var newRow = [now, name, phone, plates, couponCode, 'Yes', now, 'Walk-in'];
    sheet.appendRow(newRow);
  }

  // NO messaging for walk-ins

  return {
    status: 'success',
    couponCode: couponCode,
    name: name,
    phone: phone,
    date: date,
    plates: plates
  };
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
 * Read all data rows from a date-tab sheet and return as entry objects.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} date - The date this tab represents
 * @returns {Object[]}
 */
function readEntriesFromSheet_(sheet, date) {
  var data = sheet.getDataRange().getValues();
  var entries = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    entries.push({
      timestamp: row[COL.TIMESTAMP - 1],
      name: row[COL.NAME - 1],
      phone: row[COL.PHONE - 1],
      date: date,
      plates: row[COL.PLATES - 1],
      couponCode: row[COL.COUPON_CODE - 1],
      redeemed: row[COL.REDEEMED - 1] === 'Yes',
      redeemedAt: row[COL.REDEEMED_AT - 1] || '',
      source: row[COL.SOURCE - 1] || 'Online'
    });
  }
  return entries;
}

/**
 * Get a summary of all event dates — registrations, walk-ins,
 * coupon counts, redeemed, pending, and token limits.
 *
 * @returns {Object}
 */
function getSummary() {
  var today = getTodayDate_();
  var dateSettingsMap = getDateSettingsMap_();
  var dateSheets = getAllDateSheets_();
  var dateSummaries = [];
  var todaySummary = null;

  for (var i = 0; i < dateSheets.length; i++) {
    var sheet = dateSheets[i];
    var tabDate = sheet.getName();
    var data = sheet.getDataRange().getValues();

    var registered = 0;   // Online source count
    var walkins = 0;       // Walk-in source count
    var totalCoupons = 0;  // Sum of Plates
    var redeemed = 0;      // Redeemed = Yes count

    for (var r = 1; r < data.length; r++) {
      var source = data[r][COL.SOURCE - 1] || 'Online';
      var plates = parseInt(data[r][COL.PLATES - 1], 10) || 0;
      var isRedeemed = data[r][COL.REDEEMED - 1] === 'Yes';

      if (source === 'Online') {
        registered++;
      } else if (source === 'Walk-in') {
        walkins++;
      }
      totalCoupons += plates;
      if (isRedeemed) redeemed++;
    }

    var totalRegistrations = registered + walkins;
    var pending = totalRegistrations - redeemed;
    var maxTokens = (dateSettingsMap[tabDate] && dateSettingsMap[tabDate].maxTokens) || 0;

    // pending = people who have NOT redeemed (headcount, not plate count)
    // This is consistent: registered/walkins/redeemed/pending are all headcounts.
    // totalCoupons remains the plate (food portion) count.
    var summaryObj = {
      date: tabDate,
      registered: registered,
      walkins: walkins,
      totalRegistrations: totalRegistrations,
      totalCoupons: totalCoupons,
      redeemed: redeemed,
      pending: pending,           // headcount of unredeemed registrations
      maxTokens: maxTokens
    };

    dateSummaries.push(summaryObj);

    if (tabDate === today) {
      todaySummary = Object.assign({}, summaryObj, {
        remaining: (maxTokens > 0) ? Math.max(0, maxTokens - totalCoupons) : undefined
      });
    }
  }

  return {
    status: 'success',
    days: dateSummaries,    // FIX: was 'dates' — admin.html reads 'days'
    today: todaySummary
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
 * Search a single sheet for a coupon code.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} code
 * @param {string} date - The date this tab represents
 * @returns {Object|null}
 */
function findCodeInSheet_(sheet, code, date) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL.COUPON_CODE - 1]).toUpperCase() === code) {
      return {
        row: i + 1, // 1-based sheet row
        date: date,
        timestamp: data[i][COL.TIMESTAMP - 1],
        name: data[i][COL.NAME - 1],
        phone: data[i][COL.PHONE - 1],
        plates: data[i][COL.PLATES - 1],
        couponCode: String(data[i][COL.COUPON_CODE - 1]),
        redeemed: data[i][COL.REDEEMED - 1] === 'Yes',
        redeemedAt: data[i][COL.REDEEMED_AT - 1] || '',
        source: data[i][COL.SOURCE - 1] || 'Online'
      };
    }
  }
  return null;
}

/**
 * Find a registration by phone number within a specific date's sheet.
 * Used for duplicate / update checking.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} phone
 * @returns {Object|null}
 */
function findByPhoneAndDateInSheet_(sheet, phone) {
  var data = sheet.getDataRange().getValues();
  var normalizedPhone = phone.replace(/[\s\-()]/g, '');

  for (var i = 1; i < data.length; i++) {
    var rowPhone = String(data[i][COL.PHONE - 1]).replace(/[\s\-()]/g, '');
    if (rowPhone === normalizedPhone) {
      return {
        row: i + 1,
        couponCode: String(data[i][COL.COUPON_CODE - 1]),
        plates: data[i][COL.PLATES - 1],
        redeemed: data[i][COL.REDEEMED - 1] === 'Yes'
      };
    }
  }
  return null;
}

/**
 * Generate a unique coupon code like BHOG-A3X7 (or TEST-A3X7 in non-prod).
 * Attempts up to 100 times to avoid collisions.
 *
 * @returns {string}
 */
function generateCouponCode() {
  var config = getConfig();
  var prefix = config.COUPON_PREFIX || 'FOOD';
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // Removed confusing chars (0,O,1,I)
  var code;
  var attempts = 0;

  do {
    code = prefix + '-';
    for (var j = 0; j < 4; j++) {
      code += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    attempts++;
  } while (findByCode(code) && attempts < 100);

  return code;
}

// ============================================================
// DAILY TOKEN LIMIT HELPERS
// ============================================================

/**
 * Get total issued plates (sum of Plates column) for a given date.
 *
 * @param {string} date
 * @returns {number}
 */
function getIssuedCountForDate_(date) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(date);
  if (!sheet) return 0;

  var data = sheet.getDataRange().getValues();
  var total = 0;
  for (var i = 1; i < data.length; i++) {
    total += parseInt(data[i][COL.PLATES - 1], 10) || 0;
  }
  return total;
}

/**
 * Build a map of date -> { enabled, maxTokens } from the DateSettings sheet.
 * Used internally for fast lookups.
 *
 * @returns {Object} map keyed by date string
 */
function getDateSettingsMap_() {
  var sheet = getDateSettingsSheet();
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var dateStr = String(data[i][0]);
    map[dateStr] = {
      enabled: data[i][1] === 'Yes',
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
 * Get or create the DateSettings sheet.
 * Pre-populated with EVENT_DATES from config, all enabled, MaxTokens=0 (unlimited).
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
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

    // Pre-populate with event dates, all enabled, unlimited tokens
    if (config.EVENT_DATES && config.EVENT_DATES.length > 0) {
      var rows = config.EVENT_DATES.map(function(dateStr) {
        return [dateStr, 'Yes', 0];
      });
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
  var sheet = getDateSettingsSheet();
  var data = sheet.getDataRange().getValues();
  var settings = [];

  for (var i = 1; i < data.length; i++) {
    var dateStr = String(data[i][0]);
    var enabled = data[i][1] === 'Yes';
    var maxTokens = parseInt(data[i][2], 10) || 0;
    var issuedCount = getIssuedCountForDate_(dateStr);
    var isFull = (maxTokens > 0) && (issuedCount >= maxTokens);

    settings.push({
      date: dateStr,
      enabled: enabled,
      maxTokens: maxTokens,
      issuedCount: issuedCount,
      isFull: isFull
    });
  }

  return {
    status: 'success',
    dates: settings
  };
}

/**
 * Toggle a date's enabled/disabled status.
 *
 * @param {string} date    - Date string (YYYY-MM-DD)
 * @param {string} enabled - 'true' or 'false'
 * @returns {Object}
 */
function toggleDate(date, enabled) {
  if (!date) return { status: 'error', message: 'Date is required' };

  var sheet = getDateSettingsSheet();
  var data = sheet.getDataRange().getValues();
  var newValue = (enabled === 'true') ? 'Yes' : 'No';

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === date) {
      sheet.getRange(i + 1, 2).setValue(newValue);
      return {
        status: 'success',
        date: date,
        enabled: newValue === 'Yes'
      };
    }
  }

  // Date not found in settings — add it
  sheet.appendRow([date, newValue, 0]);
  return {
    status: 'success',
    date: date,
    enabled: newValue === 'Yes'
  };
}

/**
 * Set the maximum token (coupon) limit for a specific date.
 *
 * @param {string} date      - Date string (YYYY-MM-DD)
 * @param {string|number} maxTokens - Maximum coupons allowed (0 = unlimited)
 * @returns {Object}
 */
function setDateLimit(date, maxTokens) {
  if (!date) return { status: 'error', message: 'Date is required' };

  var limit = parseInt(maxTokens, 10);
  if (isNaN(limit) || limit < 0) {
    return { status: 'error', message: 'maxTokens must be a non-negative integer' };
  }

  var sheet = getDateSettingsSheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === date) {
      sheet.getRange(i + 1, 3).setValue(limit);
      return {
        status: 'success',
        date: date,
        maxTokens: limit
      };
    }
  }

  // Date not found — add it (enabled by default)
  sheet.appendRow([date, 'Yes', limit]);
  return {
    status: 'success',
    date: date,
    maxTokens: limit
  };
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
// MEMBERS LOOKUP (cross-references external members sheet)
// ============================================================

/**
 * Get registered members who have NOT redeemed their coupons for a given date.
 * Cross-references the external members sheet with the registration data.
 *
 * @param {string} date - Date to check (YYYY-MM-DD). Defaults to today.
 * @returns {Object} { status, date, pending: [{name, phone, plates, couponCode}], totalMembers, totalRegistered, totalPending }
 */
function getMembersStatus(date) {
  var config = getConfig();

  // Default to today
  if (!date) {
    var today = new Date();
    date = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  // Validate members config
  if (!config.MEMBERS_SPREADSHEET_ID || config.MEMBERS_SPREADSHEET_ID.indexOf('YOUR_') === 0) {
    return { status: 'error', message: 'Members spreadsheet is not configured. Update MEMBERS_SPREADSHEET_ID in config.gs.' };
  }

  // Read members sheet
  var memberPhones = {};
  try {
    var memberSS = SpreadsheetApp.openById(config.MEMBERS_SPREADSHEET_ID);
    var memberSheet = memberSS.getSheetByName(config.MEMBERS_SHEET_NAME || 'Sheet1');
    if (!memberSheet) {
      return { status: 'error', message: 'Members sheet tab "' + (config.MEMBERS_SHEET_NAME || 'Sheet1') + '" not found.' };
    }

    var memberData = memberSheet.getDataRange().getValues();
    var nameCol = (config.MEMBERS_NAME_COLUMN || 1) - 1;    // Convert to 0-based
    var phoneCol = (config.MEMBERS_PHONE_COLUMN || 2) - 1;  // Convert to 0-based
    var filterCol = (config.MEMBERS_FILTER_COLUMN || 0) - 1; // -1 means no filter
    var filterVal = (config.MEMBERS_FILTER_VALUE || 'Membership').toLowerCase();

    for (var i = 1; i < memberData.length; i++) { // Skip header row
      // Apply purpose/type filter if configured
      if (filterCol >= 0) {
        var cellVal = String(memberData[i][filterCol] || '').trim().toLowerCase();
        if (cellVal !== filterVal) continue; // Skip non-membership rows
      }

      var phone = String(memberData[i][phoneCol] || '').replace(/[\s\-()]/g, '');
      var name = String(memberData[i][nameCol] || '').trim();
      if (phone) {
        memberPhones[phone] = name;
      }
    }
  } catch (err) {
    return { status: 'error', message: 'Could not read members sheet: ' + err.toString() };
  }

  var totalMembers = Object.keys(memberPhones).length;

  // Read today's registrations
  var ss = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  var dateSheet = ss.getSheetByName(date);

  if (!dateSheet) {
    // No registrations for this date — all members are unregistered
    return {
      status: 'success',
      date: date,
      pending: [],
      totalMembers: totalMembers,
      totalRegistered: 0,
      totalPending: 0
    };
  }

  var regData = dateSheet.getDataRange().getValues();
  // Columns in date tab: Timestamp(0), Name(1), Phone(2), Plates(3), CouponCode(4), Redeemed(5), RedeemedAt(6), Source(7)

  // Build a map of registered members for this date
  var registeredMap = {};
  for (var j = 1; j < regData.length; j++) {
    var regPhone = String(regData[j][2] || '').replace(/[\s\-()]/g, '');
    registeredMap[regPhone] = {
      name: regData[j][1],
      phone: regData[j][2],
      plates: regData[j][3],
      couponCode: regData[j][4],
      redeemed: regData[j][5] === 'Yes'
    };
  }

  // Find members who registered but have NOT redeemed
  var pendingList = [];
  var totalRegistered = 0;

  var memberPhoneKeys = Object.keys(memberPhones);
  for (var k = 0; k < memberPhoneKeys.length; k++) {
    var mPhone = memberPhoneKeys[k];
    var reg = registeredMap[mPhone];

    if (reg) {
      totalRegistered++;
      if (!reg.redeemed) {
        pendingList.push({
          name: reg.name || memberPhones[mPhone],
          phone: reg.phone,
          plates: reg.plates,
          couponCode: reg.couponCode
        });
      }
    }
  }

  return {
    status: 'success',
    date: date,
    pending: pendingList,
    totalMembers: totalMembers,
    totalRegistered: totalRegistered,
    totalPending: pendingList.length
  };
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