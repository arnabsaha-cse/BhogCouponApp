// ============================================================
// CONFIGURATION FILE - config.gs
// ============================================================
//
// This file contains all configuration for the Food Distribution
// Coupon System. Switch between NON-PROD and PROD by changing
// the ACTIVE_ENV variable below.
//
// SETUP:
// 1. Create this as a separate file in your Apps Script project
//    (click "+" next to Files > Script > name it "config")
// 2. Fill in your values for each environment
// 3. Set ACTIVE_ENV to 'NON_PROD' for testing, 'PROD' for live event
//
// ============================================================

/**
 * ┌─────────────────────────────────────────────────────────┐
 * │  ACTIVE ENVIRONMENT                                     │
 * │  Change this to switch between NON_PROD and PROD        │
 * │  'NON_PROD' = testing / development                     │
 * │  'PROD'     = live event                                │
 * └─────────────────────────────────────────────────────────┘
 */
const ACTIVE_ENV = 'NON_PROD';


// ============================================================
// NON-PROD (Testing / Development)
// ============================================================
// Use this for testing before the actual event.
// Points to a separate test spreadsheet so your real data
// stays clean. Messaging is disabled by default.
// ============================================================
const NON_PROD = {

  // ---------- Google Sheet ----------
  // Create a SEPARATE test spreadsheet for non-prod
  SPREADSHEET_ID: 'YOUR_TEST_GOOGLE_SHEET_ID_HERE',
  SHEET_NAME: 'Registrations_Test',

  // ---------- Twilio ----------
  // Use Twilio test credentials (from Twilio Console > Account > Test Credentials)
  // Test credentials don't send real messages or cost money
  TWILIO_ACCOUNT_SID: 'YOUR_TWILIO_TEST_ACCOUNT_SID',
  TWILIO_AUTH_TOKEN: 'YOUR_TWILIO_TEST_AUTH_TOKEN',
  TWILIO_PHONE_NUMBER: '+15005550006',                      // Twilio test phone number
  TWILIO_WHATSAPP_NUMBER: 'whatsapp:+14155238886',          // Twilio WhatsApp sandbox

  // ---------- App URLs ----------
  // Local or staging URL
  APP_BASE_URL: 'http://localhost:8080',

  // ---------- Feature Flags ----------
  ENABLE_WHATSAPP: false,                                    // Keep OFF for testing
  ENABLE_SMS: false,                                         // Keep OFF for testing
  ENABLE_DUPLICATE_CHECK: true,                              // Test duplicate detection
  ENABLE_LOGGING: true,                                      // Extra logging for debugging

  // ---------- Event Settings ----------
  EVENT_DATES: [
    '2026-10-01',
    '2026-10-02',
    '2026-10-03',
    '2026-10-04',
    '2026-10-05'
  ],
  MAX_PLATES_PER_REGISTRATION: 5,
  COUPON_PREFIX: 'TEST',                                     // Coupons will be TEST-XXXX (non-prod only)

  // ---------- Members Sheet (external) ----------
  // Google Sheet ID that contains the members list
  // The script reads Name and Phone from this sheet
  MEMBERS_SPREADSHEET_ID: 'YOUR_TEST_MEMBERS_SHEET_ID_HERE',
  MEMBERS_SHEET_NAME: 'Sheet1',                              // Tab name in the members/receipts spreadsheet
  MEMBERS_PHONE_COLUMN: 2,                                   // Column number for phone (1-based, e.g., B=2)
  MEMBERS_NAME_COLUMN: 1,                                    // Column number for name (1-based, e.g., A=1)
  MEMBERS_FILTER_COLUMN: 0,                                  // Column number for purpose/type filter (1-based, e.g., C=3). Set to 0 to skip filtering.
  MEMBERS_FILTER_VALUE: 'Membership',                        // Only rows matching this value are treated as members (case-insensitive)

  // ---------- Admin ----------
  ADMIN_PIN: '0000',                                         // Simple PIN for testing

  // ---------- Phone Defaults ----------
  DEFAULT_COUNTRY_CODE: '+1',                                // Change to your country code
};


// ============================================================
// PROD (Live Event)
// ============================================================
// Use this for the actual food distribution event.
// Points to the real spreadsheet. Messaging is enabled.
// ============================================================
const PROD = {

  // ---------- Google Sheet ----------
  // Your REAL spreadsheet for production data
  SPREADSHEET_ID: 'YOUR_PROD_GOOGLE_SHEET_ID_HERE',
  SHEET_NAME: 'Registrations',

  // ---------- Twilio ----------
  // Use REAL Twilio credentials (from Twilio Console > Account > Live Credentials)
  TWILIO_ACCOUNT_SID: 'YOUR_TWILIO_LIVE_ACCOUNT_SID',
  TWILIO_AUTH_TOKEN: 'YOUR_TWILIO_LIVE_AUTH_TOKEN',
  TWILIO_PHONE_NUMBER: 'YOUR_TWILIO_LIVE_PHONE_NUMBER',     // e.g., '+1234567890'
  TWILIO_WHATSAPP_NUMBER: 'whatsapp:+14155238886',          // Your approved WhatsApp number

  // ---------- App URLs ----------
  // Your live hosted URL
  APP_BASE_URL: 'https://yourusername.github.io/FoodCouponApp',

  // ---------- Feature Flags ----------
  ENABLE_WHATSAPP: true,                                     // ON for live event
  ENABLE_SMS: true,                                          // ON as fallback
  ENABLE_DUPLICATE_CHECK: true,                              // Prevent double registration
  ENABLE_LOGGING: false,                                     // Less logging in prod

  // ---------- Event Settings ----------
  EVENT_DATES: [
    '2026-10-01',
    '2026-10-02',
    '2026-10-03',
    '2026-10-04',
    '2026-10-05'
  ],
  MAX_PLATES_PER_REGISTRATION: 5,
  COUPON_PREFIX: 'BHOG',                                     // Coupons will be BHOG-XXXX

  // ---------- Members Sheet (external) ----------
  // Google Sheet ID that contains the members list
  MEMBERS_SPREADSHEET_ID: 'YOUR_MEMBERS_SHEET_ID_HERE',
  MEMBERS_SHEET_NAME: 'Sheet1',                              // Tab name in the members/receipts spreadsheet
  MEMBERS_PHONE_COLUMN: 2,                                   // Column number for phone (1-based, e.g., B=2)
  MEMBERS_NAME_COLUMN: 1,                                    // Column number for name (1-based, e.g., A=1)
  MEMBERS_FILTER_COLUMN: 0,                                  // Column number for purpose/type filter (1-based, e.g., C=3). Set to 0 to skip filtering.
  MEMBERS_FILTER_VALUE: 'Membership',                        // Only rows matching this value are treated as members (case-insensitive)

  // ---------- Admin ----------
  ADMIN_PIN: '1234',                                         // CHANGE THIS to a secure PIN!

  // ---------- Phone Defaults ----------
  DEFAULT_COUNTRY_CODE: '+1',                                // Change to your country code
};


// ============================================================
// CONFIG RESOLVER - Do not modify below this line
// ============================================================

/**
 * Returns the active configuration based on ACTIVE_ENV.
 * All other files use getConfig() to read settings.
 */
function getConfig() {
  switch (ACTIVE_ENV) {
    case 'PROD':
      return PROD;
    case 'NON_PROD':
      return NON_PROD;
    default:
      Logger.log('⚠️ Unknown ACTIVE_ENV: "' + ACTIVE_ENV + '". Defaulting to NON_PROD.');
      return NON_PROD;
  }
}

/**
 * Utility: Log the active environment (run manually to verify)
 */
function logActiveConfig() {
  const cfg = getConfig();
  Logger.log('============================================');
  Logger.log('Active Environment: ' + ACTIVE_ENV);
  Logger.log('Spreadsheet ID:    ' + cfg.SPREADSHEET_ID);
  Logger.log('Sheet Name:        ' + cfg.SHEET_NAME);
  Logger.log('App Base URL:      ' + cfg.APP_BASE_URL);
  Logger.log('Coupon Prefix:     ' + cfg.COUPON_PREFIX);
  Logger.log('WhatsApp Enabled:  ' + cfg.ENABLE_WHATSAPP);
  Logger.log('SMS Enabled:       ' + cfg.ENABLE_SMS);
  Logger.log('Duplicate Check:   ' + cfg.ENABLE_DUPLICATE_CHECK);
  Logger.log('Logging Enabled:   ' + cfg.ENABLE_LOGGING);
  Logger.log('Admin PIN:         ' + cfg.ADMIN_PIN);
  Logger.log('============================================');
}

/**
 * Utility: Validate that required config values are set
 * Run this before deploying to catch missing values.
 */
function validateConfig() {
  const cfg = getConfig();
  const errors = [];

  // Required fields
  if (!cfg.SPREADSHEET_ID || cfg.SPREADSHEET_ID.includes('YOUR_'))
    errors.push('SPREADSHEET_ID is not set');

  if (cfg.ENABLE_SMS || cfg.ENABLE_WHATSAPP) {
    if (!cfg.TWILIO_ACCOUNT_SID || cfg.TWILIO_ACCOUNT_SID.includes('YOUR_'))
      errors.push('TWILIO_ACCOUNT_SID is not set');
    if (!cfg.TWILIO_AUTH_TOKEN || cfg.TWILIO_AUTH_TOKEN.includes('YOUR_'))
      errors.push('TWILIO_AUTH_TOKEN is not set');
    if (!cfg.TWILIO_PHONE_NUMBER || cfg.TWILIO_PHONE_NUMBER.includes('YOUR_'))
      errors.push('TWILIO_PHONE_NUMBER is not set');
  }

  if (!cfg.APP_BASE_URL || cfg.APP_BASE_URL === 'http://localhost:8080') {
    if (ACTIVE_ENV === 'PROD')
      errors.push('APP_BASE_URL is still set to localhost (update for PROD)');
  }

  if (cfg.ADMIN_PIN === '0000' && ACTIVE_ENV === 'PROD')
    errors.push('ADMIN_PIN is still the default "0000" (change for PROD)');

  if (cfg.ADMIN_PIN === '1234' && ACTIVE_ENV === 'PROD')
    errors.push('ADMIN_PIN is "1234" — consider a more secure PIN for PROD');

  if (!cfg.EVENT_DATES || cfg.EVENT_DATES.length === 0)
    errors.push('EVENT_DATES is empty');

  // Report
  if (errors.length === 0) {
    Logger.log('✅ Configuration is valid for ' + ACTIVE_ENV);
  } else {
    Logger.log('❌ Configuration errors for ' + ACTIVE_ENV + ':');
    errors.forEach(function(err, i) {
      Logger.log('   ' + (i + 1) + '. ' + err);
    });
  }

  return errors;
}
