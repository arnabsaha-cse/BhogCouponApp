
    // ============================================================
// FRONTEND CONFIGURATION - config.js
// ============================================================
// This is the ONLY file you need to update for the frontend.
// All HTML pages (index.html, coupon.html, admin.html) read from here.
//
// After updating values, just refresh the browser — no rebuild needed.
// ============================================================

const CONFIG = {

    // ──────────────────────────────────────────────────────────
    // Google Apps Script Web App URL
    // ──────────────────────────────────────────────────────────
    // Get this after deploying your Apps Script:
    //   Apps Script Editor > Deploy > New deployment > Web app
    //   Copy the URL (looks like: https://script.google.com/macros/s/AKfycb.../exec)
    //
    APPS_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbwa0gqREep3ZzsC2pyT98cdTw6Q7tDZuFI6L7JZn2h2cE_LOlFQwDs8ET08t23Gk8weTg/exec',

    // ──────────────────────────────────────────────────────────
    // Base URL where this app is hosted
    // ──────────────────────────────────────────────────────────
    // Used to generate coupon page links in QR codes.
    // Examples:
    //   GitHub Pages:  'https://arnabsaha-cse.github.io/BhogCouponApp'
    //   Local testing: '' (empty string — uses relative paths)
    //
    BASE_URL: '',

    // ──────────────────────────────────────────────────────────
    // Event Dates
    // ──────────────────────────────────────────────────────────
    // Your 5 distribution dates in YYYY-MM-DD format.
    // These populate the date dropdown on the registration form
    // and the date filter on the admin page.
    //
    EVENT_DATES: [
        '2026-10-01',
        '2026-10-02',
        '2026-10-03',
        '2026-10-04',
        '2026-10-05'
    ],

    // ──────────────────────────────────────────────────────────
    // Registration Settings
    // ──────────────────────────────────────────────────────────
    MAX_PLATES: 5,

    // ──────────────────────────────────────────────────────────
    // Admin Settings
    // ──────────────────────────────────────────────────────────
    // 4-digit PIN to access the admin page.
    // CHANGE THIS before going live!
    //
    ADMIN_PIN: '1234'
};
