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
    //   GitHub Pages:  'https://yourusername.github.io/FoodCouponApp'
    //   Local testing: '' (empty string — uses relative paths)
    //
    BASE_URL: '',

    // ──────────────────────────────────────────────────────────
    // QR Code Library
    // ──────────────────────────────────────────────────────────
    // Path to the qrcode.js library used by index.html, coupon.html
    // and qrcode.html to render QR codes.
    // Using cdnjs's mirror of the classic davidshimjs/qrcodejs
    // library (API: `new QRCode(element, options)`), since it's a
    // stable, permanently-hosted file — unlike the npm "qrcode"
    // package, whose precompiled browser bundle is no longer
    // published for recent versions.
    // If you'd rather self-host: download the file below, put it
    // in this same frontend/ folder as e.g. qrcode.min.js, and
    // point this at 'qrcode.min.js' instead.
    //
    QRCODE_LIB_URL: 'https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js',

    // ──────────────────────────────────────────────────────────
    // Event Dates
    // ──────────────────────────────────────────────────────────
    // Your 5 distribution dates in YYYY-MM-DD format.
    // These populate the date dropdown on the registration form
    // and the date filter on the admin page.
    //
    EVENT_DATES: [
        '2026-10-16',
        '2026-10-17',
        '2026-10-18',
        '2026-10-19',
        '2026-10-20'
    ],

    // ──────────────────────────────────────────────────────────
    // Event Logo
    // ──────────────────────────────────────────────────────────
    // URL or relative path to your event logo image.
    // Displayed at the top of the registration page.
    // Recommended size: 80-120px height, PNG or JPG.
    // Set to '' to show the 🪷 emoji fallback.
    //
    EVENT_LOGO_URL: 'agomoni_small_logo.png',

    // ──────────────────────────────────────────────────────────
    // Registration Settings
    // ──────────────────────────────────────────────────────────
    // Max coupons (plates) per date per registration
    MAX_PLATES: 5,

    // ──────────────────────────────────────────────────────────
    // Admin Settings
    // ──────────────────────────────────────────────────────────
    // 4-digit PIN to access the admin page.
    // CHANGE THIS before going live!
    //
    ADMIN_PIN: '1234'
};