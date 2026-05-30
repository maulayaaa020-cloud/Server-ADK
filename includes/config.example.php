<?php
// Salin file ini ke config.php dan isi dengan nilai yang sesuai.
// JANGAN commit config.php ke repository.

define('MIDTRANS_SERVER_KEY', 'Mid-server-xxx');
define('MIDTRANS_CLIENT_KEY', 'Mid-client-xxx');

// Preview signing key — generate dengan: bin2hex(random_bytes(32))
define('PREVIEW_SECRET_KEY', 'ganti-dengan-random-64-char-hex');

define('DB_HOST', 'localhost');
define('DB_USER', 'root');
define('DB_PASS', '');
define('DB_NAME', 'adk_photocopy');

// Resend Email Gateway (https://resend.com)
define('RESEND_API_KEY', 're_xxxxxxxxxxxxxxxxxxxxxx');
define('RESEND_FROM',    'ADK Photocopy <noreply@domain-kamu.com>');

// Admin dashboard — generate hash: password_hash('password_kamu', PASSWORD_BCRYPT, ['cost'=>12])
define('ADMIN_EMAIL',         'email@kamu.com');
define('ADMIN_PASSWORD_HASH', '$2y$12$...');

// URL aplikasi
// define('APP_URL', 'http://localhost/adk'); // lokal
define('APP_URL',   'https://adkivia.com');   // production
// Base path: '' di VPS/production (domain root), '/adk' di localhost XAMPP
define('BASE_PATH', '');

// Path Python
define('PYTHON_EXE',        'C:\\path\\to\\python.exe');
define('PYTHON_SCRIPT_DIR', 'C:\\path\\to\\adk\\python');

define('CACERT_PATH', __DIR__ . '/../vendor/midtrans/midtrans-php/data/cacert.pem');
