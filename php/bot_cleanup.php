<?php
/**
 * Jalankan via cron: php /var/www/html/adk/php/bot_cleanup.php
 * Hapus file bot > 7 hari + token expired
 */

define('MAX_AGE_SECONDS', 7 * 24 * 60 * 60); // 7 hari

require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';
require_once __DIR__ . '/../includes/log.php';

$now      = time();
$deleted  = 0;
$errors   = 0;

$dirs = [
    __DIR__ . '/../uploads/bot/',
    __DIR__ . '/../uploads/bot/hasil/',
];

foreach ($dirs as $dir) {
    if (!is_dir($dir)) continue;
    foreach (glob($dir . '*') as $file) {
        if (!is_file($file)) continue;
        if (($now - filemtime($file)) >= MAX_AGE_SECONDS) {
            if (unlink($file)) {
                $deleted++;
            } else {
                $errors++;
            }
        }
    }
}

// Hapus token expired dari DB
try {
    $db   = getDB();
    $rows = $db->exec("DELETE FROM bot_login_tokens WHERE expires_at < NOW()");
} catch (Exception $e) {
    $rows = 0;
}

adk_log('cleanup', 'Bot cleanup selesai', [
    'files_deleted'   => $deleted,
    'files_error'     => $errors,
    'tokens_deleted'  => $rows,
]);

echo "[" . date('Y-m-d H:i:s') . "] Cleanup: {$deleted} file dihapus, {$rows} token dihapus\n";
