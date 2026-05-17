<?php
ob_start();
session_start();
require_once __DIR__ . '/includes/config.php';

// Ambil path file dari session
$relPath = $_SESSION['di_file_output'] ?? '';

if (!$relPath) {
    ob_end_clean();
    http_response_code(404);
    exit('File tidak tersedia.');
}

$absPath = __DIR__ . '/' . $relPath;
if (!file_exists($absPath)) {
    ob_end_clean();
    http_response_code(404);
    exit('File tidak ditemukan di server.');
}

$basename = basename($absPath);

ob_end_clean();
header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
header('Content-Disposition: attachment; filename="' . $basename . '"');
header('Content-Length: ' . filesize($absPath));
header('Cache-Control: private, no-cache');
header('X-Content-Type-Options: nosniff');

readfile($absPath);
exit;
