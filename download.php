<?php
ob_start();
session_start();
require_once __DIR__ . '/includes/config.php';
require_once __DIR__ . '/includes/db.php';

$dbId = (int)($_GET['id'] ?? 0);

if (!$dbId) {
    ob_end_clean();
    http_response_code(400);
    exit('Parameter tidak valid.');
}

try {
    $db   = getDB();
    $stmt = $db->prepare("SELECT * FROM orders WHERE id = :id LIMIT 1");
    $stmt->execute([':id' => $dbId]);
    $order = $stmt->fetch();
} catch (Exception $e) {
    http_response_code(500);
    exit('Kesalahan server.');
}

if (!$order) {
    ob_end_clean();
    http_response_code(404);
    exit('Order tidak ditemukan.');
}

// Hanya order yang sudah lunas yang bisa diunduh
if ($order['status'] !== 'paid') {
    ob_end_clean();
    http_response_code(403);
    exit('File hanya tersedia setelah pembayaran selesai.');
}

// Validasi kepemilikan: admin, pemilik email, atau pemilik guest_token
$isAdmin    = !empty($_SESSION['adk_admin']);
$sessEmail  = $_SESSION['email'] ?? $_SESSION['cek_email'] ?? null;
$guestToken = $_SESSION['guest_token'] ?? $_COOKIE['adk_guest'] ?? null;

$isOwner = $isAdmin
    || ($sessEmail  && $order['phone']       === $sessEmail)
    || ($guestToken && $order['guest_token'] === $guestToken);

if (!$isOwner) {
    ob_end_clean();
    http_response_code(403);
    exit('Akses ditolak.');
}

$relPath = $order['file_output'] ?? '';

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
