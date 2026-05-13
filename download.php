<?php
session_start();
require_once __DIR__ . '/includes/config.php';
require_once __DIR__ . '/includes/db.php';

$dbId = (int)($_GET['id']   ?? 0);
$type = $_GET['type'] ?? '';

if (!$dbId || !in_array($type, ['docx', 'pdf'])) {
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
    http_response_code(404);
    exit('Order tidak ditemukan.');
}

// Hanya order yang sudah lunas yang bisa diunduh
if ($order['status'] !== 'paid') {
    http_response_code(403);
    exit('File hanya tersedia setelah pembayaran selesai.');
}

// Validasi kepemilikan: admin, pemilik nomor HP, atau pemilik guest_token
$isAdmin    = !empty($_SESSION['adk_admin']);
$sessPhone  = $_SESSION['phone'] ?? $_SESSION['cek_phone'] ?? null;
$guestToken = $_SESSION['guest_token'] ?? $_COOKIE['adk_guest'] ?? null;

$isOwner = $isAdmin
    || ($sessPhone  && $order['phone']       === $sessPhone)
    || ($guestToken && $order['guest_token'] === $guestToken);

if (!$isOwner) {
    http_response_code(403);
    exit('Akses ditolak.');
}

// Pilih file sesuai tipe
$relPath = ($type === 'pdf') ? ($order['file_output_pdf'] ?? '') : ($order['file_output'] ?? '');

if (!$relPath) {
    http_response_code(404);
    exit('File tidak tersedia.');
}

$absPath = __DIR__ . '/' . $relPath;
if (!file_exists($absPath)) {
    http_response_code(404);
    exit('File tidak ditemukan di server.');
}

$basename = basename($absPath);
$mime     = ($type === 'pdf') ? 'application/pdf'
          : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

header('Content-Type: '         . $mime);
header('Content-Disposition: attachment; filename="' . $basename . '"');
header('Content-Length: '       . filesize($absPath));
header('Cache-Control: private, no-cache');
header('X-Content-Type-Options: nosniff');

readfile($absPath);
exit;
