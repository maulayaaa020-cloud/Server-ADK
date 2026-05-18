<?php
ob_start();
session_start();
require_once __DIR__ . '/includes/config.php';
require_once __DIR__ . '/includes/db.php';

$id = (int)($_GET['id'] ?? 0);

// Fallback: jika tidak ada ?id, coba dari session (alur langsung dari hasil)
if (!$id) {
    $relPath = $_SESSION['di_file_output'] ?? '';
    if (!$relPath) { ob_end_clean(); http_response_code(400); exit('Parameter tidak valid.'); }
    $absPath = __DIR__ . '/' . $relPath;
    if (!file_exists($absPath)) { ob_end_clean(); http_response_code(404); exit('File tidak ditemukan.'); }
    $basename = basename($absPath);
    ob_end_clean();
    header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    header('Content-Disposition: attachment; filename="' . $basename . '"');
    header('Content-Length: ' . filesize($absPath));
    header('Cache-Control: private, no-cache');
    readfile($absPath);
    exit;
}

// Mode dari history: cek kepemilikan
try {
    $db   = getDB();
    $stmt = $db->prepare("SELECT * FROM daftar_isi_orders WHERE id = :id LIMIT 1");
    $stmt->execute([':id' => $id]);
    $order = $stmt->fetch();
} catch (Exception $e) {
    ob_end_clean(); http_response_code(500); exit('Kesalahan server.');
}

if (!$order) { ob_end_clean(); http_response_code(404); exit('Order tidak ditemukan.'); }
if (empty($order['file_output'])) { ob_end_clean(); http_response_code(404); exit('File belum tersedia.'); }

$isAdmin    = !empty($_SESSION['adk_admin']);
$sessPhone  = $_SESSION['phone'] ?? $_SESSION['cek_phone'] ?? $_SESSION['di_phone'] ?? null;
$guestToken = $_SESSION['guest_token'] ?? $_SESSION['di_guest_token'] ?? $_COOKIE['adk_guest'] ?? null;

$isOwner = $isAdmin
    || ($sessPhone  && $order['phone']       === $sessPhone)
    || ($guestToken && $order['guest_token'] === $guestToken);

if (!$isOwner) { ob_end_clean(); http_response_code(403); exit('Akses ditolak.'); }

$absPath = __DIR__ . '/' . $order['file_output'];
if (!file_exists($absPath)) { ob_end_clean(); http_response_code(404); exit('File tidak ditemukan di server.'); }

$basename = basename($absPath);
ob_end_clean();
header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
header('Content-Disposition: attachment; filename="' . $basename . '"');
header('Content-Length: ' . filesize($absPath));
header('Cache-Control: private, no-cache');
header('X-Content-Type-Options: nosniff');
readfile($absPath);
exit;
