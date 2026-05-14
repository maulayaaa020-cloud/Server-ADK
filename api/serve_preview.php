<?php
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

$db_id   = (int)($_GET['id']  ?? 0);
$expires = (int)($_GET['t']   ?? 0);
$sig     = $_GET['sig'] ?? '';

if (!$db_id || !$expires || !$sig) { http_response_code(400); exit; }

// Token sudah expired
if (time() > $expires) { http_response_code(403); exit; }

// Verifikasi HMAC — pastikan token tidak dipalsukan
$expected = hash_hmac('sha256', $db_id . ':' . $expires, DOKU_SECRET_KEY);
if (!hash_equals($expected, $sig)) { http_response_code(403); exit; }

try {
    $db   = getDB();
    $stmt = $db->prepare("SELECT file_output FROM orders WHERE id = ?");
    $stmt->execute([$db_id]);
    $order = $stmt->fetch();
} catch (Exception $e) { http_response_code(500); exit; }

if (!$order || empty($order['file_output'])) { http_response_code(404); exit; }

$filePath = __DIR__ . '/../' . $order['file_output'];
if (!file_exists($filePath)) { http_response_code(404); exit; }

header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
header('Content-Disposition: inline; filename="preview.docx"');
header('Content-Length: ' . filesize($filePath));
header('Cache-Control: private, no-store');
readfile($filePath);
