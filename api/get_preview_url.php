<?php
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';
header('Content-Type: application/json');

$isAdmin    = !empty($_SESSION['adk_admin']);
$email      = $_SESSION['email'] ?? $_SESSION['cek_email'] ?? null;
$guestToken = $_SESSION['guest_token'] ?? $_COOKIE['adk_guest'] ?? null;

if (!$isAdmin && !$email && !$guestToken) {
    echo json_encode(['ok' => false, 'error' => 'Unauthorized']); exit;
}

$input = json_decode(file_get_contents('php://input'), true);
$db_id = (int)($input['db_id'] ?? 0);
if (!$db_id) { echo json_encode(['ok' => false, 'error' => 'Invalid ID']); exit; }

try {
    $db = getDB();
    if ($isAdmin) {
        $stmt = $db->prepare("SELECT * FROM orders WHERE id = ?");
        $stmt->execute([$db_id]);
    } elseif ($email) {
        $stmt = $db->prepare("SELECT * FROM orders WHERE id = ? AND phone = ?");
        $stmt->execute([$db_id, $email]);
    } else {
        $stmt = $db->prepare("SELECT * FROM orders WHERE id = ? AND guest_token = ?");
        $stmt->execute([$db_id, $guestToken]);
    }
    $order = $stmt->fetch();
} catch (Exception $e) {
    echo json_encode(['ok' => false, 'error' => 'Database error']); exit;
}

if (!$order || empty($order['file_output'])) {
    echo json_encode(['ok' => false, 'error' => 'Order tidak ditemukan']); exit;
}

if (!file_exists(__DIR__ . '/../' . $order['file_output'])) {
    echo json_encode(['ok' => false, 'error' => 'File belum tersedia']); exit;
}

// Token berlaku 5 menit — cukup untuk Office Online fetch file
$expires = time() + 300;
$sig     = hash_hmac('sha256', $db_id . ':' . $expires, PREVIEW_SECRET_KEY);

$fileUrl = APP_URL . '/api/serve_preview.php'
        . '?id=' . $db_id . '&t=' . $expires . '&sig=' . urlencode($sig);

$viewerUrl = 'https://view.officeapps.live.com/op/embed.aspx?src='
           . urlencode($fileUrl) . '&wdScale=PageWidth';

echo json_encode(['ok' => true, 'viewer_url' => $viewerUrl]);
