<?php
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';
header('Content-Type: application/json');

$isAdmin    = !empty($_SESSION['adk_admin']);
$phone      = $_SESSION['phone'] ?? $_SESSION['cek_phone'] ?? null;
$guestToken = $_SESSION['guest_token'] ?? $_COOKIE['adk_guest'] ?? null;

if (!$isAdmin && !$phone && !$guestToken) {
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
    } elseif ($phone) {
        $stmt = $db->prepare("SELECT * FROM orders WHERE id = ? AND phone = ?");
        $stmt->execute([$db_id, $phone]);
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
$sig     = hash_hmac('sha256', $db_id . ':' . $expires, DOKU_SECRET_KEY);

$fileUrl = 'https://adkphotocopy.com/api/serve_preview.php'
         . '?id=' . $db_id . '&t=' . $expires . '&sig=' . urlencode($sig);

$vw = (int)($input['vw'] ?? 1200);
$wdScale = ($vw < 768) ? '50' : 'PageWidth';
$viewerUrl = 'https://view.officeapps.live.com/op/embed.aspx?src='
           . urlencode($fileUrl) . '&wdScale=' . $wdScale;

echo json_encode(['ok' => true, 'viewer_url' => $viewerUrl]);
