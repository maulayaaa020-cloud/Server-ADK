<?php
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';
require_once __DIR__ . '/../includes/log.php';

$rawBody           = file_get_contents('php://input');
$callbackSignature = $_SERVER['HTTP_X_CALLBACK_SIGNATURE'] ?? '';

// Verifikasi signature TriPay
$expectedSignature = hash_hmac('sha256', $rawBody, TRIPAY_PRIVATE_KEY);
if (!hash_equals($expectedSignature, $callbackSignature)) {
    http_response_code(401);
    exit(json_encode(['success' => false, 'message' => 'Invalid signature']));
}

$data = json_decode($rawBody, true);

$merchantRef = $data['merchant_ref'] ?? '';
$status      = $data['status']       ?? '';

if (!$merchantRef || !$status) {
    http_response_code(400);
    exit(json_encode(['success' => false, 'message' => 'Invalid payload']));
}

adk_log('payment', 'TriPay callback', ['ref' => $merchantRef, 'status' => $status]);

try {
    $db = getDB();

    if ($status === 'PAID') {
        $affected = $db->prepare("UPDATE orders SET status = 'paid' WHERE order_id = :ref AND status = 'pending'");
        $affected->execute([':ref' => $merchantRef]);
        adk_log('payment', 'Order PAID', ['ref' => $merchantRef, 'rows' => $affected->rowCount()]);
    } elseif (in_array($status, ['EXPIRED', 'FAILED', 'REFUND'])) {
        $db->prepare("UPDATE orders SET status = 'failed' WHERE order_id = :ref AND status = 'pending'")
           ->execute([':ref' => $merchantRef]);
        adk_log('payment', 'Order ' . $status, ['ref' => $merchantRef]);
    }
} catch (Exception $e) {
    adk_log('error', 'TriPay callback DB error', ['ref' => $merchantRef, 'err' => $e->getMessage()]);
    http_response_code(500);
    exit;
}

http_response_code(200);
echo json_encode(['success' => true]);
