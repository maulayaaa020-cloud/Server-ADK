<?php
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

header('Content-Type: application/json');
$data    = json_decode(file_get_contents('php://input'), true);
$dbId    = (int)($data['db_id'] ?? 0);
$orderId = $data['order_id'] ?? '';

if (!$dbId || !$orderId) { echo json_encode(['status' => 'error']); exit; }

try {
    $db  = getDB();
    $row = $db->prepare("SELECT status FROM orders WHERE id = :id");
    $row->execute([':id' => $dbId]);
    $order = $row->fetch();

    if (!$order) { echo json_encode(['status' => 'error']); exit; }
    if ($order['status'] !== 'pending') {
        echo json_encode(['status' => $order['status']]);
        exit;
    }

    // Cek status ke DOKU
    $requestId = uniqid('adk-', true);
    $timestamp = gmdate('Y-m-d\TH:i:s\Z');
    $target    = '/orders/v1/status/' . urlencode($orderId);

    $strToSign = "Client-Id:" . DOKU_CLIENT_ID . "\n"
               . "Request-Id:" . $requestId . "\n"
               . "Request-Timestamp:" . $timestamp . "\n"
               . "Request-Target:" . $target;
    $hmacKey   = strncmp(DOKU_SECRET_KEY, 'doku_key_', 9) === 0 ? substr(DOKU_SECRET_KEY, 9) : DOKU_SECRET_KEY;
    $signature = 'HMACSHA256=' . base64_encode(hash_hmac('sha256', $strToSign, $hmacKey, true));

    $ch = curl_init(DOKU_BASE_URL . $target);
    curl_setopt_array($ch, [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_HTTPHEADER     => [
            'Content-Type: application/json',
            'Client-Id: ' . DOKU_CLIENT_ID,
            'Request-Id: ' . $requestId,
            'Request-Timestamp: ' . $timestamp,
            'Signature: ' . $signature,
        ],
    ]);
    $resp   = curl_exec($ch);
    $status = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    $result = json_decode($resp, true);
    $txStatus = $result['transaction']['status'] ?? '';

    if (in_array($txStatus, ['SUCCESS', 'SETTLEMENT'])) {
        $db->prepare("UPDATE orders SET status = 'paid' WHERE id = :id")->execute([':id' => $dbId]);
        echo json_encode(['status' => 'paid']);
    } elseif (in_array($txStatus, ['EXPIRED', 'FAILED', 'CANCELLED'])) {
        $db->prepare("UPDATE orders SET status = 'failed' WHERE id = :id")->execute([':id' => $dbId]);
        echo json_encode(['status' => 'failed']);
    } else {
        echo json_encode(['status' => 'pending', 'doku' => $txStatus]);
    }

} catch (Exception $e) {
    echo json_encode(['status' => 'error', 'error' => $e->getMessage()]);
}
