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
    $row = $db->prepare("SELECT status, snap_token FROM orders WHERE id = :id");
    $row->execute([':id' => $dbId]);
    $order = $row->fetch();

    if (!$order) { echo json_encode(['status' => 'error']); exit; }
    if ($order['status'] !== 'pending') {
        echo json_encode(['status' => $order['status']]);
        exit;
    }

    $reference = $order['snap_token'] ?? '';
    if (!$reference) {
        echo json_encode(['status' => 'pending']);
        exit;
    }

    // Cek status ke TriPay
    $ch = curl_init(TRIPAY_BASE_URL . '/transaction/detail?reference=' . urlencode($reference));
    curl_setopt_array($ch, [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_HTTPHEADER     => ['Authorization: Bearer ' . TRIPAY_API_KEY],
        CURLOPT_TIMEOUT        => 10,
    ]);
    $resp   = curl_exec($ch);
    $status = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    $result   = json_decode($resp, true);
    $txStatus = $result['data']['status'] ?? '';

    if ($txStatus === 'PAID') {
        $db->prepare("UPDATE orders SET status = 'paid' WHERE id = :id")->execute([':id' => $dbId]);
        echo json_encode(['status' => 'paid']);
    } elseif (in_array($txStatus, ['EXPIRED', 'FAILED', 'REFUND'])) {
        $db->prepare("UPDATE orders SET status = 'failed' WHERE id = :id")->execute([':id' => $dbId]);
        echo json_encode(['status' => 'failed']);
    } else {
        echo json_encode(['status' => 'pending', 'tripay' => $txStatus]);
    }

} catch (Exception $e) {
    echo json_encode(['status' => 'error', 'error' => $e->getMessage()]);
}
