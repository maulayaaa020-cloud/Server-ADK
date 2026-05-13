<?php
require_once 'vendor/autoload.php';
require_once 'config.php';
require_once 'db.php';

header('Content-Type: application/json');
$data    = json_decode(file_get_contents('php://input'), true);
$dbId    = (int)($data['db_id']   ?? 0);
$orderId = $data['order_id'] ?? '';

if (!$dbId || !$orderId) { echo json_encode(['status' => 'error']); exit; }

try {
    $db = getDB();
    $row = $db->prepare("SELECT status FROM orders WHERE id = :id");
    $row->execute([':id' => $dbId]);
    $order = $row->fetch();

    if (!$order) { echo json_encode(['status' => 'error']); exit; }

    // Kalau sudah bukan pending di DB, kembalikan langsung
    if ($order['status'] !== 'pending') {
        echo json_encode(['status' => $order['status']]);
        exit;
    }

    // Cek ke Midtrans
    \Midtrans\Config::$serverKey    = MIDTRANS_SERVER_KEY;
    \Midtrans\Config::$isProduction = false;

    $mt = \Midtrans\Transaction::status($orderId);
    $ts = $mt->transaction_status ?? '';

    if ($ts === 'settlement' || $ts === 'capture') {
        $db->prepare("UPDATE orders SET status = 'paid' WHERE id = :id")
           ->execute([':id' => $dbId]);
        echo json_encode(['status' => 'paid']);
    } elseif (in_array($ts, ['expire', 'cancel', 'deny'])) {
        $db->prepare("UPDATE orders SET status = 'failed' WHERE id = :id")
           ->execute([':id' => $dbId]);
        echo json_encode(['status' => 'failed']);
    } else {
        echo json_encode(['status' => 'pending', 'midtrans' => $ts]);
    }

} catch (Exception $e) {
    echo json_encode(['status' => 'error', 'error' => $e->getMessage()]);
}
