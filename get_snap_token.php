<?php
require_once 'vendor/autoload.php';
require_once 'config.php';
require_once 'db.php';

header('Content-Type: application/json');
$data = json_decode(file_get_contents('php://input'), true);

$dbId    = (int)($data['db_id']   ?? 0);
$orderId = $data['order_id'] ?? '';
$harga   = (int)($data['harga']   ?? 0);

if (!$dbId || !$orderId || !$harga) {
    echo json_encode(['token' => null]);
    exit;
}

try {
    $db = getDB();

    // Cek token yang sudah tersimpan sebelumnya
    $row = $db->prepare("SELECT snap_token, status FROM orders WHERE id = :id");
    $row->execute([':id' => $dbId]);
    $order = $row->fetch();

    if (!$order || $order['status'] !== 'pending') {
        echo json_encode(['token' => null, 'error' => 'order_not_pending']);
        exit;
    }

    if (!empty($order['snap_token'])) {
        // Kembalikan token lama — user bisa buka popup yang sama lagi
        echo json_encode(['token' => $order['snap_token']]);
        exit;
    }

    // Belum ada token → generate baru dari Midtrans
    \Midtrans\Config::$serverKey    = MIDTRANS_SERVER_KEY;
    \Midtrans\Config::$isProduction = false;
    \Midtrans\Config::$isSanitized  = true;
    \Midtrans\Config::$is3ds        = true;

    $params = [
        'transaction_details' => [
            'order_id'     => $orderId,
            'gross_amount' => $harga,
        ],
        'customer_details' => ['first_name' => 'User', 'email' => 'user@gmail.com'],
    ];

    $token = \Midtrans\Snap::getSnapToken($params);

    // Simpan token ke DB agar klik berikutnya pakai token yang sama
    $db->prepare("UPDATE orders SET snap_token = :token WHERE id = :id")
       ->execute([':token' => $token, ':id' => $dbId]);

    echo json_encode(['token' => $token]);
} catch (Exception $e) {
    echo json_encode(['token' => null, 'error' => $e->getMessage()]);
}
