<?php
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

header('Content-Type: application/json');
$data    = json_decode(file_get_contents('php://input'), true);
$dbId    = (int)($data['db_id']   ?? 0);
$orderId = $data['order_id']      ?? '';
$harga   = (int)($data['harga']   ?? 0);
$method  = strtoupper(trim($data['method'] ?? 'QRIS'));

if (!$dbId || !$orderId || !$harga) {
    echo json_encode(['url' => null, 'error' => 'Missing parameters']);
    exit;
}

try {
    $db  = getDB();
    $row = $db->prepare("SELECT snap_token, status FROM orders WHERE id = :id");
    $row->execute([':id' => $dbId]);
    $order = $row->fetch();

    if (!$order || $order['status'] !== 'pending') {
        echo json_encode(['url' => null, 'error' => 'order_not_pending']);
        exit;
    }

    // Kembalikan URL yang sudah ada
    if (!empty($order['snap_token'])) {
        $existing = json_decode($order['snap_token'], true);
        if ($existing && isset($existing['url'])) {
            echo json_encode(['url' => $existing['url']]);
            exit;
        }
        echo json_encode(['url' => $order['snap_token']]);
        exit;
    }

    $signature   = hash_hmac('sha256', TRIPAY_MERCHANT_CODE . $orderId . $harga, TRIPAY_PRIVATE_KEY);
    $expiredTime = time() + (24 * 60 * 60);

    $body = json_encode([
        'method'         => $method,
        'merchant_ref'   => $orderId,
        'amount'         => $harga,
        'customer_name'  => 'Pelanggan ADK',
        'customer_email' => 'pelanggan@adkphotocopy.com',
        'customer_phone' => '08000000000',
        'order_items'    => [[
            'sku'      => 'ADK-PENOMORAN',
            'name'     => 'Layanan Penomoran Halaman ADK',
            'price'    => $harga,
            'quantity' => 1,
        ]],
        'callback_url'   => 'https://adkphotocopy.com/api/tripay_callback.php',
        'return_url'     => 'https://adkphotocopy.com/history.php',
        'expired_time'   => $expiredTime,
        'signature'      => $signature,
    ]);

    $ch = curl_init(TRIPAY_BASE_URL . '/transaction/create');
    curl_setopt_array($ch, [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_POST           => true,
        CURLOPT_POSTFIELDS     => $body,
        CURLOPT_HTTPHEADER     => [
            'Content-Type: application/json',
            'Authorization: Bearer ' . TRIPAY_API_KEY,
        ],
        CURLOPT_TIMEOUT        => 15,
    ]);
    $resp   = curl_exec($ch);
    $status = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    error_log("[TriPay create] HTTP {$status} response: " . $resp);

    $result = json_decode($resp, true);

    if (empty($result['success'])) {
        echo json_encode(['url' => null, 'error' => $result['message'] ?? $resp]);
        exit;
    }

    $payUrl    = $result['data']['pay_url'] ?? $result['data']['checkout_url'] ?? null;
    $reference = $result['data']['reference'] ?? null;

    if (!$payUrl) {
        echo json_encode(['url' => null, 'error' => 'No pay_url returned', 'debug' => $result['data'] ?? $resp]);
        exit;
    }

    $db->prepare("UPDATE orders SET snap_token = :data WHERE id = :id")
       ->execute([':data' => json_encode(['reference' => $reference, 'url' => $payUrl]), ':id' => $dbId]);

    echo json_encode(['url' => $payUrl, 'reference' => $reference]);

} catch (Throwable $e) {
    echo json_encode(['url' => null, 'error' => $e->getMessage()]);
}
