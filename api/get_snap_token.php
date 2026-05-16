<?php
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

header('Content-Type: application/json');
$data    = json_decode(file_get_contents('php://input'), true);
$dbId    = (int)($data['db_id']   ?? 0);
$orderId = $data['order_id'] ?? '';
$harga   = (int)($data['harga']   ?? 0);

if (!$dbId || !$orderId || !$harga) {
    echo json_encode(['url' => null]);
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

    // Kembalikan URL yang sudah ada jika sudah dibuat sebelumnya
    if (!empty($order['snap_token'])) {
        echo json_encode(['url' => $order['snap_token']]);
        exit;
    }

    // Buat payment order di DOKU
    $requestId = uniqid('adk-', true);
    $timestamp = gmdate('Y-m-d\TH:i:s\Z');
    $target    = '/checkout/v1/payment';

    $body = json_encode([
        'order' => [
            'amount'          => $harga,
            'invoice_number'  => $orderId,
            'currency'        => 'IDR',
            'callback_url'    => APP_URL . '/history.php',
            'auto_redirect'   => false,
        ],
        'payment' => [
            'payment_due_date' => 30,
        ],
        'customer' => [
            'name'  => 'Pelanggan ADK',
            'email' => 'pelanggan@adk.com',
        ],
    ], JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE);

    $digest    = base64_encode(hash('sha256', $body, true));
    $strToSign = "Client-Id:" . DOKU_CLIENT_ID . "\n"
               . "Request-Id:" . $requestId . "\n"
               . "Request-Timestamp:" . $timestamp . "\n"
               . "Request-Target:" . $target . "\n"
               . "Digest:" . $digest;
    $signature = 'HMACSHA256=' . base64_encode(hash_hmac('sha256', $strToSign, DOKU_SECRET_KEY, true));

    // DEBUG — hapus setelah masalah teratasi
    error_log("[DOKU DEBUG] timestamp=" . $timestamp);
    error_log("[DOKU DEBUG] client_id=" . DOKU_CLIENT_ID);
    error_log("[DOKU DEBUG] secret_len=" . strlen(DOKU_SECRET_KEY));
    error_log("[DOKU DEBUG] digest=" . $digest);
    error_log("[DOKU DEBUG] strToSign=" . str_replace("\n", "|", $strToSign));
    error_log("[DOKU DEBUG] signature=" . $signature);

    $ch = curl_init(DOKU_BASE_URL . $target);
    $curlOpts = [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_POST           => true,
        CURLOPT_POSTFIELDS     => $body,
        CURLOPT_HTTPHEADER     => [
            'Content-Type: application/json',
            'Client-Id: ' . DOKU_CLIENT_ID,
            'Request-Id: ' . $requestId,
            'Request-Timestamp: ' . $timestamp,
            'Signature: ' . $signature,
        ],
        CURLOPT_SSL_VERIFYPEER => true,
        CURLOPT_TIMEOUT        => 15,
    ];
    // Pakai CAINFO custom hanya jika file ada (Windows/XAMPP); di Linux curl pakai sistem CA
    if (defined('CACERT_PATH') && file_exists(CACERT_PATH)) {
        $curlOpts[CURLOPT_CAINFO] = CACERT_PATH;
    }
    curl_setopt_array($ch, $curlOpts);
    $resp      = curl_exec($ch);
    $curlErr   = curl_error($ch);
    $curlErrNo = curl_errno($ch);
    $status    = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    if ($resp === false || $curlErrNo) {
        error_log("[DOKU] curl error #{$curlErrNo}: {$curlErr}");
        echo json_encode(['url' => null, 'error' => "curl #{$curlErrNo}: {$curlErr}"]);
        exit;
    }

    error_log("[DOKU] HTTP {$status} response: " . $resp);

    $result = json_decode($resp, true);

    $paymentUrl = $result['response']['payment']['url'] ?? null;
    if ($status !== 200 || empty($paymentUrl)) {
        echo json_encode(['url' => null, 'error' => $resp]);
        exit;
    }

    // Simpan URL ke kolom snap_token (reuse kolom)
    $db->prepare("UPDATE orders SET snap_token = :url WHERE id = :id")
       ->execute([':url' => $paymentUrl, ':id' => $dbId]);

    echo json_encode(['url' => $paymentUrl]);

} catch (Exception $e) {
    echo json_encode(['url' => null, 'error' => $e->getMessage()]);
}
