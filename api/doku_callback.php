<?php
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

// DOKU JOKUL notification endpoint
// Daftarkan URL ini di DOKU merchant dashboard: APP_URL/api/doku_callback.php

$rawBody = file_get_contents('php://input');

// Ambil headers yang dikirim DOKU
$clientId  = $_SERVER['HTTP_CLIENT_ID']          ?? '';
$requestId = $_SERVER['HTTP_REQUEST_ID']          ?? '';
$timestamp = $_SERVER['HTTP_REQUEST_TIMESTAMP']   ?? '';
$incoming  = $_SERVER['HTTP_SIGNATURE']           ?? '';

// Verifikasi Client-Id
if ($clientId !== DOKU_CLIENT_ID) {
    http_response_code(401);
    exit;
}

// Verifikasi signature DOKU
$target  = '/api/doku_callback.php';
$digest  = base64_encode(hash('sha256', $rawBody, true));
$strSign = "Client-Id:{$clientId}\n"
         . "Request-Id:{$requestId}\n"
         . "Request-Timestamp:{$timestamp}\n"
         . "Request-Target:{$target}\n"
         . "Digest:{$digest}";
$expected = 'HMACSHA256=' . base64_encode(hash_hmac('sha256', $strSign, DOKU_SECRET_KEY, true));

if (!hash_equals($expected, $incoming)) {
    error_log("[DOKU callback] Signature mismatch. Expected: {$expected} | Got: {$incoming}");
    http_response_code(401);
    exit;
}

$data     = json_decode($rawBody, true);
$invoiceNo = $data['order']['invoice_number'] ?? '';
$txStatus  = $data['transaction']['status']   ?? '';

if (!$invoiceNo || !$txStatus) {
    http_response_code(400);
    exit;
}

error_log("[DOKU callback] invoice={$invoiceNo} status={$txStatus}");

try {
    $db = getDB();

    if (in_array($txStatus, ['SUCCESS', 'SETTLEMENT'])) {
        $db->prepare("UPDATE orders SET status = 'paid' WHERE order_id = :inv AND status = 'pending'")
           ->execute([':inv' => $invoiceNo]);
    } elseif (in_array($txStatus, ['EXPIRED', 'FAILED', 'CANCELLED'])) {
        $db->prepare("UPDATE orders SET status = 'failed' WHERE order_id = :inv AND status = 'pending'")
           ->execute([':inv' => $invoiceNo]);
    }
} catch (Exception $e) {
    error_log("[DOKU callback] DB error: " . $e->getMessage());
    http_response_code(500);
    exit;
}

http_response_code(200);
echo json_encode(['status' => 'ok']);
