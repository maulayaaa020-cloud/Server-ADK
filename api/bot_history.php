<?php
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

header('Content-Type: application/json');
header('Access-Control-Allow-Origin: *');

// Simple API key auth
$apiKey = $_SERVER['HTTP_X_BOT_KEY'] ?? '';
if ($apiKey !== 'adkivia-bot-2026') {
    http_response_code(401);
    echo json_encode(['error' => 'Unauthorized']);
    exit;
}

$waNumber = trim($_GET['wa'] ?? '');
if (empty($waNumber)) {
    http_response_code(400);
    echo json_encode(['error' => 'Parameter wa dibutuhkan']);
    exit;
}

// Bersihkan format WhatsApp (@c.us, @lid, dll)
$phone = preg_replace('/@.*$/', '', $waNumber);
$phone = preg_replace('/[^0-9]/', '', $phone);

try {
    $db = getDB();

    // Ambil 5 order penomoran terakhir
    $stmt = $db->prepare("
        SELECT order_id, paket, harga, status, created_at
        FROM orders
        WHERE phone = ?
        ORDER BY created_at DESC
        LIMIT 5
    ");
    $stmt->execute([$phone]);
    $penomoran = $stmt->fetchAll();

    $daftarIsi = [];

    $total = count($penomoran) + count($daftarIsi);

    echo json_encode([
        'success'   => true,
        'wa_number' => $waNumber,
        'phone'     => $phone,
        'total'     => $total,
        'penomoran' => $penomoran,
        'daftar_isi'=> $daftarIsi,
    ]);

} catch (Exception $e) {
    http_response_code(500);
    echo json_encode(['error' => $e->getMessage()]);
}
