<?php
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

header('Content-Type: application/json');
header('Access-Control-Allow-Origin: *');

$apiKey = $_SERVER['HTTP_X_BOT_KEY'] ?? '';
if ($apiKey !== 'adkivia-bot-2026') {
    http_response_code(401);
    echo json_encode(['error' => 'Unauthorized']);
    exit;
}

$data  = json_decode(file_get_contents('php://input'), true);
$phone = trim($data['phone'] ?? '');

if (empty($phone)) {
    http_response_code(400);
    echo json_encode(['error' => 'Parameter phone dibutuhkan']);
    exit;
}

try {
    $db   = getDB();
    $stmt = $db->prepare("SELECT order_done, email, paket, paket_confirmed, specs_confirmed, font, size, hidden, posisi, pos_bab, pos_isi, dimulai, semb_dafus, semb_lamprn, num_cover FROM bot_pending_orders WHERE phone = ?");
    $stmt->execute([$phone]);
    $row  = $stmt->fetch(PDO::FETCH_ASSOC);

    echo json_encode($row
        ? [
            'order_done'      => (int)$row['order_done'],
            'email'           => $row['email'],
            'paket'           => $row['paket'] ?? 'paket3',
            'paket_confirmed' => (int)$row['paket_confirmed'],
            'specs_confirmed' => (int)$row['specs_confirmed'],
            'font'            => $row['font'] ?? '',
            'size'            => $row['size'] ?? '',
            'hidden'          => $row['hidden'] ?? '',
            'posisi'          => $row['posisi'] ?? '',
            'pos_bab'         => $row['pos_bab'] ?? '',
            'pos_isi'         => $row['pos_isi'] ?? '',
            'dimulai'         => $row['dimulai'] ?? '',
            'semb_dafus'      => $row['semb_dafus'] ?? '',
            'semb_lamprn'     => $row['semb_lamprn'] ?? '',
            'num_cover'       => $row['num_cover'] ?? '',
            'phone'           => $phone,
            'exists'          => true,
          ]
        : ['order_done' => 0, 'paket_confirmed' => 0, 'specs_confirmed' => 0, 'exists' => false, 'phone' => $phone]
    );
} catch (Exception $e) {
    http_response_code(500);
    echo json_encode(['error' => $e->getMessage()]);
}
