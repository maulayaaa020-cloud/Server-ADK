<?php
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

header('Content-Type: application/json');

$data       = json_decode(file_get_contents('php://input'), true);
$phone      = trim($data['phone']       ?? '');
$guestToken = trim($data['guest_token'] ?? '');

if (!$phone || !$guestToken) {
    echo json_encode(['ok' => false, 'error' => 'Data tidak valid']);
    exit;
}

// Verifikasi OTP sudah dilakukan client-side via device token cookie
$dtoken = $_COOKIE['adk_dtoken'] ?? '';
$dphone = $_COOKIE['adk_dphone'] ?? '';

if ($dphone !== $phone || !$dtoken) {
    echo json_encode(['ok' => false, 'error' => 'Verifikasi OTP diperlukan']);
    exit;
}

try {
    $db = getDB();

    // Pastikan device token valid di DB
    $chk = $db->prepare(
        "SELECT id FROM trusted_devices WHERE phone = :p AND token = :t AND expires_at > NOW() LIMIT 1"
    );
    $chk->execute([':p' => $phone, ':t' => $dtoken]);
    if (!$chk->fetch()) {
        echo json_encode(['ok' => false, 'error' => 'Sesi OTP tidak valid, silahkan verifikasi ulang']);
        exit;
    }

    // Pindahkan semua order tamu ke nomor telepon
    $upd = $db->prepare(
        "UPDATE orders SET phone = :phone WHERE guest_token = :gt AND (phone IS NULL OR phone = '')"
    );
    $upd->execute([':phone' => $phone, ':gt' => $guestToken]);

    // Update session
    $_SESSION['phone']     = $phone;
    $_SESSION['cek_phone'] = $phone;
    unset($_SESSION['guest_token']);
    unset($_SESSION['is_guest']);

    echo json_encode(['ok' => true]);

} catch (Exception $e) {
    echo json_encode(['ok' => false, 'error' => 'Database error']);
}
