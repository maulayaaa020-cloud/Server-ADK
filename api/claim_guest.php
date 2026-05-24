<?php
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

header('Content-Type: application/json');

$data       = json_decode(file_get_contents('php://input'), true);
$email      = trim($data['email']       ?? '');
$guestToken = trim($data['guest_token'] ?? '');

if (!$email || !$guestToken) {
    echo json_encode(['ok' => false, 'error' => 'Data tidak valid']);
    exit;
}

// Verifikasi OTP sudah dilakukan client-side via device token cookie
$dtoken = $_COOKIE['adk_dtoken'] ?? '';
$demail = $_COOKIE['adk_demail'] ?? '';

if ($demail !== $email || !$dtoken) {
    echo json_encode(['ok' => false, 'error' => 'Verifikasi OTP diperlukan']);
    exit;
}

try {
    $db = getDB();

    // Pastikan device token valid di DB
    $chk = $db->prepare(
        "SELECT id FROM trusted_devices WHERE phone = :p AND token = :t AND expires_at > NOW() LIMIT 1"
    );
    $chk->execute([':p' => $email, ':t' => $dtoken]);
    if (!$chk->fetch()) {
        echo json_encode(['ok' => false, 'error' => 'Sesi OTP tidak valid, silahkan verifikasi ulang']);
        exit;
    }

    // Pindahkan semua order tamu ke email
    $upd = $db->prepare(
        "UPDATE orders SET phone = :phone WHERE guest_token = :gt AND (phone IS NULL OR phone = '')"
    );
    $upd->execute([':phone' => $email, ':gt' => $guestToken]);

    // Update session
    $_SESSION['email']     = $email;
    $_SESSION['cek_email'] = $email;
    unset($_SESSION['guest_token']);
    unset($_SESSION['is_guest']);

    echo json_encode(['ok' => true]);

} catch (Exception $e) {
    echo json_encode(['ok' => false, 'error' => 'Database error']);
}
