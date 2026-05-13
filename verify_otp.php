<?php
require_once 'config.php';
require_once 'db.php';

header('Content-Type: application/json');
$data  = json_decode(file_get_contents('php://input'), true);
$phone = trim($data['phone'] ?? '');
$code  = trim($data['code']  ?? '');

if (!$phone || !$code) {
    echo json_encode(['ok' => false, 'error' => 'Data tidak lengkap.']);
    exit;
}

try {
    $db   = getDB();
    $stmt = $db->prepare(
        "SELECT id FROM otp_codes
         WHERE phone = :p AND code = :c AND expires_at > NOW() AND used = 0
         ORDER BY created_at DESC LIMIT 1"
    );
    $stmt->execute([':p' => $phone, ':c' => $code]);
    $row = $stmt->fetch();

    if (!$row) {
        echo json_encode(['ok' => false, 'error' => 'Kode salah atau sudah kedaluwarsa.']);
        exit;
    }

    // Tandai kode sebagai sudah dipakai
    $db->prepare("UPDATE otp_codes SET used = 1 WHERE id = :id")->execute([':id' => $row['id']]);

    // Generate device token (64 char hex, berlaku 30 hari)
    $token = bin2hex(random_bytes(32));

    // Gunakan NOW() MySQL langsung — bebas timezone PHP
    $db->prepare("INSERT INTO trusted_devices (phone, token, expires_at) VALUES (:p, :t, NOW() + INTERVAL 30 DAY)")
       ->execute([':p' => $phone, ':t' => $token]);

    echo json_encode(['ok' => true, 'token' => $token]);

} catch (Exception $e) {
    echo json_encode(['ok' => false, 'error' => $e->getMessage()]);
}
