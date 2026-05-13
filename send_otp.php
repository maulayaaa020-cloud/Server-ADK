<?php
require_once 'config.php';
require_once 'db.php';

header('Content-Type: application/json');
$data  = json_decode(file_get_contents('php://input'), true);
$phone = trim($data['phone'] ?? '');

if (!$phone) { echo json_encode(['ok' => false, 'error' => 'Nomor tidak valid.']); exit; }

try {
    $db = getDB();

    // Rate limit: maks 3 OTP per 10 menit
    $chk = $db->prepare("SELECT COUNT(*) FROM otp_codes WHERE phone = :p AND created_at > DATE_SUB(NOW(), INTERVAL 10 MINUTE)");
    $chk->execute([':p' => $phone]);
    if ((int)$chk->fetchColumn() >= 3) {
        echo json_encode(['ok' => false, 'error' => 'Terlalu banyak permintaan. Tunggu beberapa menit.']);
        exit;
    }

    // Hapus kode lama yang sudah expired
    $db->prepare("DELETE FROM otp_codes WHERE expires_at < NOW()")->execute();

    // Generate kode 6 digit
    $code = str_pad(random_int(0, 999999), 6, '0', STR_PAD_LEFT);

    // Gunakan NOW() MySQL langsung — bebas timezone PHP
    $db->prepare("INSERT INTO otp_codes (phone, code, expires_at) VALUES (:p, :c, NOW() + INTERVAL 5 MINUTE)")
       ->execute([':p' => $phone, ':c' => $code]);

    // TODO: Kirim via WhatsApp/SMS
    // DEV MODE: kembalikan kode langsung ke client
    echo json_encode(['ok' => true, 'otp' => $code, 'dev' => true]);

} catch (Exception $e) {
    echo json_encode(['ok' => false, 'error' => $e->getMessage()]);
}
