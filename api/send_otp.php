<?php
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

header('Content-Type: application/json');
$data  = json_decode(file_get_contents('php://input'), true);
$phone = trim($data['phone'] ?? '');

if (!$phone) { echo json_encode(['ok' => false, 'error' => 'Nomor tidak valid.']); exit; }

if (strpos($phone, '@') !== false) {
    if ($phone === ADMIN_EMAIL) {
        echo json_encode(['ok' => true, 'admin_mode' => true]);
    } else {
        echo json_encode(['ok' => false, 'error' => 'Nomor tidak valid.']);
    }
    exit;
}

try {
    $db = getDB();

    $chk = $db->prepare("SELECT COUNT(*) FROM otp_codes WHERE phone = :p AND created_at > DATE_SUB(NOW(), INTERVAL 10 MINUTE)");
    $chk->execute([':p' => $phone]);
    if ((int)$chk->fetchColumn() >= 3) {
        echo json_encode(['ok' => false, 'error' => 'Terlalu banyak permintaan. Tunggu beberapa menit.']);
        exit;
    }

    $db->prepare("DELETE FROM otp_codes WHERE expires_at < NOW()")->execute();

    $code = str_pad(random_int(0, 999999), 6, '0', STR_PAD_LEFT);

    $db->prepare("INSERT INTO otp_codes (phone, code, expires_at) VALUES (:p, :c, NOW() + INTERVAL 5 MINUTE)")
       ->execute([':p' => $phone, ':c' => $code]);

    $ch = curl_init('https://api.fonnte.com/send');
    curl_setopt_array($ch, [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_POST           => true,
        CURLOPT_HTTPHEADER     => ['Authorization: ' . FONNTE_TOKEN],
        CURLOPT_POSTFIELDS     => http_build_query([
            'target'  => $phone,
            'message' => "Kode OTP ADK Photocopy kamu: *{$code}*\n\nBerlaku 5 menit. Jangan berikan ke siapapun.",
        ]),
    ]);
    curl_exec($ch);
    curl_close($ch);

    echo json_encode(['ok' => true]);

} catch (Exception $e) {
    echo json_encode(['ok' => false, 'error' => $e->getMessage()]);
}
