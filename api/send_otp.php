<?php
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

header('Content-Type: application/json');
$data  = json_decode(file_get_contents('php://input'), true);
$email = trim($data['email'] ?? '');

if (!$email || !filter_var($email, FILTER_VALIDATE_EMAIL)) {
    echo json_encode(['ok' => false, 'error' => 'Email tidak valid.']);
    exit;
}

if ($email === ADMIN_EMAIL) {
    echo json_encode(['ok' => true, 'admin_mode' => true]);
    exit;
}

try {
    $db = getDB();

    $chk = $db->prepare("SELECT COUNT(*) FROM otp_codes WHERE phone = :p AND created_at > DATE_SUB(NOW(), INTERVAL 10 MINUTE)");
    $chk->execute([':p' => $email]);
    if ((int)$chk->fetchColumn() >= 3) {
        echo json_encode(['ok' => false, 'error' => 'Terlalu banyak permintaan. Tunggu beberapa menit.']);
        exit;
    }

    $db->prepare("DELETE FROM otp_codes WHERE expires_at < NOW()")->execute();

    $code = str_pad(random_int(0, 999999), 6, '0', STR_PAD_LEFT);

    $db->prepare("INSERT INTO otp_codes (phone, code, expires_at) VALUES (:p, :c, NOW() + INTERVAL 5 MINUTE)")
       ->execute([':p' => $email, ':c' => $code]);

    $ch = curl_init('https://api.resend.com/emails');
    curl_setopt_array($ch, [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_POST           => true,
        CURLOPT_HTTPHEADER     => [
            'Authorization: Bearer ' . RESEND_API_KEY,
            'Content-Type: application/json',
        ],
        CURLOPT_POSTFIELDS => json_encode([
            'from'    => RESEND_FROM,
            'to'      => [$email],
            'subject' => 'Kode OTP ADK Photocopy',
            'text'    => "Kode OTP ADK Photocopy kamu: {$code}\n\nBerlaku 5 menit. Jangan berikan ke siapapun.",
        ]),
    ]);
    $res      = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    if ($httpCode < 200 || $httpCode >= 300) {
        error_log('Resend error: HTTP ' . $httpCode . ' — ' . $res);
        echo json_encode(['ok' => false, 'error' => 'Gagal mengirim OTP. Coba lagi.']);
        exit;
    }

    echo json_encode(['ok' => true]);

} catch (Exception $e) {
    error_log('send_otp error: ' . $e->getMessage());
    echo json_encode(['ok' => false, 'error' => 'Gagal mengirim OTP. Coba lagi.']);
}
