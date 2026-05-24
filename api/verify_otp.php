<?php
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

header('Content-Type: application/json');
$data  = json_decode(file_get_contents('php://input'), true);
$email = trim($data['email'] ?? '');
$code  = trim($data['code']  ?? '');

if (!$email || !$code) {
    echo json_encode(['ok' => false, 'error' => 'Data tidak lengkap.']);
    exit;
}

if ($email === ADMIN_EMAIL) {
    $lockFile = sys_get_temp_dir() . '/adk_admin_lock.json';

    $lock = file_exists($lockFile)
        ? (json_decode(file_get_contents($lockFile), true) ?: [])
        : [];

    $attempts     = (int)($lock['attempts']     ?? 0);
    $lockoutUntil = (int)($lock['lockout_until'] ?? 0);

    if ($lockoutUntil > time()) {
        $remaining = $lockoutUntil - time();
        $m = floor($remaining / 60);
        $s = $remaining % 60;
        $label = $m > 0 ? "{$m} menit" . ($s ? " {$s} detik" : '') : "{$s} detik";
        echo json_encode([
            'ok'        => false,
            'error'     => "Terlalu banyak percobaan. Coba lagi dalam {$label}.",
            'locked'    => true,
            'remaining' => $remaining,
        ]);
        exit;
    }

    if (password_verify($code, ADMIN_PASSWORD_HASH)) {
        file_put_contents($lockFile, json_encode(['attempts' => 0, 'lockout_until' => 0]));
        session_regenerate_id(true);
        $_SESSION['adk_admin'] = true;
        echo json_encode(['ok' => true, 'redirect' => BASE_PATH . '/admin/dashboard.php']);
    } else {
        $attempts++;
        if ($attempts >= 3) {
            $until = time() + 600;
            file_put_contents($lockFile, json_encode(['attempts' => 0, 'lockout_until' => $until]));
            echo json_encode([
                'ok'        => false,
                'error'     => 'Terlalu banyak percobaan. Coba lagi dalam 10 menit.',
                'locked'    => true,
                'remaining' => 600,
            ]);
        } else {
            file_put_contents($lockFile, json_encode(['attempts' => $attempts, 'lockout_until' => 0]));
            echo json_encode([
                'ok'    => false,
                'error' => 'Kode salah atau sudah kedaluwarsa.',
            ]);
        }
    }
    exit;
}

try {
    $db   = getDB();
    $stmt = $db->prepare(
        "SELECT id FROM otp_codes
         WHERE phone = :p AND code = :c AND expires_at > NOW() AND used = 0
         ORDER BY created_at DESC LIMIT 1"
    );
    $stmt->execute([':p' => $email, ':c' => $code]);
    $row = $stmt->fetch();

    if (!$row) {
        echo json_encode(['ok' => false, 'error' => 'Kode salah atau sudah kedaluwarsa.']);
        exit;
    }

    $db->prepare("UPDATE otp_codes SET used = 1 WHERE id = :id")->execute([':id' => $row['id']]);

    $token = bin2hex(random_bytes(32));

    $db->prepare("INSERT INTO trusted_devices (phone, token, expires_at) VALUES (:p, :t, NOW() + INTERVAL 30 DAY)")
       ->execute([':p' => $email, ':t' => $token]);

    echo json_encode(['ok' => true, 'token' => $token]);

} catch (Exception $e) {
    echo json_encode(['ok' => false, 'error' => $e->getMessage()]);
}
