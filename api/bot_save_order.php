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

$data = json_decode(file_get_contents('php://input'), true);
if (empty($data['phone'])) {
    http_response_code(400);
    echo json_encode(['error' => 'Parameter phone dibutuhkan']);
    exit;
}

$phone      = trim($data['phone']);
$paket      = $data['paket']      ?? 'paket3';
$font       = $data['font']       ?? 'Times New Roman';
$size       = $data['size']       ?? '12 pt';
$hidden     = $data['hidden']     ?? 'Ya';
$posisi     = $data['posisi']     ?? 'Tengah Bawah';
$pos_bab    = $data['pos_bab']    ?? 'Tengah Bawah';
$pos_isi    = $data['pos_isi']    ?? 'Kanan Atas';
$dimulai    = $data['dimulai']    ?? 'i';
$semb_dafus = $data['semb_dafus'] ?? 'Tidak';
$semb_lamprn= $data['semb_lamprn']?? 'Tidak';
$num_cover  = (int)($data['num_cover'] ?? 1);

try {
    $db = getDB();
    $db->prepare("
        INSERT INTO bot_pending_orders
            (phone, paket, font, size, hidden, posisi, pos_bab, pos_isi, dimulai, semb_dafus, semb_lamprn, num_cover)
        VALUES
            (:phone, :paket, :font, :size, :hidden, :posisi, :pos_bab, :pos_isi, :dimulai, :semb_dafus, :semb_lamprn, :num_cover)
        ON DUPLICATE KEY UPDATE
            paket=VALUES(paket), font=VALUES(font), size=VALUES(size),
            hidden=VALUES(hidden), posisi=VALUES(posisi), pos_bab=VALUES(pos_bab),
            pos_isi=VALUES(pos_isi), dimulai=VALUES(dimulai),
            semb_dafus=VALUES(semb_dafus), semb_lamprn=VALUES(semb_lamprn),
            num_cover=VALUES(num_cover), updated_at=NOW()
    ")->execute([
        ':phone'      => $phone,
        ':paket'      => $paket,
        ':font'       => $font,
        ':size'       => $size,
        ':hidden'     => $hidden,
        ':posisi'     => $posisi,
        ':pos_bab'    => $pos_bab,
        ':pos_isi'    => $pos_isi,
        ':dimulai'    => $dimulai,
        ':semb_dafus' => $semb_dafus,
        ':semb_lamprn'=> $semb_lamprn,
        ':num_cover'  => $num_cover,
    ]);

    echo json_encode(['success' => true, 'phone' => $phone]);
} catch (Exception $e) {
    http_response_code(500);
    echo json_encode(['error' => 'Server error']);
}
