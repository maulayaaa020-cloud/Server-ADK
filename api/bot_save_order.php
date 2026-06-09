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

// GET: cek status order_done
if ($_SERVER['REQUEST_METHOD'] === 'GET') {
    $phone = trim($_GET['phone'] ?? '');
    if (empty($phone)) {
        http_response_code(400);
        echo json_encode(['error' => 'Parameter phone dibutuhkan']);
        exit;
    }
    try {
        $db   = getDB();
        $stmt = $db->prepare("SELECT phone, email, paket, order_done, paket_confirmed, specs_confirmed, last_activity FROM bot_pending_orders WHERE phone = ?");
        $stmt->execute([$phone]);
        $row  = $stmt->fetch(PDO::FETCH_ASSOC);
        echo json_encode($row ? [
            'order_done'      => (int)$row['order_done'],
            'email'           => $row['email'],
            'paket'           => $row['paket'],
            'paket_confirmed' => (int)$row['paket_confirmed'],
            'specs_confirmed' => (int)$row['specs_confirmed'],
            'last_activity'   => $row['last_activity'],
            'exists'          => true,
        ] : [
            'order_done'      => 0,
            'paket_confirmed' => 0,
            'specs_confirmed' => 0,
            'last_activity'   => null,
            'exists'          => false,
        ]);
    } catch (Exception $e) {
        http_response_code(500);
        echo json_encode(['error' => 'Server error']);
    }
    exit;
}

$data = json_decode(file_get_contents('php://input'), true);
if (empty($data['phone'])) {
    http_response_code(400);
    echo json_encode(['error' => 'Parameter phone dibutuhkan']);
    exit;
}

// Helper: pakai nilai dari request jika tidak kosong, else null (pertahankan nilai DB)
function val($data, $key, $default = '') {
    $v = isset($data[$key]) ? trim((string)$data[$key]) : '';
    return $v !== '' ? $v : null;
}

$phone       = trim($data['phone']);
$email       = val($data, 'email');
$paket_raw   = val($data, 'paket');
$paket       = $paket_raw ?? 'paket3';
$font        = val($data, 'font');
$size        = val($data, 'size');
$hidden      = val($data, 'hidden');
$posisi      = val($data, 'posisi');
$pos_bab     = val($data, 'pos_bab');
$pos_isi     = val($data, 'pos_isi');
$dimulai     = val($data, 'dimulai');
$semb_dafus  = val($data, 'semb_dafus');
$semb_lamprn = val($data, 'semb_lamprn');
$num_cover   = isset($data['num_cover']) && $data['num_cover'] !== '' ? (int)$data['num_cover'] : null;

// paket_confirmed: 1 hanya kalau user eksplisit kirim paket
$ins_paket_confirmed = ($paket_raw !== null) ? 1 : 0;
$upd_paket_confirmed = ($paket_raw !== null) ? 1 : null;

// specs_confirmed: hanya set kalau eksplisit dikirim (0 = reset, 1 = confirmed), null = pertahankan DB
$specs_confirmed_raw = isset($data['specs_confirmed']) && $data['specs_confirmed'] !== '' ? (int)$data['specs_confirmed'] : null;
$upd_specs_confirmed = $specs_confirmed_raw;
$ins_specs_confirmed = 0;

// order_done: hanya set kalau eksplisit dikirim (0 = reset untuk order baru)
$upd_order_done = isset($data['order_done']) && $data['order_done'] !== '' ? (int)$data['order_done'] : null;

// Untuk INSERT baru: pakai default jika null
$ins_font        = $font        ?? 'Times New Roman';
$ins_size        = $size        ?? '12 pt';
$ins_hidden      = $hidden      ?? 'Ya';
$ins_posisi      = $posisi      ?? 'Tengah Bawah';
$ins_pos_bab     = $pos_bab     ?? 'Tengah Bawah';
$ins_pos_isi     = $pos_isi     ?? 'Kanan Atas';
$ins_dimulai     = $dimulai     ?? 'i';
$ins_semb_dafus  = $semb_dafus  ?? 'Tidak';
$ins_semb_lamprn = $semb_lamprn ?? 'Tidak';
$ins_num_cover   = $num_cover   ?? 1;

try {
    $db = getDB();
    $db->prepare("
        INSERT INTO bot_pending_orders
            (phone, email, paket, paket_confirmed, specs_confirmed, font, size, hidden, posisi, pos_bab, pos_isi, dimulai, semb_dafus, semb_lamprn, num_cover, last_activity)
        VALUES
            (:phone, :email, :paket, :paket_confirmed, :specs_confirmed, :font, :size, :hidden, :posisi, :pos_bab, :pos_isi, :dimulai, :semb_dafus, :semb_lamprn, :num_cover, NOW())
        ON DUPLICATE KEY UPDATE
            email           = IF(:u_email           IS NULL, email,           :u_email),
            paket           = IF(:u_paket IS NULL, paket, :u_paket),
            paket_confirmed = IF(:u_paket_confirmed IS NULL, paket_confirmed, :u_paket_confirmed),
            specs_confirmed = IF(:u_specs_confirmed IS NULL, specs_confirmed, :u_specs_confirmed),
            font            = IF(:u_font            IS NULL, font,            :u_font),
            size         = IF(:u_size         IS NULL, size,         :u_size),
            hidden       = IF(:u_hidden       IS NULL, hidden,       :u_hidden),
            posisi       = IF(:u_posisi       IS NULL, posisi,       :u_posisi),
            pos_bab      = IF(:u_pos_bab      IS NULL, pos_bab,      :u_pos_bab),
            pos_isi      = IF(:u_pos_isi      IS NULL, pos_isi,      :u_pos_isi),
            dimulai      = IF(:u_dimulai      IS NULL, dimulai,      :u_dimulai),
            semb_dafus   = IF(:u_semb_dafus   IS NULL, semb_dafus,   :u_semb_dafus),
            semb_lamprn  = IF(:u_semb_lamprn  IS NULL, semb_lamprn,  :u_semb_lamprn),
            num_cover    = IF(:u_num_cover    IS NULL, num_cover,    :u_num_cover),
            order_done    = IF(:u_order_done IS NULL, order_done, :u_order_done),
            last_activity = NOW(),
            updated_at    = NOW()
    ")->execute([
        ':phone'              => $phone,
        ':email'              => $email ?? '',
        ':paket'              => $paket,
        ':paket_confirmed'    => $ins_paket_confirmed,
        ':specs_confirmed'    => $ins_specs_confirmed,
        ':font'               => $ins_font,
        ':size'          => $ins_size,
        ':hidden'        => $ins_hidden,
        ':posisi'        => $ins_posisi,
        ':pos_bab'       => $ins_pos_bab,
        ':pos_isi'       => $ins_pos_isi,
        ':dimulai'       => $ins_dimulai,
        ':semb_dafus'    => $ins_semb_dafus,
        ':semb_lamprn'   => $ins_semb_lamprn,
        ':num_cover'     => $ins_num_cover,
        ':u_email'            => $email,
        ':u_paket'            => $paket_raw,
        ':u_paket_confirmed'  => $upd_paket_confirmed,
        ':u_specs_confirmed'  => $upd_specs_confirmed,
        ':u_order_done'       => $upd_order_done,
        ':u_font'             => $font,
        ':u_size'        => $size,
        ':u_hidden'      => $hidden,
        ':u_posisi'      => $posisi,
        ':u_pos_bab'     => $pos_bab,
        ':u_pos_isi'     => $pos_isi,
        ':u_dimulai'     => $dimulai,
        ':u_semb_dafus'  => $semb_dafus,
        ':u_semb_lamprn' => $semb_lamprn,
        ':u_num_cover'   => $num_cover,
    ]);

    echo json_encode(['success' => true, 'phone' => $phone]);
} catch (Exception $e) {
    http_response_code(500);
    echo json_encode(['error' => 'Server error']);
}
