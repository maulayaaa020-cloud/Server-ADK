<?php
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';
require_once __DIR__ . '/../includes/log.php';

header('Content-Type: application/json');
header('Access-Control-Allow-Origin: *');

$apiKey = $_SERVER['HTTP_X_BOT_KEY'] ?? '';
if ($apiKey !== 'adkivia-bot-2026') {
    http_response_code(401);
    echo json_encode(['error' => 'Unauthorized']);
    exit;
}

set_time_limit(300);

$data      = json_decode(file_get_contents('php://input'), true);
$phone     = trim($data['phone']     ?? '');
$filename  = trim($data['filename']  ?? '');
$origName  = trim($data['orig_name'] ?? $filename);

if (strpos($origName, '_ADK_') !== false) {
    http_response_code(400);
    echo json_encode(['error' => 'FILE_IS_OUTPUT', 'message' => 'File ini adalah hasil pemrosesan sebelumnya']);
    exit;
}

if (!$phone || !$filename) {
    http_response_code(400);
    echo json_encode(['error' => 'Parameter phone dan filename dibutuhkan']);
    exit;
}

$inputPath = __DIR__ . '/../uploads/bot/' . basename($filename);
if (!file_exists($inputPath)) {
    http_response_code(404);
    echo json_encode(['error' => 'File tidak ditemukan']);
    exit;
}

// Ambil parameter order dari DB
try {
    $db   = getDB();
    $stmt = $db->prepare("SELECT * FROM bot_pending_orders WHERE phone = ?");
    $stmt->execute([$phone]);
    $order = $stmt->fetch();
} catch (Exception $e) {
    http_response_code(500);
    echo json_encode(['error' => 'DB error']);
    exit;
}

if ($order) {
    if (empty($order['email'])) {
        http_response_code(422);
        echo json_encode(['error' => 'EMAIL_REQUIRED', 'message' => 'Email user belum diisi. Minta user konfirmasi email terlebih dahulu.']);
        exit;
    }
    if (empty($order['paket_confirmed'])) {
        http_response_code(422);
        echo json_encode(['error' => 'PAKET_REQUIRED', 'message' => 'Paket belum dikonfirmasi user. Minta user pilih paket terlebih dahulu.']);
        exit;
    }
}

if (!$order) {
    // Gunakan default jika tidak ada parameter tersimpan
    $order = [
        'paket'      => 'paket3',
        'font'       => 'Times New Roman',
        'size'       => '12 pt',
        'hidden'     => 'Ya',
        'posisi'     => 'Tengah Bawah',
        'pos_bab'    => 'Tengah Bawah',
        'pos_isi'    => 'Kanan Atas',
        'dimulai'    => 'i',
        'semb_dafus' => 'Tidak',
        'semb_lamprn'=> 'Tidak',
        'num_cover'  => 1,
    ];
}

// Siapkan path output — format sama seperti website
$hasilDir = __DIR__ . '/../uploads/bot/hasil/';
if (!is_dir($hasilDir)) mkdir($hasilDir, 0755, true);

$namaAsli  = pathinfo($origName, PATHINFO_FILENAME);
$ext       = strtolower(pathinfo($origName, PATHINFO_EXTENSION)) ?: 'docx';
$namaAsli  = preg_replace('/^\d+_/', '', $namaAsli);
$namaAsli  = preg_replace('/[^a-zA-Z0-9]/', '_', $namaAsli);
$namaAsli  = trim(preg_replace('/_+/', '_', $namaAsli), '_');
$timestamp = date('YmdHis');
$random    = bin2hex(random_bytes(4));
$outputName = $namaAsli . '_ADK_' . $timestamp . '_' . $random . '.' . $ext;
$outputPath = $hasilDir . $outputName;

// Build command
$python = PYTHON_EXE;
$script = PYTHON_SCRIPT_DIR . '/main.py';

$args = implode(' ', [
    escapeshellarg($inputPath),
    escapeshellarg($outputPath),
    escapeshellarg($order['paket']       ?? 'paket3'),
    escapeshellarg($order['font']        ?? 'Times New Roman'),
    escapeshellarg($order['size']        ?? '12 pt'),
    escapeshellarg($order['hidden']      ?? 'Ya'),
    escapeshellarg($order['posisi']      ?? 'Tengah Bawah'),
    escapeshellarg($order['pos_bab']     ?? 'Tengah Bawah'),
    escapeshellarg($order['pos_isi']     ?? 'Kanan Atas'),
    escapeshellarg($order['dimulai']     ?? 'i'),
    escapeshellarg($order['semb_dafus']  ?? 'Tidak'),
    escapeshellarg($order['semb_lamprn'] ?? 'Tidak'),
    escapeshellarg((string)($order['num_cover'] ?? 1)),
]);

$cmd = '"' . $python . '" "' . $script . '" ' . $args . ' 2>&1';

adk_log('processing', 'Bot process mulai', ['phone' => $phone, 'file' => $filename]);

exec($cmd, $out, $exitCode);
$rawOutput = implode('', $out);
$result    = @json_decode($rawOutput, true);

if ($exitCode !== 0 || !file_exists($outputPath)) {
    $errCode = $result['code']    ?? 'PROCESSING_ERROR';
    $errMsg  = $result['message'] ?? $rawOutput;
    adk_log('error', 'Bot process gagal', ['phone' => $phone, 'code' => $errCode]);
    http_response_code(500);
    echo json_encode(['error' => $errMsg ?: 'Gagal memproses file']);
    exit;
}

adk_log('processing', 'Bot process selesai', ['phone' => $phone, 'output' => $outputName]);

// Tandai order selesai diproses
try {
    $db = getDB();
    $db->prepare("UPDATE bot_pending_orders SET order_done = 1 WHERE phone = ?")
       ->execute([$phone]);
} catch (Exception $e) {
    adk_log('error', 'Gagal set order_done', ['err' => $e->getMessage()]);
}

// Buat order di tabel orders
$hargaMap         = ['paket1'=>3000,'paket2'=>5000,'paket3'=>8000,'paket4'=>10000];
$paketKey         = $order['paket'] ?? 'paket3';
$hargaPaket       = $hargaMap[$paketKey] ?? 10000;
$biayaOperasional = 2000;
$harga            = $hargaPaket + $biayaOperasional;
$email            = $order['email'] ?? $phone;
$orderId          = 'ADK-BOT-' . date('YmdHis') . '-' . strtoupper(bin2hex(random_bytes(3)));
$outputRel        = 'uploads/bot/hasil/' . $outputName;

$extraData = json_encode([
    'pos_bab'      => $order['pos_bab']    ?? 'Tengah Bawah',
    'pos_isi_bab'  => $order['pos_isi']    ?? 'Kanan Atas',
    'semb_dafus'   => $order['semb_dafus'] ?? 'Tidak',
    'semb_lamprn'  => $order['semb_lamprn']?? 'Tidak',
    'num_cover'    => $order['num_cover']  ?? 1,
    'dimulai_dari' => $order['dimulai']    ?? 'i',
    'via'          => 'bot',
    'harga_paket'  => $hargaPaket,
    'bot_surcharge'=> $biayaOperasional,
]);

try {
    $db = getDB();
    $db->prepare("
        INSERT INTO orders
            (order_id, phone, file_input, file_output, paket, font, size, hidden_cover, posisi, harga, status, extra_data, created_at)
        VALUES
            (:order_id, :phone, :file_input, :file_output, :paket, :font, :size, :hidden_cover, :posisi, :harga, 'pending', :extra_data, NOW())
    ")->execute([
        ':order_id'    => $orderId,
        ':phone'       => $email,
        ':file_input'  => $origName,
        ':file_output' => $outputRel,
        ':paket'       => $paketKey,
        ':font'        => $order['font']   ?? 'Times New Roman',
        ':size'        => $order['size']   ?? '12 pt',
        ':hidden_cover'=> $order['hidden'] ?? 'Ya',
        ':posisi'      => $order['posisi'] ?? 'Tengah Bawah',
        ':harga'       => $harga,
        ':extra_data'  => $extraData,
    ]);
} catch (Exception $e) {
    adk_log('error', 'Gagal simpan order bot', ['err' => $e->getMessage()]);
}

// Generate magic link token (reusable, valid 7 days)
$botToken   = bin2hex(random_bytes(32));
$expiresAt  = date('Y-m-d H:i:s', strtotime('+7 days'));
$historyUrl = APP_URL . '/history.php?bot_token=' . $botToken;

try {
    $db->prepare("
        INSERT INTO bot_login_tokens (token, email, expires_at)
        VALUES (:token, :email, :expires_at)
    ")->execute([
        ':token'      => $botToken,
        ':email'      => $email,
        ':expires_at' => $expiresAt,
    ]);
} catch (Exception $e) {
    adk_log('error', 'Gagal simpan bot_login_token', ['err' => $e->getMessage()]);
    $historyUrl = APP_URL . '/cek_pembelian.php';
}

echo json_encode([
    'success'           => true,
    'order_id'          => $orderId,
    'output_file'       => $outputName,
    'output_url'        => APP_URL . '/uploads/bot/hasil/' . $outputName,
    'history_url'       => $historyUrl,
    'cek_url'           => APP_URL . '/cek_pembelian.php',
    'harga_paket'       => 'Rp ' . number_format($hargaPaket, 0, ',', '.'),
    'biaya_operasional' => 'Rp ' . number_format($biayaOperasional, 0, ',', '.'),
    'total'             => 'Rp ' . number_format($harga, 0, ',', '.'),
]);
