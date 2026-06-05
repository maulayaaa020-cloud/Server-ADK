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

$data       = json_decode(file_get_contents('php://input'), true);
$phone      = trim($data['phone']    ?? '');
$filename   = trim($data['filename'] ?? '');

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

// Siapkan path output
$hasilDir  = __DIR__ . '/../uploads/bot/hasil/';
if (!is_dir($hasilDir)) mkdir($hasilDir, 0755, true);

$outputName = 'hasil_' . basename($filename);
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

echo json_encode([
    'success'     => true,
    'output_file' => $outputName,
    'output_url'  => APP_URL . '/uploads/bot/hasil/' . $outputName,
]);
