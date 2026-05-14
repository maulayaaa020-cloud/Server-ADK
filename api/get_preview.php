<?php
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';
header('Content-Type: application/json');

$isAdmin    = !empty($_SESSION['adk_admin']);
$phone      = $_SESSION['phone'] ?? $_SESSION['cek_phone'] ?? null;
$guestToken = $_SESSION['guest_token'] ?? $_COOKIE['adk_guest'] ?? null;

if (!$isAdmin && !$phone && !$guestToken) {
    echo json_encode(['ok' => false, 'error' => 'Unauthorized']); exit;
}

$input = json_decode(file_get_contents('php://input'), true);
$db_id = (int)($input['db_id'] ?? 0);
if (!$db_id) { echo json_encode(['ok' => false, 'error' => 'Invalid ID']); exit; }

try {
    $db = getDB();
    if ($isAdmin) {
        $stmt = $db->prepare("SELECT * FROM orders WHERE id = ?");
        $stmt->execute([$db_id]);
    } elseif ($phone) {
        $stmt = $db->prepare("SELECT * FROM orders WHERE id = ? AND phone = ?");
        $stmt->execute([$db_id, $phone]);
    } else {
        $stmt = $db->prepare("SELECT * FROM orders WHERE id = ? AND guest_token = ?");
        $stmt->execute([$db_id, $guestToken]);
    }
    $order = $stmt->fetch();
} catch (Exception $e) {
    echo json_encode(['ok' => false, 'error' => 'Database error']); exit;
}

if (!$order || empty($order['file_output'])) {
    echo json_encode(['ok' => false, 'error' => 'Order tidak ditemukan']); exit;
}

$docxPath = __DIR__ . '/../' . $order['file_output'];
if (!file_exists($docxPath)) {
    echo json_encode(['ok' => false, 'error' => 'File belum tersedia']); exit;
}

$cacheDir = __DIR__ . '/../preview_cache/' . $db_id;

// Kembalikan cache jika sudah ada
if (is_dir($cacheDir)) {
    $pages = glob($cacheDir . '/page-*.png');
    natsort($pages);
    $pages = array_values($pages);
    if (count($pages) > 0) {
        echo json_encode(['ok' => true, 'pages' => count($pages)]); exit;
    }
}

if (!is_dir($cacheDir)) mkdir($cacheDir, 0755, true);

// Dir temp per proses agar tidak konflik antar request
$tmpDir = sys_get_temp_dir() . '/adk_pv_' . $db_id . '_' . getmypid();
if (!is_dir($tmpDir)) mkdir($tmpDir, 0755, true);

// Salin file ke tmpDir agar LibreOffice tidak menyentuh file asli
$tmpDocx = $tmpDir . '/' . basename($docxPath);
if (!copy($docxPath, $tmpDocx)) {
    echo json_encode(['ok' => false, 'error' => 'Gagal menyalin file untuk konversi']); exit;
}

// Step 1: DOCX → PDF via LibreOffice (dari salinan, bukan file asli)
$cmd1 = 'HOME=/tmp libreoffice --headless --convert-to pdf '
      . escapeshellarg($tmpDocx) . ' --outdir ' . escapeshellarg($tmpDir)
      . ' 2>&1';
exec($cmd1, $out1, $ret1);

if ($ret1 !== 0) {
    @rmdir($tmpDir);
    echo json_encode(['ok' => false, 'error' => 'Gagal konversi ke PDF', 'log' => implode("\n", $out1)]); exit;
}

$pdfFiles = glob($tmpDir . '/*.pdf');
if (empty($pdfFiles)) {
    echo json_encode(['ok' => false, 'error' => 'PDF tidak ditemukan setelah konversi']); exit;
}
$pdfPath = $pdfFiles[0];

// Step 2: PDF → PNG per halaman via pdftoppm
$pageBase = $cacheDir . '/page';
$cmd2 = 'pdftoppm -r 130 -png ' . escapeshellarg($pdfPath) . ' ' . escapeshellarg($pageBase) . ' 2>&1';
exec($cmd2, $out2, $ret2);

// Bersihkan file temp
@unlink($pdfPath);
@unlink($tmpDocx);
@rmdir($tmpDir);

if ($ret2 !== 0) {
    echo json_encode(['ok' => false, 'error' => 'Gagal konversi ke gambar', 'log' => implode("\n", $out2)]); exit;
}

$pages = glob($cacheDir . '/page-*.png');
natsort($pages);
$pages = array_values($pages);

if (count($pages) === 0) {
    echo json_encode(['ok' => false, 'error' => 'Tidak ada halaman yang dihasilkan']); exit;
}

echo json_encode(['ok' => true, 'pages' => count($pages)]);
