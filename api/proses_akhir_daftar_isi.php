<?php
date_default_timezone_set('Asia/Jakarta');
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

if (!isset($_SESSION['di_file'])) {
    header("Location: ../daftar-isi.html");
    exit;
}

$uploadDir  = __DIR__ . '/../upload';
$input_full = $uploadDir . '/' . $_SESSION['di_file'];
if (!file_exists($input_full)) {
    die("File input tidak ditemukan.");
}

$namaAsli  = pathinfo($_SESSION['di_file'], PATHINFO_FILENAME);
$ext       = strtolower(pathinfo($_SESSION['di_file'], PATHINFO_EXTENSION));
$namaAsli  = preg_replace('/^\d+_/', '', $namaAsli);
$namaAsli  = preg_replace("/[^a-zA-Z0-9]/", "_", $namaAsli);
$namaAsli  = trim(preg_replace("/_+/", "_", $namaAsli), "_");

$timestamp   = date("YmdHis");
$random      = bin2hex(random_bytes(4));
$outputRel   = "hasil/" . $namaAsli . "_DaftarIsi_" . $timestamp . "_" . $random . "." . $ext;
$output_full = __DIR__ . '/../' . $outputRel;

$kedalaman    = $_SESSION['di_kedalaman']    ?? 'H1';
$format_titik = $_SESSION['di_format_titik'] ?? 'titik';

$python = PYTHON_EXE;
$script = PYTHON_SCRIPT_DIR . '/daftar_isi.py';

$hasilDir = __DIR__ . '/../hasil';
if (!is_dir($hasilDir)) {
    mkdir($hasilDir, 0775, true);
}

$cmd = "\"$python\" \"$script\" " .
    escapeshellarg($input_full)  . " " .
    escapeshellarg($output_full) . " " .
    escapeshellarg($kedalaman)   . " " .
    escapeshellarg($format_titik) . " 2>&1";

exec($cmd, $out, $status);

if ($status !== 0 || !file_exists($output_full)) {
    $logDir = __DIR__ . '/../logs';
    if (!is_dir($logDir)) mkdir($logDir, 0775, true);
    $logLine = date('Y-m-d H:i:s') . " | daftar_isi | order=" . ($_SESSION['di_order_db_id'] ?? 'unknown')
             . " | status={$status}\n" . implode("\n", $out) . "\n---\n";
    file_put_contents($logDir . '/error_daftar_isi.log', $logLine, FILE_APPEND | LOCK_EX);

    echo "<div style='font-family:sans-serif;padding:40px;text-align:center;color:white;background:#0d0b1e;min-height:100vh'>";
    echo "<h2 style='color:#f87171'>Gagal memproses file</h2>";
    echo "<p style='color:#9ca3af'>Terjadi kesalahan saat membuat daftar isi.<br>Pastikan dokumen menggunakan style Heading 1/2/3 dari Word.</p>";
    echo "<a href='../daftar-isi.html' style='color:#a78bfa'>← Coba Lagi</a>";
    echo "</div>";
    exit;
}

$dbId = $_SESSION['di_order_db_id'] ?? null;
if ($dbId) {
    try {
        $db = getDB();
        $db->prepare("UPDATE daftar_isi_orders SET file_output = :out, status = 'selesai' WHERE id = :id")
           ->execute([':out' => $outputRel, ':id' => $dbId]);
    } catch (Exception $e) {}
}

$_SESSION['di_file_output'] = $outputRel;

header("Location: ../daftar-isi-hasil.php");
exit;
