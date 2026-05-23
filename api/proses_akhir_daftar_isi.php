<?php
/**
 * api/proses_akhir_daftar_isi.php
 * Simpan metadata job daftar isi → redirect ke menunggu.php.
 * Python dijalankan via AJAX dari menunggu.php (run_job_daftar_isi.php).
 */
date_default_timezone_set('Asia/Jakarta');
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/log.php';

if (!isset($_SESSION['di_file'])) {
    header("Location: ../daftar-isi.html");
    exit;
}

$uploadDir  = __DIR__ . '/../upload';
$input_full = $uploadDir . '/' . $_SESSION['di_file'];
if (!file_exists($input_full)) {
    adk_log('error', 'File input daftar isi tidak ditemukan', ['file' => $_SESSION['di_file']]);
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

$hasilDir = __DIR__ . '/../hasil';
if (!is_dir($hasilDir)) mkdir($hasilDir, 0775, true);

$dbId    = $_SESSION['di_order_db_id'] ?? null;
$orderId = $_SESSION['di_order_id']    ?? null;

$jobId = $orderId
    ? preg_replace('/[^a-zA-Z0-9_\-]/', '_', $orderId)
    : 'job_di_' . uniqid();

$logDir = __DIR__ . '/../logs';
if (!is_dir($logDir)) mkdir($logDir, 0775, true);

$metaFile = $logDir . '/job_' . $jobId . '_meta.json';
file_put_contents($metaFile, json_encode([
    'type'        => 'daftar_isi',
    'db_id'       => $dbId,
    'order_id'    => $orderId,
    'input_full'  => $input_full,
    'output_full' => $output_full,
    'output_rel'  => $outputRel,
    'kedalaman'   => $_SESSION['di_kedalaman']    ?? 'H1',
    'format_titik'=> $_SESSION['di_format_titik'] ?? 'titik',
    'font'        => $_SESSION['di_font']         ?? 'Times New Roman',
    'size'        => $_SESSION['di_size']         ?? '12',
    'space'       => $_SESSION['di_space']        ?? '1.0',
    'created_at'  => date('Y-m-d H:i:s'),
]));

$_SESSION['current_job_id'] = $jobId;

adk_log('daftar_isi', 'Job siap, redirect ke menunggu', ['job' => $jobId, 'order' => $orderId]);

header("Location: ../menunggu.php?job=" . urlencode($jobId) . "&db_id=" . (int)$dbId . "&type=daftar_isi");
exit;