<?php
/**
 * api/proses_akhir.php
 * Simpan metadata job → redirect ke menunggu.php.
 * Python dijalankan via AJAX dari menunggu.php (run_job.php), bukan di sini.
 */
date_default_timezone_set('Asia/Jakarta');
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/log.php';

if (!isset($_SESSION['file'])) {
    header("Location: ../jasa.html");
    exit;
}

$uploadDir  = __DIR__ . '/../upload';
$input_full = $uploadDir . '/' . $_SESSION['file'];
if (!file_exists($input_full)) {
    adk_log('error', 'File input tidak ditemukan', ['file' => $_SESSION['file']]);
    die("File input tidak ditemukan.");
}

// Output path
$namaAsli  = pathinfo($_SESSION['file'], PATHINFO_FILENAME);
$ext       = strtolower(pathinfo($_SESSION['file'], PATHINFO_EXTENSION));
$namaAsli  = preg_replace('/^\d+_/', '', $namaAsli);
$namaAsli  = preg_replace("/[^a-zA-Z0-9]/", "_", $namaAsli);
$namaAsli  = trim(preg_replace("/_+/", "_", $namaAsli), "_");

$timestamp   = date("YmdHis");
$random      = bin2hex(random_bytes(4));
$outputRel   = "hasil/" . $namaAsli . "_ADK_" . $timestamp . "_" . $random . "." . $ext;
$output_full = __DIR__ . '/../' . $outputRel;

$hasilDir = __DIR__ . '/../hasil';
if (!is_dir($hasilDir)) mkdir($hasilDir, 0775, true);

$dbId    = $_SESSION['order_db_id'] ?? null;
$orderId = $_SESSION['order_id']    ?? null;

// Job ID
$jobId = $orderId
    ? preg_replace('/[^a-zA-Z0-9_\-]/', '_', $orderId)
    : 'job_' . uniqid();

// Simpan metadata — dibaca oleh run_job.php
$logDir = __DIR__ . '/../logs';
if (!is_dir($logDir)) mkdir($logDir, 0775, true);

$metaFile = $logDir . '/job_' . $jobId . '_meta.json';
file_put_contents($metaFile, json_encode([
    'db_id'       => $dbId,
    'order_id'    => $orderId,
    'input_full'  => $input_full,
    'output_full' => $output_full,
    'output_rel'  => $outputRel,
    'paket'       => $_SESSION['paket']        ?? 'paket3',
    'font'        => $_SESSION['font']         ?? 'Times New Roman',
    'size'        => $_SESSION['size']         ?? '12 pt',
    'hidden'      => $_SESSION['hidden_cover'] ?? 'Ya',
    'posisi'      => $_SESSION['posisi']       ?? 'Tengah Bawah',
    'pos_bab'     => $_SESSION['pos_bab']      ?? 'Tengah Bawah',
    'pos_isi'     => $_SESSION['pos_isi_bab']  ?? 'Kanan Atas',
    'dimulai'     => $_SESSION['dimulai_dari'] ?? 'i',
    'semb_dafus'  => $_SESSION['semb_dafus']   ?? 'Tidak',
    'semb_lamprn' => $_SESSION['semb_lamprn']  ?? 'Tidak',
    'num_cover'   => (int)($_SESSION['num_cover'] ?? 1),
    'created_at'  => date('Y-m-d H:i:s'),
]));

$_SESSION['current_job_id'] = $jobId;

adk_log('processing', 'Job siap, redirect ke menunggu', ['job' => $jobId, 'order' => $orderId]);

header("Location: ../menunggu.php?job=" . urlencode($jobId) . "&db_id=" . (int)$dbId);
exit;