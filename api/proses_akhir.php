<?php
/**
 * api/proses_akhir.php
 * Menjalankan Python di background (async), lalu redirect ke menunggu.php.
 * User tidak perlu menunggu loading — polling dilakukan dari JavaScript.
 */
date_default_timezone_set('Asia/Jakarta');
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';
require_once __DIR__ . '/../includes/log.php';

// ── Guard ─────────────────────────────────────────────────────────────────────
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

// ── Setup output path ─────────────────────────────────────────────────────────
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

// ── Session params ─────────────────────────────────────────────────────────────
$paket      = $_SESSION['paket']        ?? 'paket3';
$font       = $_SESSION['font']         ?? 'Times New Roman';
$size       = $_SESSION['size']         ?? '12 pt';
$hidden     = $_SESSION['hidden_cover'] ?? 'Ya';
$posisi     = $_SESSION['posisi']       ?? 'Tengah Bawah';
$pos_bab    = $_SESSION['pos_bab']      ?? 'Tengah Bawah';
$pos_isi    = $_SESSION['pos_isi_bab']  ?? 'Kanan Atas';
$dimulai    = $_SESSION['dimulai_dari'] ?? 'i';
$semb_dafus = $_SESSION['semb_dafus']   ?? 'Tidak';
$semb_lamprn= $_SESSION['semb_lamprn']  ?? 'Tidak';
$num_cover  = (int)($_SESSION['num_cover'] ?? 1);
$dbId       = $_SESSION['order_db_id']  ?? null;
$orderId    = $_SESSION['order_id']     ?? null;

// ── Queue: batasi maks 3 job bersamaan ────────────────────────────────────────
$logDir     = __DIR__ . '/../logs';
if (!is_dir($logDir)) mkdir($logDir, 0775, true);

define('MAX_CONCURRENT_JOBS', 3);

function count_active_jobs(string $logDir): int {
    $count = 0;
    foreach (glob($logDir . '/job_*.bat') ?: [] as $f) {
        // Bat file masih ada = job belum selesai
        if (filemtime($f) > time() - 300) { // dalam 5 menit terakhir
            $count++;
        }
    }
    return $count;
}

$activeJobs = count_active_jobs($logDir);
if ($activeJobs >= MAX_CONCURRENT_JOBS) {
    // Tandai sebagai queued di session, polling akan cek ulang
    $_SESSION['job_queued'] = true;
    adk_log('processing', 'Job dimasukkan antrian (server sibuk)', [
        'order' => $orderId, 'active' => $activeJobs
    ]);
    // Tetap lanjut — menunggu.php akan polling dan akan diproses saat slot kosong
}

// ── Build job ID dan file path ────────────────────────────────────────────────
$jobId   = ($orderId ? preg_replace('/[^a-zA-Z0-9_\-]/', '_', $orderId) : 'job_' . uniqid());
$batFile = $logDir . DIRECTORY_SEPARATOR . 'job_' . $jobId . '.bat';
$jobFile = $logDir . DIRECTORY_SEPARATOR . 'job_' . $jobId . '.json';
$tmpFile = $jobFile . '.tmp';

// ── Simpan metadata job ───────────────────────────────────────────────────────
$metaFile = $logDir . '/job_' . $jobId . '_meta.json';
file_put_contents($metaFile, json_encode([
    'db_id'      => $dbId,
    'order_id'   => $orderId,
    'output_rel' => $outputRel,
    'started_at' => date('Y-m-d H:i:s'),
]));

// ── Build command ─────────────────────────────────────────────────────────────
$python = PYTHON_EXE;
$script = PYTHON_SCRIPT_DIR . '/main.py';

$args = implode(' ', [
    escapeshellarg($input_full),
    escapeshellarg($output_full),
    escapeshellarg($paket),
    escapeshellarg($font),
    escapeshellarg($size),
    escapeshellarg($hidden),
    escapeshellarg($posisi),
    escapeshellarg($pos_bab),
    escapeshellarg($pos_isi),
    escapeshellarg($dimulai),
    escapeshellarg($semb_dafus),
    escapeshellarg($semb_lamprn),
    escapeshellarg((string)$num_cover),
]);

// ── Tulis .bat untuk eksekusi background (Windows-reliable) ──────────────────
// Hasil Python ditulis ke .tmp dulu, lalu rename ke .json — menghindari polling
// membaca file yang belum selesai ditulis.
$tmpFileEsc = str_replace('/', '\\', $tmpFile);
$jobFileEsc = str_replace('/', '\\', $jobFile);
$batFileEsc = str_replace('/', '\\', $batFile);

$batContent  = '@echo off' . "\r\n";
$batContent .= '"' . $python . '" "' . $script . '" ' . $args . ' > "' . $tmpFileEsc . '" 2>&1' . "\r\n";
$batContent .= 'move /Y "' . $tmpFileEsc . '" "' . $jobFileEsc . '"' . "\r\n";
$batContent .= 'del "' . $batFileEsc . '"' . "\r\n"; // bat hapus diri sendiri saat selesai

file_put_contents($batFile, $batContent);

// ── Jalankan .bat di background ───────────────────────────────────────────────
pclose(popen('start /B "" "' . $batFile . '"', 'r'));

adk_log('processing', 'Job dimulai (async)', [
    'job'    => $jobId,
    'order'  => $orderId,
    'db_id'  => $dbId,
    'paket'  => $paket,
]);

// ── Simpan job_id ke session ──────────────────────────────────────────────────
$_SESSION['current_job_id'] = $jobId;

// ── Redirect ke halaman menunggu ──────────────────────────────────────────────
header("Location: ../menunggu.php?job=" . urlencode($jobId) . "&db_id=" . (int)$dbId);
exit;