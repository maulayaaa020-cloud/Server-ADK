<?php
date_default_timezone_set('Asia/Jakarta');
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

if (!isset($_SESSION['file'])) {
    header("Location: ../jasa.html");
    exit;
}

$uploadDir  = __DIR__ . '/../upload';
$input_full = $uploadDir . '/' . $_SESSION['file'];
if (!file_exists($input_full)) {
    die("File input tidak ditemukan.");
}

$namaAsli  = pathinfo($_SESSION['file'], PATHINFO_FILENAME);
$ext       = strtolower(pathinfo($_SESSION['file'], PATHINFO_EXTENSION));
$namaAsli  = preg_replace('/^\d+_/', '', $namaAsli);
$namaAsli  = preg_replace("/[^a-zA-Z0-9]/", "_", $namaAsli);
$namaAsli  = trim(preg_replace("/_+/", "_", $namaAsli), "_");

$timestamp   = date("YmdHis");
$random      = bin2hex(random_bytes(4));
$outputRel   = "hasil/" . $namaAsli . "_ADK_" . $timestamp . "_" . $random . "." . $ext;
$output_full = __DIR__ . '/../' . $outputRel;

$paket      = $_SESSION['paket']        ?? 'paket3';
$font       = $_SESSION['font']         ?? 'Times New Roman';
$size       = $_SESSION['size']         ?? '12 pt';
$hidden     = $_SESSION['hidden_cover'] ?? 'Ya';
$posisi     = $_SESSION['posisi']       ?? 'Tengah Bawah';
// Paket 4 — custom per-zona
$pos_bab    = $_SESSION['pos_bab']      ?? 'Tengah Bawah';
$pos_isi    = $_SESSION['pos_isi_bab']  ?? 'Kanan Atas';
$dimulai    = $_SESSION['dimulai_dari'] ?? 'i';
$semb_dafus = $_SESSION['semb_dafus']   ?? 'Tidak';
$semb_lamprn= $_SESSION['semb_lamprn']  ?? 'Tidak';
$num_cover  = (int)($_SESSION['num_cover'] ?? 1);

$python = PYTHON_EXE;
$script = PYTHON_SCRIPT_DIR . '/main.py';

$hasilDir = __DIR__ . '/../hasil';
if (!is_dir($hasilDir)) {
    mkdir($hasilDir, 0775, true);
}

$cmd = "\"$python\" \"$script\" " .
    escapeshellarg($input_full)  . " " .
    escapeshellarg($output_full) . " " .
    escapeshellarg($paket)       . " " .
    escapeshellarg($font)        . " " .
    escapeshellarg($size)        . " " .
    escapeshellarg($hidden)      . " " .
    escapeshellarg($posisi)      . " " .
    escapeshellarg($pos_bab)     . " " .
    escapeshellarg($pos_isi)     . " " .
    escapeshellarg($dimulai)     . " " .
    escapeshellarg($semb_dafus)  . " " .
    escapeshellarg($semb_lamprn) . " " .
    escapeshellarg($num_cover)   . " 2>&1";

exec($cmd, $out, $status);

if ($status !== 0 || !file_exists($output_full)) {
    $logDir = __DIR__ . '/../logs';
    if (!is_dir($logDir)) mkdir($logDir, 0775, true);
    $logLine = date('Y-m-d H:i:s') . " | order=" . ($_SESSION['order_db_id'] ?? 'unknown') . " | status={$status}\n" . implode("\n", $out) . "\n---\n";
    file_put_contents($logDir . '/error.log', $logLine, FILE_APPEND | LOCK_EX);

    // Parse output Python untuk pesan error spesifik
    $pyJson   = @json_decode(implode('', $out), true);
    $errCode  = $pyJson['code']    ?? '';
    $errMsg   = $pyJson['message'] ?? '';

    $pesanMap = [
        'FORMAT_NOT_SUPPORTED' => 'Sepertinya file kamu format <b>.doc</b> (versi lama). Buka dulu di Microsoft Word, lalu klik <b>Simpan Sebagai → .docx</b>, terus upload lagi ya.',
        'FILE_TOO_LARGE'       => 'Ups, file kamu kebesaran nih — batas maksimalnya 30 MB. Coba hapus gambar yang nggak perlu atau kompres dulu sebelum upload.',
        'INVALID_DOCX'         => 'File-nya kayaknya rusak atau bukan Word yang beneran. Coba buka di Microsoft Word, simpan ulang, terus upload lagi.',
        'MACRO_DETECTED'       => 'File kamu mengandung macro (biasanya format <b>.docm</b>). Simpan ulang sebagai <b>.docx</b> biasa di Word, lalu upload lagi.',
        'FILE_READ_ERROR'      => 'File kamu nggak bisa kebaca, mungkin ada bagian yang rusak (misalnya gambar corrupt). Coba buka di Word → Simpan Sebagai → .docx baru, lalu upload lagi.',
        'PROCESSING_ERROR'     => 'Sistemnya kesulitan memproses file ini. Kemungkinan struktur dokumennya tidak biasa. Kalau masalah berlanjut, hubungi admin ya.',
        'FILE_SAVE_ERROR'      => 'Prosesnya sudah selesai tapi gagal nyimpen hasilnya. Coba ulangi sekali lagi, biasanya langsung berhasil.',
    ];

    $detail = $pesanMap[$errCode] ?? 'Ada yang nggak beres waktu mengerjakan dokumen kamu. Coba upload ulang, atau hubungi admin kalau masalahnya terus muncul.';

    echo "<div style='font-family:sans-serif;padding:40px;text-align:center;color:white;background:#0d0b1e;min-height:100vh'>";
    echo "<h2 style='color:#f87171'>Gagal memproses file</h2>";
    echo "<p style='color:#9ca3af;max-width:480px;margin:0 auto 16px'>{$detail}</p>";
    echo "<a href='../jasa.html' style='color:#a78bfa'>← Kembali</a>";
    echo "</div>";
    exit;
}

$dbId = $_SESSION['order_db_id'] ?? null;

// ── Simpan ke DB ─────────────────────────────────────────
if ($dbId) {
    try {
        $db = getDB();
        $db->prepare("UPDATE orders SET file_output = :out WHERE id = :id")
           ->execute([':out' => $outputRel, ':id' => $dbId]);
    } catch (Exception $e) {}
}

$_SESSION['file_output'] = $outputRel;

header("Location: ../history.php");
exit;
