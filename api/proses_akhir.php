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

$paket  = $_SESSION['paket']        ?? 'paket3';
$font   = $_SESSION['font']         ?? 'Times New Roman';
$size   = $_SESSION['size']         ?? '12 pt';
$hidden = $_SESSION['hidden_cover'] ?? 'Ya';
$posisi = $_SESSION['posisi']       ?? 'Tengah Bawah';

$python = PYTHON_EXE;
$script = PYTHON_SCRIPT_DIR . '\\proses.py';

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
    escapeshellarg($posisi)      . " 2>&1";

exec($cmd, $out, $status);

if ($status !== 0 || !file_exists($output_full)) {
    $logDir = __DIR__ . '/../logs';
    if (!is_dir($logDir)) mkdir($logDir, 0775, true);
    $logLine = date('Y-m-d H:i:s') . " | order={$orderId} | status={$status}\n" . implode("\n", $out) . "\n---\n";
    file_put_contents($logDir . '/error.log', $logLine, FILE_APPEND | LOCK_EX);

    echo "<div style='font-family:sans-serif;padding:40px;text-align:center;color:white;background:#0d0b1e;min-height:100vh'>";
    echo "<h2 style='color:#f87171'>Gagal memproses file</h2>";
    echo "<p style='color:#9ca3af'>Terjadi kesalahan saat mengerjakan dokumen Anda.<br>Silakan coba lagi atau hubungi admin.</p>";
    echo "<a href='../jasa.html' style='color:#a78bfa'>← Kembali</a>";
    echo "</div>";
    exit;
}

$dbId = $_SESSION['order_db_id'] ?? null;

// ── Konversi ke PDF via Microsoft Word ──────────────────
$pdfRel      = preg_replace('/\.(docx?)$/i', '.pdf', $outputRel);
$pdf_full    = __DIR__ . '/../' . $pdfRel;
$pdfScript   = "D:\\Freelaces\\Server\\htdocs\\adk\\python\\convert_pdf.py";
$pdfCmd      = "\"$python\" \"$pdfScript\" " .
               escapeshellarg($output_full) . " " .
               escapeshellarg($pdf_full) . " 2>&1";
exec($pdfCmd, $pdfOut, $pdfStatus);
$pdfOk = ($pdfStatus === 0 && file_exists($pdf_full));

// ── Simpan ke DB ─────────────────────────────────────────
if ($dbId) {
    try {
        $db = getDB();
        $db->prepare("UPDATE orders SET file_output = :out, file_output_pdf = :pdf WHERE id = :id")
           ->execute([
               ':out' => $outputRel,
               ':pdf' => $pdfOk ? $pdfRel : null,
               ':id'  => $dbId,
           ]);
    } catch (Exception $e) {}
}

$_SESSION['file_output']     = $outputRel;
$_SESSION['file_output_pdf'] = $pdfOk ? $pdfRel : null;

header("Location: ../history.php");
exit;
