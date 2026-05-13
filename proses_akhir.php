<?php
date_default_timezone_set('Asia/Jakarta');
session_start();
require_once 'config.php';
require_once 'db.php';

if (!isset($_SESSION['file'])) {
    header("Location: jasa.html");
    exit;
}

$input_full = __DIR__ . "/upload/" . $_SESSION['file'];
if (!file_exists($input_full)) {
    die("File input tidak ditemukan.");
}

// Nama output
$namaAsli  = pathinfo($_SESSION['file'], PATHINFO_FILENAME);
$ext       = strtolower(pathinfo($_SESSION['file'], PATHINFO_EXTENSION));
$namaAsli  = preg_replace('/^\d+_/', '', $namaAsli);
$namaAsli  = preg_replace("/[^a-zA-Z0-9]/", "_", $namaAsli);
$namaAsli  = trim(preg_replace("/_+/", "_", $namaAsli), "_");

$timestamp   = date("YmdHis");
$random      = bin2hex(random_bytes(4));
$outputRel   = "hasil/" . $namaAsli . "_ADK_" . $timestamp . "_" . $random . "." . $ext;
$output_full = __DIR__ . "/" . $outputRel;

$paket  = $_SESSION['paket']        ?? 'paket3';
$font   = $_SESSION['font']         ?? 'Times New Roman';
$size   = $_SESSION['size']         ?? '12 pt';
$hidden = $_SESSION['hidden_cover'] ?? 'Ya';
$posisi = $_SESSION['posisi']       ?? 'Tengah Bawah';

$python = "D:\\Freelaces\\Server\\python.exe";
$script = "D:\\Freelaces\\Server\\htdocs\\adk\\python\\main.py";

if (!is_dir(__DIR__ . "/hasil")) {
    mkdir(__DIR__ . "/hasil", 0775, true);
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
    echo "<pre style='color:red;padding:20px'>";
    echo "❌ Gagal memproses file.\n\n";
    print_r($out);
    echo "</pre>";
    exit;
}

// Simpan nama output ke DB
$dbId = $_SESSION['order_db_id'] ?? null;
if ($dbId) {
    try {
        $db = getDB();
        $db->prepare("UPDATE orders SET file_output = :out WHERE id = :id")
           ->execute([':out' => $outputRel, ':id' => $dbId]);
    } catch (Exception $e) {}
}

// Simpan ke session agar history.php bisa pakai
$_SESSION['file_output'] = $outputRel;

// Redirect ke history
header("Location: history.php");
exit;
