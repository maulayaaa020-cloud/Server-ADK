<?php
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

$phone      = $_SESSION['phone'] ?? $_SESSION['cek_phone'] ?? null;
if (empty($phone)) $phone = null;
if (!empty($_SESSION['is_guest'])) $phone = null;
$guestToken = $_SESSION['guest_token'] ?? $_COOKIE['adk_guest'] ?? null;

if (!$phone && !$guestToken) { http_response_code(401); exit; }

$id = (int)($_GET['id'] ?? 0);
if (!$id) { http_response_code(400); exit; }

try {
    $db = getDB();
    if ($phone) {
        $stmt = $db->prepare("SELECT * FROM orders WHERE id = :id AND phone = :phone");
        $stmt->execute([':id' => $id, ':phone' => $phone]);
    } else {
        $stmt = $db->prepare("SELECT * FROM orders WHERE id = :id AND guest_token = :gt");
        $stmt->execute([':id' => $id, ':gt' => $guestToken]);
    }
    $order = $stmt->fetch();
} catch (Exception $e) { http_response_code(500); exit; }

if (!$order || empty($order['file_output_pdf'])) {
    http_response_code(404); exit;
}

$pdf_full = __DIR__ . '/../' . $order['file_output_pdf'];
if (!file_exists($pdf_full)) {
    http_response_code(404); exit;
}

// Folder preview watermark
$preview_dir  = __DIR__ . '/../hasil/preview';
if (!is_dir($preview_dir)) mkdir($preview_dir, 0775, true);

$preview_name = 'wm_' . $id . '.pdf';
$preview_full = $preview_dir . '/' . $preview_name;

// Generate watermark jika belum ada atau PDF asli lebih baru
$perlu_generate = !file_exists($preview_full)
               || filemtime($pdf_full) > filemtime($preview_full);

if ($perlu_generate) {
    $python = PYTHON_EXE;
    $script = PYTHON_SCRIPT_DIR . '/add_watermark.py';
    $cmd    = "\"$python\" \"$script\" " .
              escapeshellarg($pdf_full) . " " .
              escapeshellarg($preview_full) . " 2>&1";
    exec($cmd, $out, $status);

    if ($status !== 0 || !file_exists($preview_full)) {
        // Fallback: serve PDF asli jika watermark gagal
        $preview_full = $pdf_full;
    }
}

// Stream PDF ke browser (inline — supaya iframe bisa tampilkan)
header('Content-Type: application/pdf');
header('Content-Disposition: inline; filename="preview.pdf"');
header('Content-Length: ' . filesize($preview_full));
header('Cache-Control: private, max-age=600');
header('X-Frame-Options: SAMEORIGIN');
readfile($preview_full);
exit;
