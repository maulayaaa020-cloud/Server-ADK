<?php
date_default_timezone_set('Asia/Jakarta');
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

header('Content-Type: application/json');

$data  = json_decode(file_get_contents('php://input'), true);
$phone = $_SESSION['phone'] ?? $_SESSION['cek_phone'] ?? '-';

$orderId   = trim($data['order_id']  ?? '');
$fileName  = trim($data['file_name'] ?? '');
$paket     = trim($data['paket']     ?? '');
$jenis     = trim($data['jenis']     ?? '');
$deskripsi = trim($data['deskripsi'] ?? '');

if (!$orderId || !$jenis) {
    echo json_encode(['ok' => false, 'error' => 'Data tidak lengkap']);
    exit;
}

try {
    $db = getDB();

    // Cegah laporan duplikat untuk order yang sama
    $cek = $db->prepare("SELECT COUNT(*) FROM bug_reports WHERE order_id = :oid");
    $cek->execute([':oid' => $orderId]);
    if ((int)$cek->fetchColumn() > 0) {
        echo json_encode(['ok' => false, 'error' => 'Laporan untuk order ini sudah pernah dikirim.']);
        exit;
    }

    // Ambil path file dari database
    $row = $db->prepare("SELECT file_input, file_output, file_output_pdf FROM orders WHERE order_id = :oid LIMIT 1");
    $row->execute([':oid' => $orderId]);
    $order = $row->fetch();

    if (!$order) {
        echo json_encode(['ok' => false, 'error' => 'Order tidak ditemukan']);
        exit;
    }

    // Buat folder LAPORAN BUG jika belum ada
    $baseDir    = __DIR__ . '/../LAPORAN BUG';
    if (!is_dir($baseDir)) mkdir($baseDir, 0775, true);

    // Nama subfolder: tanggal_jam_orderid
    $folderName = date('Y-m-d_His') . '_' . preg_replace('/[^a-zA-Z0-9\-]/', '', $orderId);
    $reportDir  = $baseDir . '/' . $folderName;
    mkdir($reportDir, 0775, true);

    $root = __DIR__ . '/../';

    // Salin file input (asli)
    if (!empty($order['file_input'])) {
        $src = $root . 'upload/' . basename($order['file_input']);
        if (file_exists($src)) copy($src, $reportDir . '/' . basename($src));
    }

    // Salin file output docx (hasil)
    if (!empty($order['file_output'])) {
        $src = $root . $order['file_output'];
        if (file_exists($src)) copy($src, $reportDir . '/' . basename($src));
    }

    // Salin file output pdf (hasil)
    if (!empty($order['file_output_pdf'])) {
        $src = $root . $order['file_output_pdf'];
        if (file_exists($src)) copy($src, $reportDir . '/' . basename($src));
    }

    // Buat file catatan.txt
    $catatan = implode("\n", [
        '=== LAPORAN BUG ADK PHOTOCOPY ===',
        'Tanggal   : ' . date('d-m-Y H:i:s'),
        'Order ID  : ' . $orderId,
        'Nama File : ' . $fileName,
        'Paket     : ' . $paket,
        'No. HP    : ' . $phone,
        'Masalah   : ' . $jenis,
        'Keterangan: ' . ($deskripsi ?: '-'),
        '',
        'File asli    : ' . basename($order['file_input'] ?? '-'),
        'File hasil   : ' . basename($order['file_output'] ?? '-'),
        'File PDF     : ' . basename($order['file_output_pdf'] ?? '-'),
    ]);
    file_put_contents($reportDir . '/catatan.txt', $catatan);

    // Simpan ke database
    $db->prepare("INSERT INTO bug_reports (order_id, file_name, paket, jenis, deskripsi, phone)
                  VALUES (:oid, :fn, :pk, :jn, :ds, :ph)")
       ->execute([
           ':oid' => $orderId,
           ':fn'  => $fileName,
           ':pk'  => $paket,
           ':jn'  => $jenis,
           ':ds'  => $deskripsi,
           ':ph'  => $phone,
       ]);

    echo json_encode(['ok' => true]);
} catch (Exception $e) {
    echo json_encode(['ok' => false, 'error' => $e->getMessage()]);
}
