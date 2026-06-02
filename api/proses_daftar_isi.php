<?php
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';
require_once __DIR__ . '/../includes/log.php';

if (file_exists(__DIR__ . '/../config/maintenance_daftar_isi.flag') && empty($_SESSION['adk_admin'])) {
    header('Content-Type: application/json');
    echo json_encode(['error' => true, 'message' => 'Layanan Daftar Isi sedang maintenance. Coba lagi nanti.']);
    exit;
}

if (!isset($_FILES['file']) || !isset($_POST['kedalaman']) || !isset($_POST['format_titik'])) {
    die("Data tidak lengkap!");
}

$kedalaman    = $_POST['kedalaman'];
$format_titik = $_POST['format_titik'];
$font         = $_POST['font']  ?? 'Times New Roman';
$size         = $_POST['size']  ?? '12';
$space        = $_POST['space'] ?? '1.0';
$phone        = trim($_POST['phone'] ?? '');

$valid_kedalaman = ['H1', 'H1+H2', 'H1+H2+H3'];
$valid_format    = ['titik', 'tab', 'kecualikan_bab'];
$valid_font      = ['Times New Roman', 'Arial', 'Calibri'];
$valid_size      = ['11', '12', '14'];
$valid_space     = ['1.0', '1.15', '1.5', '2.0'];

if (!in_array($kedalaman, $valid_kedalaman) || !in_array($format_titik, $valid_format)
    || !in_array($font, $valid_font) || !in_array($size, $valid_size)
    || !in_array($space, $valid_space)) {
    die("Pilihan tidak valid.");
}

$isGuest = empty($phone);

if ($isGuest) {
    $guestToken = $_SESSION['guest_token'] ?? ($_COOKIE['adk_guest'] ?? null);
    if (!$guestToken) {
        $guestToken = 'g_' . bin2hex(random_bytes(20));
    }
    setcookie('adk_guest', $guestToken, time() + 30 * 86400, '/', '', false, true);
} else {
    $guestToken = null;
}

$namaFile = $_FILES['file']['name'];
$tmpFile  = $_FILES['file']['tmp_name'];

$ext = strtolower(pathinfo($namaFile, PATHINFO_EXTENSION));
if (!in_array($ext, ['doc', 'docx'])) {
    die("Format harus doc/docx");
}

if ($_FILES['file']['size'] > 30 * 1024 * 1024) {
    die("Ukuran file maksimal 30MB.");
}

$finfo = finfo_open(FILEINFO_MIME_TYPE);
$mime  = finfo_file($finfo, $tmpFile);
finfo_close($finfo);
$allowedMime = [
    'application/msword',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
];
if (!in_array($mime, $allowedMime)) {
    die("Format file tidak valid.");
}

$uploadDir = __DIR__ . '/../upload';
if (!is_dir($uploadDir)) {
    mkdir($uploadDir, 0775, true);
}

$namaBaru = time() . "_" . preg_replace("/[^a-zA-Z0-9._]/", "", $namaFile);
if (!move_uploaded_file($tmpFile, $uploadDir . "/" . $namaBaru)) {
    die("Gagal menyimpan file. Coba lagi.");
}

$hargaConfig = @json_decode(@file_get_contents(__DIR__ . '/../config/harga.json'), true) ?: [];
$harga       = $hargaConfig['daftar_isi'] ?? 10000;

$orderId = 'DI-' . time() . '-' . bin2hex(random_bytes(3));

try {
    $db   = getDB();
    $stmt = $db->prepare("
        INSERT INTO daftar_isi_orders
            (phone, guest_token, order_id, file_input, kedalaman, format_titik, harga)
        VALUES
            (:phone, :guest_token, :order_id, :file_input, :kedalaman, :format_titik, :harga)
    ");
    $stmt->execute([
        ':phone'        => $isGuest ? null : $phone,
        ':guest_token'  => $guestToken,
        ':order_id'     => $orderId,
        ':file_input'   => $namaBaru,
        ':kedalaman'    => $kedalaman,
        ':format_titik' => $format_titik,
        ':harga'        => $harga,
    ]);
    $dbId = $db->lastInsertId();
    adk_log('daftar_isi', 'Order dibuat', ['db_id' => $dbId, 'order_id' => $orderId, 'phone' => $phone ?: 'guest']);
} catch (Exception $e) {
    adk_log('error', 'DAFTAR ISI INSERT FAIL', ['err' => $e->getMessage(), 'phone' => $phone]);
    http_response_code(500);
    die("Gagal membuat order. Coba lagi.");
}

$_SESSION['di_file']         = $namaBaru;
$_SESSION['di_kedalaman']    = $kedalaman;
$_SESSION['di_format_titik'] = $format_titik;
$_SESSION['di_font']         = $font;
$_SESSION['di_size']         = $size;
$_SESSION['di_space']        = $space;
$_SESSION['di_phone']        = $isGuest ? null : $phone;
$_SESSION['di_guest_token']  = $guestToken;
$_SESSION['di_is_guest']     = $isGuest;
$_SESSION['di_order_id']     = $orderId;
$_SESSION['di_order_db_id']  = $dbId;

// Set session yang dibutuhkan history.php
if (!$isGuest) {
    $_SESSION['phone'] = $phone;
} else {
    $_SESSION['guest_token'] = $guestToken;
    if (!empty($_SESSION['is_guest'])) {} // jaga flag
    $_SESSION['is_guest'] = true;
}

header("Location: proses_akhir_daftar_isi.php");
exit;
