<?php
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

if (file_exists(__DIR__ . '/../config/maintenance.flag') && empty($_SESSION['adk_admin'])) {
    header('Location: ../maintenance.php');
    exit;
}

if (isset($_FILES['file']) && isset($_POST['paket'])) {

    $paket      = $_POST['paket'];
    $font       = $_POST['font']              ?? '';
    $size       = $_POST['size']              ?? '';
    $hidden     = $_POST['hidden_cover']      ?? 'Ya';
    $posisi     = $_POST['posisi']            ?? '';
    $phone      = trim($_POST['phone']        ?? '');
    // Paket 4 — custom per-zona
    $pos_bab    = $_POST['pos_bab']           ?? '';
    $pos_isi    = $_POST['pos_isi_bab']       ?? '';
    $dimulai    = $_POST['dimulai_dari']      ?? 'i';
    $semb_dafus = $_POST['sembunyi_dafus']    ?? 'Tidak';
    $semb_lamprn= $_POST['sembunyi_lamprn']   ?? 'Tidak';

    $isGuest = empty($phone);

    // Mode tamu: gunakan/buat guest_token untuk identifikasi perangkat
    if ($isGuest) {
        $guestToken = $_SESSION['guest_token'] ?? ($_COOKIE['adk_guest'] ?? null);
        if (!$guestToken) {
            $guestToken = 'g_' . bin2hex(random_bytes(20));
        }
        // Simpan cookie 30 hari agar order tetap tampil setelah browser ditutup
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
    if ($paket == 'paket1')     $harga = $hargaConfig['paket1'] ?? 5000;
    elseif ($paket == 'paket2') $harga = $hargaConfig['paket2'] ?? 10000;
    elseif ($paket == 'paket4') $harga = $hargaConfig['paket4'] ?? 15000;
    else                        $harga = $hargaConfig['paket3'] ?? 10000;

    $orderId = 'ADK-' . time() . '-' . bin2hex(random_bytes(3));

    // Paket 4: kumpulkan data extra sebagai JSON
    $extraData = null;
    if ($paket === 'paket4') {
        $extraData = json_encode([
            'pos_bab'     => $pos_bab,
            'pos_isi_bab' => $pos_isi,
            'dimulai_dari'=> $dimulai,
            'semb_dafus'  => $semb_dafus,
            'semb_lamprn' => $semb_lamprn,
        ]);
    }

    try {
        $db   = getDB();
        $stmt = $db->prepare("
            INSERT INTO orders
                (phone, guest_token, order_id, file_input, paket, font, size, hidden_cover, posisi, harga, extra_data)
            VALUES
                (:phone, :guest_token, :order_id, :file_input, :paket, :font, :size, :hidden_cover, :posisi, :harga, :extra_data)
        ");
        $stmt->execute([
            ':phone'        => $isGuest ? null : $phone,
            ':guest_token'  => $guestToken,
            ':order_id'     => $orderId,
            ':file_input'   => $namaBaru,
            ':paket'        => $paket,
            ':font'         => $font,
            ':size'         => $size,
            ':hidden_cover' => $hidden,
            ':posisi'       => $posisi,
            ':extra_data'   => $extraData,
            ':harga'        => $harga,
        ]);
        $dbId = $db->lastInsertId();
    } catch (Exception $e) {
        $dbId = null;
    }

    $_SESSION['file']         = $namaBaru;
    $_SESSION['paket']        = $paket;
    $_SESSION['font']         = $font;
    $_SESSION['size']         = $size;
    $_SESSION['hidden_cover'] = $hidden;
    $_SESSION['posisi']       = $posisi;
    $_SESSION['pos_bab']      = $pos_bab;
    $_SESSION['pos_isi_bab']  = $pos_isi;
    $_SESSION['dimulai_dari'] = $dimulai;
    $_SESSION['semb_dafus']   = $semb_dafus;
    $_SESSION['semb_lamprn']  = $semb_lamprn;
    $_SESSION['phone']        = $isGuest ? null : $phone;
    $_SESSION['guest_token']  = $guestToken;
    $_SESSION['is_guest']     = $isGuest;
    // Hapus sisa session nomor lama agar history tidak tumpang tindih dengan mode tamu
    if ($isGuest) {
        unset($_SESSION['cek_phone']);
    }
    $_SESSION['order_id']     = $orderId;
    $_SESSION['order_db_id']  = $dbId;

    header("Location: proses_akhir.php");
    exit;

} else {
    die("Data tidak lengkap!");
}
