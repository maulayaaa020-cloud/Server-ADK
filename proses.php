<?php
session_start()
require_once 'config.php';
require_once 'db.php';

if (isset($_FILES['file']) && isset($_POST['paket'])) {

    $paket  = $_POST['paket'];
    $font   = $_POST['font']         ?? '';
    $size   = $_POST['size']         ?? '';
    $hidden = $_POST['hidden_cover'] ?? 'Ya';
    $posisi = $_POST['posisi']       ?? '';
    $phone  = trim($_POST['phone']   ?? '');

    $namaFile = $_FILES['file']['name'];
    $tmpFile  = $_FILES['file']['tmp_name'];

    $ext = strtolower(pathinfo($namaFile, PATHINFO_EXTENSION));
    if (!in_array($ext, ['doc', 'docx'])) {
        die("Format harus doc/docx");
    }

    if (!is_dir("upload")) {
        mkdir("upload", 0775, true);
    }

    $namaBaru  = time() . "_" . preg_replace("/[^a-zA-Z0-9._]/", "", $namaFile);
    move_uploaded_file($tmpFile, "upload/" . $namaBaru);

    if ($paket == 'paket1')     $harga = 310000
    elseif ($paket == 'paket2') $harga = 512000
    else                        $harga = 715000

    $orderId = 'ADK-' . time() . '-' . bin2hex(random_bytes(3));

    // Simpan ke database
    try {
        $db = getDB();
        $stmt = $db->prepare("
            INSERT INTO orders
                (phone, order_id, file_input, paket, font, size, hidden_cover, posisi, harga)
            VALUES
                (:phone, :order_id, :file_input, :paket, :font, :size, :hidden_cover, :posisi, :harga)
        ");
        $stmt->execute([
            ':phone'        => $phone,
            ':order_id'     => $orderId,
            ':file_input'   => $namaBaru,
            ':paket'        => $paket,
            ':font'         => $font,
            ':size'         => $size,
            ':hidden_cover' => $hidden,
            ':posisi'       => $posisi,
            ':harga'        => $harga,
        ]);
        $dbId = $db->lastInsertId();
    } catch (Exception $e) {
        $dbId = null;
    }

    // Simpan session
    $_SESSION['file']         = $namaBaru;
    $_SESSION['paket']        = $paket;
    $_SESSION['font']         = $font;
    $_SESSION['size']         = $size;
    $_SESSION['hidden_cover'] = $hidden;
    $_SESSION['posisi']       = $posisi;
    $_SESSION['phone']        = $phone;
    $_SESSION['order_id']     = $orderId;
    $_SESSION['order_db_id']  = $dbId;

    header("Location: proses_akhir.php");
    exit;

} else {
    die("Data tidak lengkap!");
}
se {
    die("Data tidak lengkap!");
} {
    die("Data tidak lengkap!");
}{
    die("Data tidak lengkap!");
}