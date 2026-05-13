<?php
require_once 'vendor/autoload.php';
require_once 'config.php';

// konfigurasi Midtrans
\Midtrans\Config::$serverKey = MIDTRANS_SERVER_KEY;
\Midtrans\Config::$isProduction = false; // sandbox
\Midtrans\Config::$isSanitized = true;
\Midtrans\Config::$is3ds = true;

// contoh harga (sementara)
$harga = 5000;

// buat ID order unik
$order_id = "ADK-" . time();

// data transaksi
$params = array(
    'transaction_details' => array(
        'order_id' => $order_id,
        'gross_amount' => $harga,
    ),
);

// ambil snap token
$snapToken = \Midtrans\Snap::getSnapToken($params);
?>

<!DOCTYPE html>
<html>
<head>
    <title>Pembayaran</title>
</head>
<body>

<h2>Bayar Sekarang</h2>

<button id="pay-button">Bayar</button>

<!-- SCRIPT MIDTRANS -->
<script src="https://app.sandbox.midtrans.com/snap/snap.js"
        data-client-key="<?= MIDTRANS_CLIENT_KEY ?>"></script>

<script>
document.getElementById('pay-button').onclick = function () {
    snap.pay('<?= $snapToken ?>', {
        onSuccess: function(result){
            alert("Pembayaran berhasil!");
            window.location.href = "proses.php";
        },
        onPending: function(result){
            alert("Menunggu pembayaran");
        },
        onError: function(result){
            alert("Pembayaran gagal");
        }
    });
};
</script>

</body>
</html>