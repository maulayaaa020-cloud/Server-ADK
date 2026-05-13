<?php
session_start();

require_once 'vendor/autoload.php';
require_once 'config.php';

if(!isset($_SESSION['file'])){
    die("Akses tidak valid!");
}

$file = $_SESSION['file'];
$paket = $_SESSION['paket'];

if($paket == "paket1") $harga = 10000;
elseif($paket == "paket2") $harga = 12000;
elseif($paket == "paket3") $harga = 13000;
else $harga = 10000;

\Midtrans\Config::$serverKey = MIDTRANS_SERVER_KEY;
\Midtrans\Config::$isProduction = false;
\Midtrans\Config::$isSanitized = true;
\Midtrans\Config::$is3ds = true;

$order_id = "ADK-" . time();

$params = [
    'transaction_details' => [
        'order_id' => $order_id,
        'gross_amount' => $harga,
    ],
    'customer_details' => [
        'first_name' => 'User',
        'email' => 'user@gmail.com',
    ]
];

$snapToken = \Midtrans\Snap::getSnapToken($params);
?>

<!DOCTYPE html>
<html>
<head>
    <title>Pembayaran</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>

<div class="hero">
    <div class="hero-right">
        <div class="card">

            <h3>Ringkasan Pesanan</h3>

            <p><b>File:</b> <?= $file ?></p>
            <p><b>Paket:</b> <?= $paket ?></p>
            <p><b>Harga:</b> Rp <?= number_format($harga) ?></p>

            <hr>

            <button id="pay-button" class="btn-Upload">
                Bayar Sekarang
            </button>

        </div>
    </div>
</div>

<script src="https://app.sandbox.midtrans.com/snap/snap.js"
data-client-key="<?= MIDTRANS_CLIENT_KEY ?>"></script>

<script>
document.getElementById('pay-button').onclick = function () {
    snap.pay('<?= $snapToken ?>', {
        onSuccess: function(result){
            window.location.href = "/adk/proses_akhir.php";
        },
        onPending: function(result){
            alert("Menunggu pembayaran");
        },
        onError: function(result){
            alert("Gagal pembayaran");
            console.log(result);
        }
    });
};
</script>

</body>
</html>