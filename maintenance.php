<?php
// Jika maintenance tidak aktif, redirect ke halaman utama
if (!file_exists(__DIR__ . '/maintenance.flag')) {
    header('Location: jasa.html');
    exit;
}
?>
<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Server Maintenance — ADK Photocopy</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'Segoe UI', sans-serif;
    background: linear-gradient(135deg, #0e0b28, #1a1043);
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    color: white;
    padding: 24px;
  }
  .card {
    background: rgba(255,255,255,0.05);
    border: 1px solid rgba(239,68,68,0.3);
    border-radius: 20px;
    padding: 48px 40px;
    max-width: 480px;
    width: 100%;
    text-align: center;
  }
  .icon { font-size: 56px; margin-bottom: 20px; }
  h1 { font-size: 22px; font-weight: 800; margin-bottom: 12px; color: #f87171; }
  p  { font-size: 14px; color: #9ca3af; line-height: 1.7; margin-bottom: 8px; }
  .btn {
    display: inline-block;
    margin-top: 28px;
    padding: 12px 28px;
    background: linear-gradient(135deg, #7c3aed, #6d28d9);
    color: white;
    text-decoration: none;
    border-radius: 12px;
    font-size: 14px;
    font-weight: 700;
  }
  .btn:hover { background: linear-gradient(135deg, #6d28d9, #5b21b6); }
</style>
</head>
<body>
<div class="card">
  <div class="icon">🔧</div>
  <h1>Server Sedang Maintenance</h1>
  <p>Kami sedang melakukan perbaikan sistem. Order baru sementara tidak dapat diproses.</p>
  <p>Jika Anda sudah memiliki order, Anda masih dapat melihat status, melakukan pembayaran, dan mengunduh hasil di halaman History.</p>
  <a href="/adk/history.php" class="btn">Cek Order Saya</a>
</div>
</body>
</html>
