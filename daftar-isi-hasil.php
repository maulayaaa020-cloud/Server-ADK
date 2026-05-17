<?php
date_default_timezone_set('Asia/Jakarta');
session_start();
require_once __DIR__ . '/includes/config.php';

if (!isset($_SESSION['di_file_output'])) {
    header("Location: daftar-isi.html");
    exit;
}

$fileRel  = $_SESSION['di_file_output'];
$fileFull = __DIR__ . '/' . $fileRel;
$fileUrl  = rtrim(APP_URL, '/') . '/' . $fileRel;
$namaFile = basename($fileRel);

if (!file_exists($fileFull)) {
    header("Location: daftar-isi.html");
    exit;
}
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hasil Daftar Isi – ADK PHOTOCOPY</title>
    <link rel="icon" type="image/png" href="favicon.png">
    <link rel="stylesheet" href="style.css">
    <style>
        .hasil-wrap {
            max-width: 560px;
            margin: 60px auto;
            padding: 0 20px;
            text-align: center;
        }
        .hasil-icon {
            font-size: 56px;
            margin-bottom: 16px;
        }
        .hasil-title {
            font-size: 22px;
            font-weight: 700;
            color: #fff;
            margin-bottom: 8px;
        }
        .hasil-sub {
            font-size: 14px;
            color: #94a3b8;
            margin-bottom: 32px;
        }
        .btn-download {
            display: inline-flex;
            align-items: center;
            gap: 10px;
            background: linear-gradient(135deg, #7c3aed, #6d28d9);
            color: #fff;
            font-weight: 700;
            font-size: 15px;
            padding: 14px 32px;
            border-radius: 10px;
            text-decoration: none;
            margin-bottom: 32px;
            transition: opacity .2s;
        }
        .btn-download:hover { opacity: .85; }
        .info-box {
            background: #12102a;
            border: 1px solid rgba(124,58,237,.3);
            border-radius: 12px;
            padding: 20px 24px;
            text-align: left;
            margin-bottom: 24px;
        }
        .info-box h4 {
            font-size: 13px;
            font-weight: 700;
            color: #a78bfa;
            margin: 0 0 12px;
            text-transform: uppercase;
            letter-spacing: .5px;
        }
        .info-box ol {
            margin: 0;
            padding-left: 18px;
            color: #cbd5e1;
            font-size: 13px;
            line-height: 1.8;
        }
        .info-box ol li b { color: #fff; }
        .btn-back {
            color: #a78bfa;
            font-size: 13px;
            text-decoration: none;
        }
        .btn-back:hover { text-decoration: underline; }
    </style>
</head>
<body>

    <!-- NAVBAR -->
    <div class="navbar">
        <a href="index.html" class="logo">
            <img src="LOGO ADK.png" alt="ADK Logo" style="width:36px;height:26px;object-fit:contain;flex-shrink:0;">
            <span>ADK PHOTOCOPY</span>
        </a>
        <div class="menu">
            <a href="index.html">Home</a>
            <a href="tutorial.html">Tutorial</a>
            <div class="dropdown">
                <a href="#" class="dropdown-toggle">Jasa ADK <span class="dropdown-arrow">▾</span></a>
                <div class="dropdown-menu">
                    <a href="jasa.html">
                        <div class="dropdown-item-icon">📄</div>
                        <div class="dropdown-item-title">Penomoran Halaman</div>
                    </a>
                    <a href="daftar-isi.html">
                        <div class="dropdown-item-icon">📋</div>
                        <div class="dropdown-item-title">Daftar Isi Otomatis</div>
                    </a>
                </div>
            </div>
        </div>
        <div class="nav-right">
            <a href="cek_pembelian.php" class="btn-nav">Cek Pembelian</a>
        </div>
    </div>

    <div class="hasil-wrap">

        <div class="hasil-icon">✅</div>
        <h1 class="hasil-title">Daftar Isi Berhasil Dibuat!</h1>
        <p class="hasil-sub">File sudah siap. Klik tombol di bawah untuk mengunduh.</p>

        <a href="download_daftar_isi.php" class="btn-download">
            ⬇ Download File
        </a>

        <div class="info-box">
            <h4>Cara Pakai File Hasil</h4>
            <ol>
                <li>Buka file di <b>Microsoft Word</b></li>
                <li>Muncul popup → klik <b>Yes</b> untuk mengisi nomor halaman otomatis</li>
                <li>Tekan <b>Ctrl+S</b> untuk menyimpan</li>
                <li>Selesai — popup tidak akan muncul lagi saat file dibuka berikutnya ✓</li>
            </ol>
        </div>

        <a href="daftar-isi.html" class="btn-back">← Upload file lain</a>

    </div>

    <footer class="site-footer" style="margin-top:60px;">
        <div class="footer-inner">
            <div class="footer-brand">
                <div class="footer-logo-wrap">
                    <img src="LOGO ADK.png" alt="ADK Logo" style="width:36px;height:26px;object-fit:contain;flex-shrink:0;">
                    <span>ADK PHOTOCOPY</span>
                </div>
                <p>© 2025 ADK PHOTOCOPY<br>All Rights Reserved.</p>
            </div>
            <div class="footer-col">
                <h4>Jasa</h4>
                <a href="jasa.html">Penomoran Halaman</a>
                <a href="daftar-isi.html">Daftar Isi Otomatis</a>
            </div>
            <div class="footer-col">
                <h4>Kontak</h4>
                <a href="https://wa.me/6287700277748" target="_blank" rel="noopener">WA Toko (Cepiring)</a>
                <a href="https://wa.me/62895341996647" target="_blank" rel="noopener">WA Toko (Rowosari)</a>
                <a href="https://wa.me/6281228790091" target="_blank" rel="noopener">WA Admin</a>
            </div>
        </div>
        <hr class="footer-divider">
        <p class="footer-bottom">© 2025 ADK PHOTOCOPY · All Rights Reserved.</p>
    </footer>

</body>
</html>
