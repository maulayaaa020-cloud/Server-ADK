<?php
date_default_timezone_set('Asia/Jakarta');
session_start();
require_once __DIR__ . '/includes/config.php';
require_once __DIR__ . '/includes/db.php';

$email = $_SESSION['email'] ?? $_SESSION['cek_email'] ?? null;
if (!$email) { header("Location: cek_pembelian.php"); exit; }

$id = (int)($_GET['id'] ?? 0);
if (!$id) { header("Location: history.php"); exit; }

try {
    $db   = getDB();
    $stmt = $db->prepare("SELECT * FROM orders WHERE id = :id AND phone = :phone");
    $stmt->execute([':id' => $id, ':phone' => $email]);
    $order = $stmt->fetch();
} catch (Exception $e) { $order = null; }

if (!$order) { header("Location: history.php"); exit; }

$isPaid   = $order['status'] === 'paid';
$hasPdf   = !empty($order['file_output_pdf']) && file_exists(__DIR__ . '/' . $order['file_output_pdf']);
$hasDocx  = !empty($order['file_output']);
$namaFile = preg_replace('/^\d+_/', '', $order['file_input']);
$namaPaket = $order['paket'] === 'paket1' ? 'Full Angka' : 'Romawi & Angka';

$chipClass = $isPaid ? 'chip-paid' : ($order['status'] === 'failed' ? 'chip-failed' : 'chip-pending');
$chipText  = $isPaid ? 'Lunas' : ($order['status'] === 'failed' ? 'Gagal' : 'Menunggu Bayar');
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Preview Hasil – ADK</title>
    <script>(function(){var t=localStorage.getItem('adkTheme')||'light';document.documentElement.setAttribute('data-theme',t);})();</script>
    <link rel="stylesheet" href="style.css?v=4">
    <style>
        .preview-wrap {
            max-width: 1100px;
            margin: 0 auto;
            padding: 32px 24px 60px;
        }

        .preview-topbar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 20px;
            flex-wrap: wrap;
            gap: 12px;
        }

        .preview-info h2 {
            font-size: 18px;
            font-weight: 800;
            color: white;
            margin: 0 0 4px;
        }

        .preview-info p {
            font-size: 12px;
            color: #a78bfa;
            margin: 0;
        }

        .preview-actions { display: flex; gap: 8px; flex-wrap: wrap; }

        .btn-action {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 9px 16px;
            border-radius: 10px;
            font-size: 12px;
            font-weight: 700;
            text-decoration: none;
            cursor: pointer;
            border: none;
            transition: 0.2s;
            letter-spacing: 0.3px;
        }

        .btn-back {
            background: rgba(255,255,255,0.07);
            color: #d1d5db;
            border: 1px solid rgba(255,255,255,0.12);
        }
        .btn-back:hover { background: rgba(255,255,255,0.13); color: white; }

        .btn-dl-docx {
            background: rgba(124,58,237,0.18);
            color: #a78bfa;
            border: 1px solid rgba(124,58,237,0.4);
        }
        .btn-dl-docx:hover { background: rgba(124,58,237,0.32); }

        .btn-dl-pdf {
            background: #7c3aed;
            color: white;
        }
        .btn-dl-pdf:hover {
            background: #a855f7;
            box-shadow: 0 4px 16px rgba(124,58,237,0.45);
        }

        .btn-locked {
            background: rgba(255,255,255,0.03);
            color: #4b5563;
            border: 1px solid rgba(255,255,255,0.06);
            cursor: not-allowed;
        }

        .status-chip {
            display: inline-block;
            font-size: 10px;
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 20px;
            margin-left: 6px;
            vertical-align: middle;
        }
        .chip-paid    { background: rgba(52,211,153,0.15); color: #34d399; }
        .chip-pending { background: rgba(251,191,36,0.12); color: #fbbf24; }
        .chip-failed  { background: rgba(239,68,68,0.1);   color: #f87171; }

        .pdf-container {
            background: #1e1e2e;
            border: 1px solid rgba(124,58,237,0.2);
            border-radius: 14px;
            overflow: hidden;
        }

        .pdf-frame {
            width: 100%;
            height: 80vh;
            border: none;
            display: block;
        }

        .info-box {
            border-radius: 14px;
            padding: 56px 24px;
            text-align: center;
        }

        .info-box.warning {
            background: rgba(251,191,36,0.06);
            border: 1px solid rgba(251,191,36,0.2);
        }

        .info-box.neutral {
            background: rgba(255,255,255,0.03);
            border: 1px solid rgba(124,58,237,0.15);
        }

        .info-icon { font-size: 44px; margin-bottom: 14px; }
        .info-title { font-size: 18px; font-weight: 800; color: white; margin-bottom: 8px; }
        .info-sub { font-size: 13px; color: #9ca3af; line-height: 1.6; margin-bottom: 24px; }

        .btn-bayar-now {
            display: inline-block;
            padding: 12px 28px;
            background: #ef4444;
            color: white;
            border-radius: 12px;
            font-size: 14px;
            font-weight: 700;
            text-decoration: none;
            transition: 0.2s;
        }
        .btn-bayar-now:hover { background: #dc2626; }
    </style>
</head>
<body>
    <div class="navbar">
        <a href="index.html" class="logo">
            <img src="Favicon Adkivia.png" alt="ADK Logo">
            <span><span style="color:#1565C0">ADK</span><span style="color:#29B6F6">IVIA</span></span>
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
                </div>
            </div>
        </div>
        <div class="nav-right">
            <button class="theme-toggle" onclick="toggleTheme()" title="Mode Siang">☀️</button>
            <a href="cek_pembelian.php" class="btn-nav">Cek Pembelian</a>
        </div>
        <div class="mobile-jasa-drop" id="jasaDrop">
            <button class="mobile-jasa-link" onclick="toggleJasaDrop()">Jasa ADK <span class="mobile-jasa-arrow">▾</span></button>
            <div class="mobile-jasa-menu">
                <a href="jasa.html">📄 Penomoran Halaman</a>
            </div>
        </div>
        <button class="theme-toggle theme-toggle-mob" onclick="toggleTheme()" title="Mode Siang">☀️</button>
        <button class="hamburger" onclick="openMobileMenu()">
            <span></span><span></span><span></span>
        </button>
    </div>

    <div class="mobile-nav" id="mobileNav">
        <div class="mobile-nav-header">
            <a href="index.html" class="logo">
                <img src="Favicon Adkivia.png" alt="ADK Logo">
                <span><span style="color:#1565C0">ADK</span><span style="color:#29B6F6">IVIA</span></span>
            </a>
            <div style="display:flex;align-items:center;gap:8px;">
                <button class="theme-toggle" onclick="toggleTheme()" title="Mode Siang">☀️</button>
                <button class="mobile-nav-close" onclick="closeMobileMenu()">✕</button>
            </div>
        </div>
        <div class="mobile-nav-links">
            <a href="index.html">Home</a>
            <a href="tutorial.html">Tutorial</a>
            <a href="jasa.html">Jasa ADK</a>
        </div>
        <div class="mobile-nav-footer">
            <a href="cek_pembelian.php" class="btn-nav">Cek Pembelian</a>
        </div>
    </div>
    <script>
        function openMobileMenu()  { document.getElementById('mobileNav').classList.add('open'); document.body.style.overflow='hidden'; }
        function closeMobileMenu() { document.getElementById('mobileNav').classList.remove('open'); document.body.style.overflow=''; }
        function toggleJasaDrop()  { document.getElementById('jasaDrop').classList.toggle('open'); }
        document.addEventListener('click', function(e) { var d = document.getElementById('jasaDrop'); if (d && !d.contains(e.target)) d.classList.remove('open'); });
    </script>

    <div class="preview-wrap">

        <div class="preview-topbar">
            <div class="preview-info">
                <h2>
                    <?= htmlspecialchars($namaFile) ?>
                    <span class="status-chip <?= $chipClass ?>"><?= $chipText ?></span>
                </h2>
                <p>Penomoran Halaman &middot; <?= htmlspecialchars($namaPaket) ?></p>
            </div>

            <div class="preview-actions">
                <a href="history.php" class="btn-action btn-back">← Kembali</a>

                <?php if ($isPaid && $hasDocx): ?>
                <a href="<?= htmlspecialchars($order['file_output']) ?>"
                   download class="btn-action btn-dl-docx">⬇ Docx</a>
                <?php endif; ?>

                <?php if ($isPaid && $hasPdf): ?>
                <a href="<?= htmlspecialchars($order['file_output_pdf']) ?>"
                   download class="btn-action btn-dl-pdf">⬇ PDF</a>
                <?php elseif ($isPaid): ?>
                <span class="btn-action btn-locked">PDF belum siap</span>
                <?php else: ?>
                <span class="btn-action btn-locked">🔒 Bayar dulu</span>
                <?php endif; ?>
            </div>
        </div>

        <?php if ($isPaid && $hasPdf): ?>
        <!-- PDF Viewer -->
        <div class="pdf-container">
            <iframe class="pdf-frame"
                    src="<?= htmlspecialchars($order['file_output_pdf']) ?>#toolbar=1&navpanes=0"
                    title="Preview PDF Hasil">
            </iframe>
        </div>

        <?php elseif ($isPaid && !$hasPdf): ?>
        <!-- Sudah bayar tapi PDF tidak tersedia -->
        <div class="info-box neutral">
            <div class="info-icon">📄</div>
            <div class="info-title">File Docx Siap</div>
            <div class="info-sub">
                PDF tidak tersedia untuk order ini.<br>
                Gunakan tombol <strong>⬇ Docx</strong> di atas untuk mengunduh hasilnya.
            </div>
        </div>

        <?php else: ?>
        <!-- Belum bayar -->
        <div class="info-box warning">
            <div class="info-icon">🔒</div>
            <div class="info-title">Selesaikan Pembayaran</div>
            <div class="info-sub">
                File sudah diproses dan siap diakses.<br>
                Lakukan pembayaran untuk melihat preview dan mengunduh hasilnya.
            </div>
            <a href="history.php" class="btn-bayar-now">Bayar Sekarang →</a>
        </div>
        <?php endif; ?>

    </div>

    <script>
        document.addEventListener('contextmenu', e => e.preventDefault());
    </script>
    <script src="theme.js?v=1"></script>
</body>
</html>



