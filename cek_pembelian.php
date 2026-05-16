<?php
session_start();
require_once __DIR__ . '/includes/config.php';
require_once __DIR__ . '/includes/db.php';

$error = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $phone = trim($_POST['phone'] ?? '');
    if (!$phone) {
        $error = 'Nomor telepon tidak boleh kosong.';
    } else {
        // Timpa semua session phone agar nomor baru langsung aktif
        $_SESSION['phone']     = $phone;
        $_SESSION['cek_phone'] = $phone;
        // Bersihkan sisa sesi tamu agar tidak tumpang tindih
        unset($_SESSION['is_guest'], $_SESSION['guest_token']);
        header("Location: history.php");
        exit;
    }
}
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Cek Pembelian - ADK</title>
    <link rel="icon" type="image/png" href="favicon.png">
    <link rel="stylesheet" href="style.css">
    <style>
        .login-wrapper {
            max-width: 420px;
            margin: 100px auto;
            padding: 0 20px;
        }

        .login-card {
            background: rgba(255,255,255,0.06);
            border: 1px solid rgba(124,58,237,0.3);
            border-radius: 20px;
            padding: 36px 32px;
        }

        .login-title {
            font-size: 20px;
            font-weight: 800;
            color: white;
            margin: 0 0 6px;
            text-align: center;
        }

        .login-sub {
            font-size: 13px;
            color: #a78bfa;
            text-align: center;
            margin-bottom: 28px;
        }

        .login-label {
            font-size: 12px;
            font-weight: 600;
            color: #9ca3af;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 8px;
            display: block;
        }

        .login-input {
            width: 100%;
            padding: 12px 16px;
            background: rgba(255,255,255,0.08);
            border: 1px solid rgba(124,58,237,0.35);
            border-radius: 10px;
            color: white;
            font-size: 15px;
            outline: none;
            box-sizing: border-box;
            transition: 0.2s;
        }

        .login-input:focus {
            border-color: #7c3aed;
            background: rgba(124,58,237,0.12);
        }

        .login-input::placeholder { color: rgba(255,255,255,0.3); }

        .login-error {
            color: #f87171;
            font-size: 12px;
            margin-top: 8px;
        }

        .btn-cek {
            display: block;
            width: 100%;
            padding: 13px;
            background: #7c3aed;
            color: white;
            border: none;
            border-radius: 12px;
            font-size: 15px;
            font-weight: 700;
            cursor: pointer;
            margin-top: 20px;
            letter-spacing: 1px;
            transition: 0.2s;
        }

        .btn-cek:hover {
            background: #a855f7;
            box-shadow: 0 6px 20px rgba(124,58,237,0.45);
        }
    </style>
</head>
<body>
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
                </div>
            </div>

            <div class="dropdown">
                <a href="#" class="dropdown-toggle">Contact Us <span class="dropdown-arrow">▾</span></a>
                <div class="dropdown-menu">
                    <a href="https://www.tiktok.com/@adk_rowosari" target="_blank" rel="noopener"><div class="dropdown-item-icon">🎵</div><div class="dropdown-item-title">TikTok</div></a>
                    <a href="https://www.instagram.com/adk_rowosari?igsh=MXV4OXdwbnQycGp5cg==" target="_blank" rel="noopener"><div class="dropdown-item-icon">📸</div><div class="dropdown-item-title">Instagram</div></a>
                    <a href="https://wa.me/6281228790091" target="_blank" rel="noopener"><div class="dropdown-item-icon">💬</div><div class="dropdown-item-title">WA Admin</div></a>
                    <a href="https://wa.me/62895341996647" target="_blank" rel="noopener"><div class="dropdown-item-icon">🏪</div><div class="dropdown-item-title">WA Toko</div></a>
                </div>
            </div>
        </div>
        <div class="nav-right">
            <a href="cek_pembelian.php" class="btn-nav active">Cek Pembelian</a>
        </div>
        <div class="mobile-jasa-drop" id="jasaDrop">
            <button class="mobile-jasa-link" onclick="toggleJasaDrop()">Jasa ADK <span class="mobile-jasa-arrow">▾</span></button>
            <div class="mobile-jasa-menu">
                <a href="jasa.html">📄 Penomoran Halaman</a>
            </div>
        </div>
        <button class="hamburger" onclick="openMobileMenu()">
            <span></span><span></span><span></span>
        </button>
    </div>

    <div class="mobile-nav" id="mobileNav">
        <div class="mobile-nav-header">
            <a href="index.html" class="logo">
                <img src="LOGO ADK.png" alt="ADK Logo" style="width:36px;height:26px;object-fit:contain;flex-shrink:0;">
                <span>ADK PHOTOCOPY</span>
            </a>
            <button class="mobile-nav-close" onclick="closeMobileMenu()">✕</button>
        </div>
        <div class="mobile-nav-links">
            <a href="index.html">Home</a>
            <a href="tutorial.html">Tutorial</a>
            <a href="jasa.html">Jasa ADK</a>
            <div class="mobile-nav-section-title">Contact Us</div>
            <a href="https://www.tiktok.com/@adk_rowosari" target="_blank" rel="noopener" class="mobile-nav-contact">🎵 TikTok</a>
            <a href="https://www.instagram.com/adk_rowosari?igsh=MXV4OXdwbnQycGp5cg==" target="_blank" rel="noopener" class="mobile-nav-contact">📸 Instagram</a>
            <a href="https://wa.me/6281228790091" target="_blank" rel="noopener" class="mobile-nav-contact">💬 WA Admin</a>
            <a href="https://wa.me/62895341996647" target="_blank" rel="noopener" class="mobile-nav-contact">🏪 WA Toko</a>
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

    <?php if (!empty($_COOKIE['adk_guest'])): ?>
    <div style="max-width:420px;margin:24px auto 0;padding:0 20px">
        <a href="history.php" style="
            display:flex;align-items:center;gap:10px;
            background:rgba(251,191,36,0.08);
            border:1px solid rgba(251,191,36,0.3);
            border-radius:14px;padding:12px 18px;
            text-decoration:none;transition:0.2s;
        " onmouseover="this.style.background='rgba(251,191,36,0.15)'"
           onmouseout="this.style.background='rgba(251,191,36,0.08)'">
            <span style="font-size:18px">👤</span>
            <div>
                <div style="font-size:13px;font-weight:700;color:#fbbf24">Anda memiliki order sebagai Tamu</div>
                <div style="font-size:11px;color:#d97706;margin-top:2px">Klik untuk melihat riwayat order tamu →</div>
            </div>
        </a>
    </div>
    <?php endif; ?>

    <div class="login-wrapper">
        <div class="login-card">
            <div class="login-title">Cek Pembelian</div>
            <div class="login-sub">Masukkan nomor telepon yang digunakan saat order</div>

            <form method="POST">
                <label class="login-label">Nomor Telepon</label>
                <input type="tel" name="phone" class="login-input"
                       placeholder="Contoh: 08123456789" autofocus>
                <?php if ($error): ?>
                <div class="login-error"><?= htmlspecialchars($error) ?></div>
                <?php endif; ?>
                <button type="submit" class="btn-cek">LIHAT RIWAYAT</button>
            </form>
        </div>
    </div>

    <script src="otp.js?v=4"></script>
    <script>
        document.addEventListener('contextmenu', e => e.preventDefault());

        document.querySelector('form').addEventListener('submit', function (e) {
            e.preventDefault();
            var phone = document.querySelector('input[name="phone"]').value.trim();
            if (!phone) return;

            requireOTP(phone, function () {
                // OTP verified — submit form secara programatik (bypass listener)
                HTMLFormElement.prototype.submit.call(document.querySelector('form'));
            });
        });
    </script>

    <footer class="site-footer">
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
            </div>
            <div class="footer-col">
                <h4>Bantuan</h4>
                <a href="tutorial.html">Tutorial</a>
                <a href="sk.html">Syarat &amp; Ketentuan</a>
            </div>
            <div class="footer-col">
                <h4>Kontak</h4>
                <a href="https://wa.me/6281228790091" target="_blank" rel="noopener">WA Admin</a>
                <a href="https://wa.me/62895341996647" target="_blank" rel="noopener">WA Toko</a>
            </div>
            <div class="footer-col">
                <h4>Ikuti Kami</h4>
                <a href="https://www.tiktok.com/@adk_rowosari" target="_blank" rel="noopener">TikTok</a>
                <a href="https://www.instagram.com/adk_rowosari?igsh=MXV4OXdwbnQycGp5cg==" target="_blank" rel="noopener">Instagram</a>
            </div>
        </div>
        <hr class="footer-divider">
        <p class="footer-bottom">© 2025 ADK PHOTOCOPY · All Rights Reserved.</p>
    </footer>
    <script src="nocopy.js"></script>
</body>
</html>



