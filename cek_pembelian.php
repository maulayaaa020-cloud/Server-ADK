<?php
session_start();
require_once __DIR__ . '/includes/config.php';
require_once __DIR__ . '/includes/db.php';

// Jika sudah login atau tamu, langsung ke history (kecuali mode ganti nomor)
if (empty($_GET['change'])) {
    if (!empty($_SESSION['email'])) { header("Location: history.php"); exit; }
    if (!empty($_SESSION['guest_token']) || !empty($_COOKIE['adk_guest'])) { header("Location: history.php"); exit; }
}

$error = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $email = trim($_POST['email'] ?? '');
    if (!$email) {
        $error = 'Email tidak boleh kosong.';
    } else {
        // Timpa semua session email agar email baru langsung aktif
        $_SESSION['email']     = $email;
        $_SESSION['cek_email'] = $email;
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
    <meta name="robots" content="noindex, nofollow">
    <title>Cek Pembelian - ADK</title>
    <link rel="icon" type="image/png" href="favicon.png">
    <script>(function(){var t=localStorage.getItem('adkTheme')||'light';document.documentElement.setAttribute('data-theme',t);})();</script>
    <link rel="stylesheet" href="style.css?v=4">
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
            font-size: 24px;
            font-weight: 800;
            color: white;
            margin: 0 0 6px;
            text-align: center;
        }

        .login-sub {
            font-size: 15px;
            color: #a78bfa;
            text-align: center;
            margin-bottom: 28px;
        }

        .login-label {
            font-size: 15px;
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
            font-size: 14px;
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

        /* ── LIGHT MODE ── */
        html[data-theme="light"] .login-card { background: #ffffff; border-color: #e2e8f0; box-shadow: 0 4px 24px rgba(0,0,0,0.06); }
        html[data-theme="light"] .login-title { color: #1a2332; }
        html[data-theme="light"] .login-sub { color: #6d28d9; }
        html[data-theme="light"] .login-label { color: #64748b; }
        html[data-theme="light"] .login-input { background: #f8fafc; color: #1a2332; border-color: rgba(109,40,217,0.3); }
        html[data-theme="light"] .login-input::placeholder { color: #9ca3af; }
        html[data-theme="light"] .login-input:focus { background: rgba(109,40,217,0.05); border-color: #6d28d9; }
    </style>
</head>
<body>
    <div class="navbar">
        <a href="index.html" class="logo">
            <img src="Favicon Adkivia.png" alt="ADK Logo" style="width:52px;height:52px;object-fit:contain;flex-shrink:0;">
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

            <div class="dropdown">
                <a href="#" class="dropdown-toggle">Contact Us <span class="dropdown-arrow">▾</span></a>
                <div class="dropdown-menu">
                    <div class="dropdown-group-label">Hubungi Kami</div>
                    <a href="https://wa.me/6281228790091" target="_blank" rel="noopener"><div class="dropdown-item-icon" style="background:rgba(37,211,102,0.15);">💬</div><div class="dropdown-item-title">WA Admin</div></a>
                    <a href="https://wa.me/6287700277748" target="_blank" rel="noopener"><div class="dropdown-item-icon" style="background:rgba(37,211,102,0.15);">🏪</div><div class="dropdown-item-title">WA Toko (Cepiring)</div></a>
                    <a href="https://wa.me/62895341996647" target="_blank" rel="noopener"><div class="dropdown-item-icon" style="background:rgba(37,211,102,0.15);">🏪</div><div class="dropdown-item-title">WA Toko (Rowosari)</div></a>
                    <div class="dropdown-divider"></div>
                    <div class="dropdown-group-label">Sosial Media</div>
                    <a href="https://www.instagram.com/adkivia" target="_blank" rel="noopener"><div class="dropdown-item-icon" style="background:rgba(225,48,108,0.15);">📸</div><div class="dropdown-item-title">Instagram</div></a>
                    <a href="https://www.tiktok.com/@adkivia" target="_blank" rel="noopener"><div class="dropdown-item-icon" style="background:rgba(0,0,0,0.12);">🎵</div><div class="dropdown-item-title">TikTok</div></a>
                </div>
            </div>
        </div>
        <div class="nav-right">
            <button class="theme-toggle" onclick="toggleTheme()" title="Mode Siang">☀️</button>
            <a href="cek_pembelian.php" class="btn-nav active">Cek Pembelian</a>
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
                <img src="Favicon Adkivia.png" alt="ADK Logo" style="width:52px;height:52px;object-fit:contain;flex-shrink:0;">
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
            <div class="mobile-nav-section-title">Contact Us</div>
            <a href="https://wa.me/6281228790091" target="_blank" rel="noopener" class="mobile-nav-contact">💬 WA Admin</a>
            <a href="https://wa.me/6287700277748" target="_blank" rel="noopener" class="mobile-nav-contact">🏪 WA Toko (Cepiring)</a>
            <a href="https://wa.me/62895341996647" target="_blank" rel="noopener" class="mobile-nav-contact">🏪 WA Toko (Rowosari)</a>
            <a href="https://www.instagram.com/adkivia" target="_blank" rel="noopener" class="mobile-nav-contact">📸 Instagram</a>
            <a href="https://www.tiktok.com/@adkivia" target="_blank" rel="noopener" class="mobile-nav-contact">🎵 TikTok</a>
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
                <div style="font-size:15px;font-weight:700;color:#fbbf24">Anda memiliki order sebagai Tamu</div>
                <div style="font-size:13px;color:#d97706;margin-top:2px">Klik untuk melihat riwayat order tamu →</div>
            </div>
        </a>
    </div>
    <?php endif; ?>

    <div class="login-wrapper">
        <div class="login-card">
            <div class="login-title">Cek Pembelian</div>
            <div class="login-sub">Masukkan email yang digunakan saat order</div>

            <form method="POST">
                <label class="login-label">Email</label>
                <input type="email" name="email" class="login-input"
                       placeholder="Contoh: email@kamu.com" autofocus>
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

        var _currentEmail = <?= json_encode($_SESSION['email'] ?? '') ?>;

        document.querySelector('form').addEventListener('submit', function (e) {
            e.preventDefault();
            var email = document.querySelector('input[name="email"]').value.trim();
            if (!email) return;

            // Email sama dengan yang sudah login → skip OTP langsung masuk
            if (_currentEmail && email === _currentEmail) {
                HTMLFormElement.prototype.submit.call(document.querySelector('form'));
                return;
            }

            requireOTP(email, function () {
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
                <a href="tentang-kami.html">Tentang Kami</a>
                <a href="kebijakan-privasi.html">Kebijakan Privasi</a>
            </div>
            <div class="footer-col">
                <h4>Kontak</h4>
                <a href="https://wa.me/6281228790091" target="_blank" rel="noopener">WA Admin</a>
                <a href="https://wa.me/62895341996647" target="_blank" rel="noopener">WA Toko</a>
            </div>
            <div class="footer-col">
                <h4>Ikuti Kami</h4>
                <a href="https://www.tiktok.com/@adkivia" target="_blank" rel="noopener">TikTok</a>
                <a href="https://www.instagram.com/adkivia" target="_blank" rel="noopener">Instagram</a>
            </div>
        </div>
        <hr class="footer-divider">
        <p class="footer-bottom">© 2025 ADK PHOTOCOPY · All Rights Reserved.</p>
    </footer>
    <script src="nocopy.js"></script>
    <script src="theme.js?v=1"></script>
</body>
</html>



