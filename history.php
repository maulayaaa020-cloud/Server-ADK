<?php
date_default_timezone_set('Asia/Jakarta');
session_start();
require_once __DIR__ . '/includes/config.php';
require_once __DIR__ . '/includes/db.php';

$isAdmin = !empty($_SESSION['adk_admin']);

$phone = $_SESSION['phone'] ?? $_SESSION['cek_phone'] ?? null;
if (empty($phone)) $phone = null;
// Jika session menandai mode tamu, abaikan nomor telepon lama
if (!empty($_SESSION['is_guest'])) $phone = null;

// Pulihkan guest_token dari cookie jika session habis (browser ditutup/dibuka lagi)
if (empty($_SESSION['guest_token']) && !empty($_COOKIE['adk_guest'])) {
    $_SESSION['guest_token'] = $_COOKIE['adk_guest'];
}
$guestToken = $_SESSION['guest_token'] ?? null;
$isGuest    = empty($phone) && !empty($guestToken);

if (!$isAdmin && !$phone && !$guestToken) { header("Location: cek_pembelian.php"); exit; }

try {
    $db = getDB();
    if ($isAdmin) {
        $stmt = $db->query("SELECT * FROM orders ORDER BY created_at DESC");
    } elseif ($isGuest) {
        $stmt = $db->prepare("SELECT * FROM orders WHERE guest_token = :gt ORDER BY created_at DESC");
        $stmt->execute([':gt' => $guestToken]);
    } else {
        $stmt = $db->prepare("SELECT * FROM orders WHERE phone = :phone ORDER BY created_at DESC");
        $stmt->execute([':phone' => $phone]);
    }
    $orders = $stmt->fetchAll();
} catch (Exception $e) { $orders = []; }

function namaPaket($p) {
    return $p === 'paket1' ? 'FULL ANGKA' : 'ROMAWI DAN ANGKA';
}

function keteranganList($row) {
    if ($row['paket'] === 'paket3') {
        $items = [['ROMAWI','Bawah Tengah'],['BAB','Bawah Tengah'],['ISI BAB','Kanan Atas']];
    } elseif ($row['paket'] === 'paket2') {
        $pos = $row['posisi'] ?: '-';
        $items = [['ROMAWI',$pos],['BAB',$pos],['ISI BAB',$pos]];
    } else {
        $items = [['POSISI', $row['posisi'] ?: '-']];
    }
    $items[] = ['Font',  ($row['font'] ?: '-') . ' ' . ($row['size'] ?: '')];
    $items[] = ['Cover', $row['hidden_cover'] === 'Tidak' ? 'Tampil' : 'Disembunyikan'];
    return $items;
}

function maskPhone($p) {
    $len = strlen($p);
    if ($len <= 8) return $p;
    return substr($p, 0, 4) . str_repeat('*', $len - 8) . substr($p, -4);
}

function tglIndo($dt) {
    $bulan = ['','Januari','Februari','Maret','April','Mei','Juni',
               'Juli','Agustus','September','Oktober','November','Desember'];
    $ts = strtotime($dt);
    return date('j',$ts).' '.$bulan[(int)date('n',$ts)].' '.date('Y',$ts).' pukul '.date('H.i',$ts);
}

// Siapkan data pending untuk JS timer
$pendingJS = [];
foreach ($orders as $o) {
    if ($o['status'] === 'pending') {
        $expiry = strtotime($o['created_at']) + 1800;
        $pendingJS[] = ['id' => (int)$o['id'], 'expiry' => $expiry];
    }
}
$hasPending = !empty($pendingJS);

// Order yang sudah pernah dilaporkan
$reportedOrderIds = [];
if (!empty($orders)) {
    try {
        $oids = array_column($orders, 'order_id');
        $placeholders = implode(',', array_fill(0, count($oids), '?'));
        $stmt2 = $db->prepare("SELECT DISTINCT order_id FROM bug_reports WHERE order_id IN ($placeholders)");
        $stmt2->execute($oids);
        $reportedOrderIds = array_column($stmt2->fetchAll(), 'order_id');
    } catch (Exception $e) {}
}
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Riwayat Orderan - ADK</title>
    <link rel="icon" type="image/png" href="LOGO ADK.png">
    <link rel="stylesheet" href="style.css">
    <style>
        .history-page {
            max-width: 960px;
            margin: 0 auto;
            padding: 40px 24px 120px;
        }

        .history-header {
            text-align: center;
            margin-bottom: 36px;
        }

        .history-heading {
            font-size: 30px;
            font-weight: 800;
            color: white;
            margin: 0 0 16px;
        }

        .history-phone-chip {
            display: inline-flex;
            align-items: center;
            gap: 10px;
            background: rgba(124,58,237,0.15);
            border: 1px solid rgba(124,58,237,0.35);
            border-radius: 40px;
            padding: 8px 18px;
        }

        .phone-masked {
            font-size: 14px;
            font-weight: 700;
            color: white;
            letter-spacing: 1px;
        }

        .phone-divider { width: 1px; height: 14px; background: rgba(255,255,255,0.2); }

        .history-phone-chip a {
            font-size: 12px;
            color: #a78bfa;
            text-decoration: none;
            font-weight: 600;
        }

        .history-phone-chip a:hover { color: white; }

        .order-list { display: flex; flex-direction: column; gap: 16px; }

        /* ===== CARD ===== */
        .order-card {
            background: rgba(255,255,255,0.04);
            border: 1px solid rgba(124,58,237,0.25);
            border-radius: 20px;
            padding: 22px 24px;
            display: flex;
            gap: 20px;
            align-items: flex-start;
            flex-wrap: wrap;
            transition: box-shadow 0.2s;
        }

        .order-card:hover {
            box-shadow: 0 0 0 1px rgba(124,58,237,0.4), 0 8px 32px rgba(124,58,237,0.1);
        }

        .order-card.paid   { border-color: rgba(52,211,153,0.3); }
        .order-card.failed { border-color: rgba(239,68,68,0.25); }

        .order-card.paid:hover {
            box-shadow: 0 0 0 1px rgba(52,211,153,0.35), 0 8px 32px rgba(52,211,153,0.08);
        }

        /* KIRI */
        .card-left { flex: 0 0 200px; }

        .card-file {
            font-size: 15px;
            font-weight: 800;
            color: white;
            margin-bottom: 3px;
            max-width: 190px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .card-service { font-size: 11px; color: #a78bfa; margin-bottom: 10px; }

        .card-status-badge {
            display: inline-block;
            font-size: 11px;
            font-weight: 700;
            padding: 3px 10px;
            border-radius: 20px;
            margin-bottom: 14px;
        }

        .badge-pending { color: #fbbf24; background: rgba(251,191,36,0.12); }
        .badge-sukses  { color: #34d399; background: rgba(52,211,153,0.12); }
        .badge-failed  { color: #f87171; background: rgba(239,68,68,0.1); }

        .card-timer {
            font-size: 11px;
            color: #fbbf24;
            font-weight: 600;
            margin-bottom: 12px;
            letter-spacing: 0.5px;
        }

        .btn-preview {
            display: block;
            width: 100%;
            box-sizing: border-box;
            padding: 8px 14px;
            background: rgba(255,255,255,0.09);
            color: white;
            border: 1px solid rgba(255,255,255,0.18);
            border-radius: 9px;
            font-size: 11px;
            font-weight: 700;
            letter-spacing: 0.8px;
            cursor: pointer;
            text-decoration: none;
            transition: 0.2s;
            text-align: center;
        }

        .btn-preview:hover {
            background: rgba(124,58,237,0.25);
            border-color: #a78bfa;
        }

        /* TENGAH */
        .card-mid { flex: 1; min-width: 180px; }

        .ket-label {
            font-size: 10px;
            font-weight: 700;
            color: #6b7280;
            text-transform: uppercase;
            letter-spacing: 0.6px;
            margin-bottom: 5px;
        }

        .ket-paket { font-size: 13px; font-weight: 700; color: white; margin-bottom: 8px; }

        .ket-list {
            list-style: none;
            margin: 0; padding: 0;
            display: flex; flex-direction: column; gap: 4px;
        }

        .ket-list li { font-size: 12px; color: #d1d5db; display: flex; gap: 5px; }
        .ket-list li::before { content: '•'; color: #7c3aed; flex-shrink: 0; }
        .ket-key { color: #e5e7eb; font-weight: 600; min-width: 66px; }

        /* KANAN */
        .card-right {
            flex: 0 0 190px;
            display: flex;
            flex-direction: column;
            align-items: flex-end;
            gap: 10px;
        }

        .card-date { font-size: 11px; color: #6b7280; text-align: right; }

        .sk-link {
            font-size: 10px;
            font-weight: 600;
            color: #f87171;
            text-decoration: none;
            text-align: right;
            display: block;
            letter-spacing: 0.3px;
            opacity: 0.85;
            transition: 0.15s;
        }
        .sk-link:hover { opacity: 1; text-decoration: underline; }

        .btn-bayar {
            display: block;
            width: 100%;
            padding: 10px;
            background: #ef4444;
            color: white;
            border: none;
            border-radius: 10px;
            font-size: 13px;
            font-weight: 700;
            cursor: pointer;
            transition: 0.2s;
            text-align: center;
        }

        .btn-bayar:hover {
            background: #dc2626;
            box-shadow: 0 4px 14px rgba(239,68,68,0.4);
        }

        .download-row {
            display: flex;
            align-items: center;
            gap: 6px;
            width: 100%;
        }

        .download-row span { font-size: 11px; color: #6b7280; flex-shrink: 0; }

        .btn-dl {
            flex: 1;
            padding: 7px 4px;
            border-radius: 8px;
            font-size: 11px;
            font-weight: 600;
            border: 1px solid;
            text-align: center;
            text-decoration: none;
            display: block;
            transition: 0.2s;
        }

        .btn-dl.locked {
            background: rgba(255,255,255,0.04);
            color: #374151;
            border-color: rgba(255,255,255,0.06);
            cursor: not-allowed;
        }

        .btn-dl.active {
            background: rgba(124,58,237,0.18);
            color: #a78bfa;
            border-color: rgba(124,58,237,0.4);
            cursor: pointer;
        }

        .btn-dl.active:hover { background: rgba(124,58,237,0.32); }

        /* ===== FLOATING BAR ===== */
        .float-bar {
            position: fixed;
            bottom: 0;
            left: 0;
            right: 0;
            z-index: 200;
            background: linear-gradient(90deg, rgba(239,68,68,0.92), rgba(220,38,38,0.92));
            backdrop-filter: blur(12px);
            border-top: 1px solid rgba(255,255,255,0.1);
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 20px;
            padding: 14px 24px;
        }

        .float-bar-text {
            font-size: 13px;
            font-weight: 700;
            color: white;
            letter-spacing: 0.3px;
        }

        .float-bar-timer {
            font-size: 20px;
            font-weight: 800;
            color: white;
            letter-spacing: 3px;
            font-variant-numeric: tabular-nums;
            background: rgba(0,0,0,0.2);
            padding: 4px 14px;
            border-radius: 8px;
        }

        .float-bar-btn {
            padding: 8px 20px;
            background: white;
            color: #ef4444;
            border: none;
            border-radius: 8px;
            font-size: 13px;
            font-weight: 800;
            cursor: pointer;
            transition: 0.2s;
        }

        .float-bar-btn:hover { background: #fff0f0; }

        /* ===== HANGING WARNING ===== */
        .card-wrap { display: flex; flex-direction: column; }

        .payment-warning {
            background: rgba(239,68,68,0.08);
            border: 1px solid rgba(239,68,68,0.28);
            border-top: none;
            border-radius: 0 0 16px 16px;
            padding: 9px 24px;
            text-align: center;
            color: #f87171;
            font-size: 12px;
            font-weight: 600;
            letter-spacing: 0.3px;
            margin-top: -4px;
        }

        .warning-timer {
            font-weight: 800;
            letter-spacing: 2px;
            margin-left: 8px;
            font-variant-numeric: tabular-nums;
        }

        /* Hilangkan radius bawah card jika ada warning */
        .card-wrap .order-card {
            border-bottom-left-radius: 4px;
            border-bottom-right-radius: 4px;
        }

        .card-wrap .payment-warning ~ * { display: none; }

        .empty-state { text-align: center; padding: 60px 20px; color: #6b7280; font-size: 14px; }

        /* ===== GUEST MODE ===== */
        .guest-section {
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 10px;
        }

        .guest-hint {
            font-size: 13px;
            font-weight: 600;
            color: #fbbf24;
            letter-spacing: 0.2px;
        }

        .btn-guest-login {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            background: rgba(124,58,237,0.15);
            border: 1px solid rgba(124,58,237,0.4);
            border-radius: 40px;
            padding: 9px 22px;
            color: #a78bfa;
            font-size: 14px;
            font-weight: 700;
            cursor: pointer;
            transition: 0.2s;
            letter-spacing: 0.3px;
        }

        .btn-guest-login:hover {
            background: rgba(124,58,237,0.28);
            color: white;
            border-color: #a78bfa;
        }

        .guest-form-wrap {
            width: 100%;
            max-width: 360px;
            margin-top: 4px;
        }

        .guest-form-inner {
            display: flex;
            gap: 8px;
            align-items: stretch;
        }

        .guest-phone-input {
            flex: 1;
            padding: 10px 14px;
            background: rgba(255,255,255,0.07);
            border: 1.5px solid rgba(124,58,237,0.4);
            border-radius: 10px;
            color: white;
            font-size: 14px;
            outline: none;
            transition: 0.2s;
        }

        .guest-phone-input::placeholder { color: #4b5563; }
        .guest-phone-input:focus {
            border-color: #7c3aed;
            background: rgba(124,58,237,0.12);
        }

        .btn-guest-submit {
            padding: 10px 16px;
            background: #7c3aed;
            border: none;
            border-radius: 10px;
            color: white;
            font-size: 13px;
            font-weight: 700;
            cursor: pointer;
            transition: 0.2s;
            white-space: nowrap;
        }

        .btn-guest-submit:hover { background: #a855f7; }
        .btn-guest-submit:disabled { opacity: 0.6; cursor: not-allowed; }

        .guest-login-error {
            font-size: 12px;
            color: #f87171;
            margin-top: 6px;
            min-height: 16px;
            text-align: center;
        }

        .guest-badge {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            background: rgba(251,191,36,0.1);
            border: 1px solid rgba(251,191,36,0.3);
            border-radius: 40px;
            padding: 5px 14px;
            font-size: 12px;
            font-weight: 700;
            color: #fbbf24;
        }

        /* ===== PREVIEW MODAL ===== */
        .pv-overlay {
            position: fixed; inset: 0; z-index: 2000;
            background: rgba(5,3,20,0.88);
            backdrop-filter: blur(10px);
            display: none;
            align-items: center;
            justify-content: center;
            padding: 16px;
        }
        .pv-overlay.open { display: flex; }

        .pv-box {
            background: #12102a;
            border: 1px solid rgba(124,58,237,0.35);
            border-radius: 18px;
            width: 100%; max-width: 880px;
            height: 90vh;
            display: flex; flex-direction: column;
            overflow: hidden;
            box-shadow: 0 28px 72px rgba(0,0,0,0.65);
            position: relative;
        }

        .pv-wm-overlay {
            position: absolute;
            top: 50px; bottom: 50px;
            left: 0; right: 0;
            pointer-events: none;
            z-index: 10;
        }

        .pv-header {
            display: flex; align-items: center;
            justify-content: space-between;
            padding: 14px 20px;
            border-bottom: 1px solid rgba(124,58,237,0.18);
            flex-shrink: 0; gap: 12px;
        }

        .pv-title {
            font-size: 13px; font-weight: 700;
            color: white; margin: 0;
            overflow: hidden; text-overflow: ellipsis;
            white-space: nowrap; flex: 1;
        }

        .pv-wm-tag {
            font-size: 10px; font-weight: 600;
            color: #fbbf24;
            background: rgba(251,191,36,0.1);
            border: 1px solid rgba(251,191,36,0.25);
            border-radius: 20px; padding: 2px 10px;
            white-space: nowrap; flex-shrink: 0;
        }

        .pv-close {
            background: none; border: none;
            color: #6b7280; font-size: 22px;
            cursor: pointer; padding: 2px 6px;
            border-radius: 6px; line-height: 1;
            transition: 0.15s; flex-shrink: 0;
        }
        .pv-close:hover { color: white; background: rgba(255,255,255,0.1); }

        .pv-loading {
            flex: 1; display: flex;
            flex-direction: column;
            align-items: center; justify-content: center;
            gap: 14px; color: #a78bfa; font-size: 13px;
        }

        .pv-spinner {
            width: 36px; height: 36px;
            border: 3px solid rgba(124,58,237,0.25);
            border-top-color: #7c3aed;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
        }
        @keyframes spin { to { transform: rotate(360deg); } }

        .pv-iframe-wrapper {
            flex: 1;
            position: relative;
            overflow: hidden;
            min-height: 0;
            display: none;
        }

        .pv-frame {
            position: absolute;
            top: 0; left: 0;
            width: 100%;
            /* 60px lebih tinggi dari wrapper agar toolbar bawah Office Online terpotong */
            height: calc(100% + 60px);
            border: none;
            background: white;
        }

        /* Mobile: scale-down iframe agar lebar halaman A4 muat di layar sempit.
           scale(0.45) → iframe harus render 222% lebar agar visual = 100% container.
           Tinggi: (100%+60px)/0.45 → setelah scale, visual = wrapper_height + 60px (toolbar ter-clip). */
        @media (max-width: 767px) {
            .pv-frame {
                transform-origin: top left;
                transform: scale(0.45);
                width: 222%;
                height: calc(222% + 133px);
            }
        }

        .pv-footer {
            display: flex; align-items: center;
            justify-content: space-between;
            padding: 10px 20px;
            border-top: 1px solid rgba(124,58,237,0.15);
            flex-shrink: 0; flex-wrap: wrap; gap: 8px;
        }

        .pv-footer-note {
            font-size: 11px; color: #6b7280; line-height: 1.4;
        }

        .btn-bayar-pv {
            padding: 8px 20px;
            background: #ef4444; color: white;
            border: none; border-radius: 8px;
            font-size: 12px; font-weight: 700;
            cursor: pointer; transition: 0.2s;
        }
        .btn-bayar-pv:hover { background: #dc2626; }

        @media (max-width: 640px) {
            .card-right { flex: 1 1 100%; align-items: stretch; }
            .card-date  { text-align: left; }
            .card-left  { flex: 1 1 100%; }
        }

        /* ===== BUG REPORT ===== */
        .btn-bug {
            display: inline-block;
            margin-top: 6px;
            padding: 6px 12px;
            background: rgba(239,68,68,0.08);
            color: #f87171;
            border: 1px solid rgba(239,68,68,0.25);
            border-radius: 8px;
            font-size: 10px;
            font-weight: 700;
            cursor: pointer;
            letter-spacing: 0.4px;
            transition: 0.2s;
            width: 100%;
            text-align: center;
        }
        .btn-bug:hover { background: rgba(239,68,68,0.16); border-color: rgba(239,68,68,0.5); }

        .btn-bug-done {
            background: rgba(52,211,153,0.06);
            color: #34d399;
            border-color: rgba(52,211,153,0.2);
            cursor: default;
        }

        .bug-overlay {
            position: fixed; inset: 0; z-index: 3000;
            background: rgba(5,3,20,0.85);
            backdrop-filter: blur(10px);
            display: none; align-items: center; justify-content: center;
            padding: 16px;
        }
        .bug-overlay.open { display: flex; }

        .bug-box {
            background: #12102a;
            border: 1px solid rgba(239,68,68,0.35);
            border-radius: 18px;
            width: 100%; max-width: 440px;
            padding: 28px 26px;
            box-shadow: 0 24px 60px rgba(0,0,0,0.6);
        }

        .bug-title {
            font-size: 17px; font-weight: 800;
            color: white; margin: 0 0 4px;
        }

        .bug-sub {
            font-size: 12px; color: #6b7280; margin: 0 0 20px;
        }

        .bug-options { display: flex; flex-direction: column; gap: 8px; margin-bottom: 16px; }

        .bug-option {
            display: flex; align-items: center; gap: 10px;
            padding: 10px 14px;
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 10px;
            cursor: pointer;
            transition: 0.15s;
            user-select: none;
        }
        .bug-option:hover { border-color: rgba(239,68,68,0.35); background: rgba(239,68,68,0.05); }
        .bug-option input[type=radio] { accent-color: #ef4444; width: 15px; height: 15px; flex-shrink: 0; }
        .bug-option label { font-size: 13px; color: #d1d5db; cursor: pointer; }

        .bug-textarea {
            width: 100%; box-sizing: border-box;
            background: rgba(255,255,255,0.05);
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 10px;
            padding: 10px 14px;
            color: #e2e8f0; font-size: 13px;
            resize: vertical; min-height: 80px;
            margin-bottom: 16px;
            display: none;
            font-family: inherit;
        }
        .bug-textarea:focus { outline: none; border-color: rgba(239,68,68,0.4); }
        .bug-textarea.show { display: block; }

        .bug-actions { display: flex; gap: 8px; }

        .bug-cancel {
            flex: 1; padding: 10px;
            background: rgba(255,255,255,0.05);
            border: 1px solid rgba(255,255,255,0.1);
            color: #9ca3af; border-radius: 10px;
            font-size: 13px; font-weight: 600;
            cursor: pointer; transition: 0.2s;
        }
        .bug-cancel:hover { background: rgba(255,255,255,0.1); color: white; }

        .bug-send {
            flex: 2; padding: 10px;
            background: #7c3aed;
            border: none; color: white;
            border-radius: 10px;
            font-size: 13px; font-weight: 700;
            cursor: pointer; transition: 0.2s;
            display: flex; align-items: center; justify-content: center; gap: 6px;
        }
        .bug-send:hover { background: #a855f7; box-shadow: 0 4px 14px rgba(124,58,237,0.4); }
        @media (max-width: 600px) {
            .footer-inner { display: grid !important; grid-template-columns: 1fr 1fr !important; gap: 24px 16px !important; }
            .footer-brand { grid-column: 1 / -1 !important; }
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
            <a href="cek_pembelian.php" class="btn-nav">Cek Pembelian</a>
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

    <div class="history-page">
        <div class="history-header">
            <h2 class="history-heading">Riwayat Orderan</h2>

            <?php if ($isAdmin): ?>
            <!-- Mode Admin -->
            <div class="history-phone-chip" style="gap:14px">
                <span>🔑</span>
                <span class="phone-masked" style="color:#a78bfa">Admin — Semua Order</span>
                <span class="phone-divider"></span>
                <a href="<?= BASE_PATH ?>/admin/dashboard.php">Dashboard</a>
            </div>
            <?php elseif ($isGuest): ?>
            <!-- Mode Tamu -->
            <div class="guest-section" id="guestSection">
                <div class="guest-badge">👤 Mode Tamu</div>
                <button class="btn-guest-login" onclick="toggleGuestLogin()">
                    📱 Login Nomor Telepon
                </button>
                <div class="guest-hint">Silahkan login untuk menyimpan history</div>
                <div class="guest-form-wrap" id="guestFormWrap" style="display:none">
                    <div class="guest-form-inner">
                        <input type="tel" id="guestPhoneInput" class="guest-phone-input"
                               placeholder="Contoh: 08123456789" autocomplete="tel">
                        <button class="btn-guest-submit" id="guestSubmitBtn" onclick="doGuestLogin()">
                            Verifikasi
                        </button>
                    </div>
                    <div id="guestLoginError" class="guest-login-error"></div>
                </div>
            </div>
            <?php else: ?>
            <!-- Mode Login -->
            <div class="history-phone-chip">
                <span>📱</span>
                <span class="phone-masked"><?= htmlspecialchars(maskPhone($phone)) ?></span>
                <span class="phone-divider"></span>
                <a href="cek_pembelian.php">Ganti Nomor</a>
            </div>
            <?php endif; ?>
        </div>

        <div class="order-list">
        <?php if (empty($orders)): ?>
            <div class="empty-state">Belum ada order untuk nomor ini.</div>
        <?php else: ?>
            <?php foreach ($orders as $o): ?>
            <?php
                $status    = $o['status'];
                $isPaid    = $status === 'paid';
                $isFailed  = $status === 'failed';
                $isPending = $status === 'pending';
                $rawName   = preg_replace('/^\d+_/', '', $o['file_input']);
                $fnBase    = pathinfo($rawName, PATHINFO_FILENAME);
                $fnExt     = pathinfo($rawName, PATHINFO_EXTENSION);
                $namaFile  = (mb_strlen($fnBase) > 12 ? mb_substr($fnBase, 0, 12) . '...' : $fnBase) . '.' . $fnExt;
                $expiry    = strtotime($o['created_at']) + 1800;
                $keterangan = keteranganList($o);
            ?>
            <div class="card-wrap">
            <div class="order-card <?= $isPaid ? 'paid' : ($isFailed ? 'failed' : '') ?>"
                 id="card-<?= $o['id'] ?>">

                <!-- KIRI -->
                <div class="card-left">
                    <div class="card-file" title="<?= htmlspecialchars($namaFile) ?>">
                        <?= htmlspecialchars($namaFile) ?>
                    </div>
                    <div class="card-service">Penomoran Halaman</div>
                    <div class="card-status-badge <?= $isPaid ? 'badge-sukses' : ($isFailed ? 'badge-failed' : 'badge-pending') ?>"
                         id="badge-<?= $o['id'] ?>">
                        <?= $isPaid ? 'Sukses' : ($isFailed ? 'Failed' : 'Waiting Payment') ?>
                    </div>
                    <?php if (!empty($o['file_output']) && !$isFailed && !$isPaid): ?>
                    <button class="btn-preview" onclick="openPreview(
                        '<?= addslashes(htmlspecialchars($namaFile)) ?>',
                        <?= $isPending ? 'true' : 'false' ?>,
                        <?= $o['id'] ?>,
                        '<?= addslashes($o['order_id']) ?>',
                        <?= (int)$o['harga'] ?>
                    )">PREVIEW HASIL</button>
                    <?php endif; ?>
                    <?php if (!empty($o['file_output'])): ?>
                    <?php if (!in_array($o['order_id'], $reportedOrderIds)): ?>
                    <button class="btn-bug" id="bugBtn-<?= $o['id'] ?>" onclick="openBugReport(
                        '<?= addslashes(htmlspecialchars($namaFile)) ?>',
                        '<?= addslashes($o['order_id']) ?>',
                        '<?= addslashes(htmlspecialchars(namaPaket($o['paket']))) ?>',
                        <?= $o['id'] ?>
                    )">Laporkan Bug</button>
                    <?php endif; ?>
                    <?php endif; ?>
                </div>

                <!-- TENGAH -->
                <div class="card-mid">
                    <div class="ket-label">Keterangan</div>
                    <div class="ket-paket"><?= namaPaket($o['paket']) ?></div>
                    <ul class="ket-list">
                        <?php foreach ($keterangan as $row): ?>
                        <li>
                            <span class="ket-key"><?= $row[0] ?></span>
                            <span>: <?= htmlspecialchars($row[1]) ?></span>
                        </li>
                        <?php endforeach; ?>
                    </ul>
                </div>

                <!-- KANAN -->
                <div class="card-right">
                    <div class="card-date"><?= tglIndo($o['created_at']) ?></div>

                    <?php if ($isPending): ?>
                    <a href="sk.html" target="_blank" class="sk-link">S&amp;K berlaku</a>
                    <button class="btn-bayar"
                            onclick="bayar(<?= $o['id'] ?>, '<?= htmlspecialchars($o['order_id']) ?>', <?= $o['harga'] ?>)">
                        Bayar — Rp <?= number_format($o['harga'],0,',','.') ?>
                    </button>
                    <?php endif; ?>

                    <div class="download-row">
                        <span>Download:</span>
                        <?php if ($isPaid && $o['file_output']): ?>
                        <a href="<?= BASE_PATH ?>/download.php?id=<?= $o['id'] ?>"
                           class="btn-dl active" download>Docx</a>
                        <?php else: ?>
                        <span class="btn-dl locked">Docx</span>
                        <?php endif; ?>
                    </div>
                </div>

            </div><!-- /order-card -->
            <?php if ($isPending): ?>
            <div class="payment-warning" id="warning-<?= $o['id'] ?>"
                 data-expiry="<?= $expiry ?>">
                !!! Lakukan Pembayaran untuk Download
                <span class="warning-timer" id="timer-<?= $o['id'] ?>">--:--:--</span>
            </div>
            <?php endif; ?>
            </div><!-- /card-wrap -->
            <?php endforeach; ?>
        <?php endif; ?>
        </div>
    </div>


    <!-- ===== PREVIEW MODAL ===== -->
    <div class="pv-overlay" id="pvOverlay" onclick="handleOverlayClick(event)">
        <div class="pv-box" id="pvBox">
            <div class="pv-header">
                <p class="pv-title" id="pvTitle">Preview</p>
                <span class="pv-wm-tag">🔍 WATERMARK PREVIEW</span>
                <button class="pv-close" onclick="closePreview()">✕</button>
            </div>

            <!-- Watermark overlay -->
            <div class="pv-wm-overlay">
                <svg width="100%" height="100%" xmlns="http://www.w3.org/2000/svg">
                    <defs>
                        <pattern id="wm-pattern" x="0" y="0" width="210" height="48"
                                 patternUnits="userSpaceOnUse" patternTransform="rotate(-35)">
                            <text x="8" y="28" font-family="Arial,sans-serif" font-size="15"
                                  font-weight="bold" fill="rgba(0,0,0,0.18)" letter-spacing="2">
                                PREVIEW HASIL ADK
                            </text>
                        </pattern>
                    </defs>
                    <rect width="100%" height="100%" fill="url(#wm-pattern)"/>
                </svg>
            </div>

            <div class="pv-loading" id="pvLoading">
                <div class="pv-spinner"></div>
                <span>Memuat preview...</span>
            </div>
            <div class="pv-iframe-wrapper" id="pvFrameWrapper">
                <iframe class="pv-frame" id="pvFrame" allowfullscreen></iframe>
            </div>

            <div class="pv-footer">
                <div class="pv-footer-note">
                    Preview adalah file docx hasil asli.<br>
                    Watermark akan hilang jika sudah dibayar.<br>
                    Jika ada kesalahan, silahkan Laporkan Bug!!
                </div>
                <button class="btn-bayar-pv" id="pvBayarBtn" style="display:none">
                    Bayar Sekarang →
                </button>
            </div>
        </div>
    </div>

    <!-- ===== BUG REPORT MODAL ===== -->
    <div class="bug-overlay" id="bugOverlay" onclick="handleBugOverlay(event)">
        <div class="bug-box" id="bugBox">
            <div class="bug-title">Laporkan Bug!</div>
            <div class="bug-sub" id="bugSubtitle"></div>

            <div class="bug-options">
                <div class="bug-option" onclick="selectBug(this)">
                    <input type="radio" name="bugType" id="bug1" value="Nomor halaman tidak sesuai">
                    <label for="bug1">Nomor halaman tidak sesuai</label>
                </div>
                <div class="bug-option" onclick="selectBug(this)">
                    <input type="radio" name="bugType" id="bug2" value="Nomor halaman tidak muncul">
                    <label for="bug2">Nomor halaman tidak muncul</label>
                </div>
                <div class="bug-option" onclick="selectBug(this, true)">
                    <input type="radio" name="bugType" id="bug3" value="Lainnya">
                    <label for="bug3">Lainnya</label>
                </div>
            </div>

            <textarea class="bug-textarea" id="bugDesc" placeholder="Jelaskan masalah yang kamu alami..."></textarea>

            <div class="bug-actions">
                <button class="bug-cancel" onclick="closeBugReport()">Batal</button>
                <button class="bug-send" id="bugSendBtn" onclick="kirimBug()">
                    <span class="bug-send-spinner" id="bugSpinner" style="display:none;width:14px;height:14px;border:2px solid rgba(255,255,255,0.35);border-top-color:white;border-radius:50%;animation:spin 0.7s linear infinite;flex-shrink:0"></span>
                    <span id="bugSendLabel">Kirim Laporan</span>
                </button>
            </div>
        </div>
    </div>

    <?php if ($isGuest): ?>
    <script src="otp.js?v=4"></script>
    <?php endif; ?>
    <script>
        const pendingOrders = <?= json_encode($pendingJS) ?>;
        const expired = {};

        function pad(n) { return String(n).padStart(2, '0'); }

        function fmtTime(sec) {
            if (sec < 0) sec = 0;
            return pad(Math.floor(sec / 3600)) + ':' +
                   pad(Math.floor((sec % 3600) / 60)) + ':' +
                   pad(sec % 60);
        }

        function expireOrder(id) {
            fetch('api/expire_order.php', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({id: id})
            }).then(r => r.json()).then(data => {
                if (data.ok) {
                    const badge   = document.getElementById('badge-'   + id);
                    const timer   = document.getElementById('timer-'   + id);
                    const card    = document.getElementById('card-'    + id);
                    const warning = document.getElementById('warning-' + id);
                    if (badge)   { badge.textContent = 'Failed'; badge.className = 'card-status-badge badge-failed'; }
                    if (timer)   timer.remove();
                    if (card)    { card.classList.remove('paid'); card.classList.add('failed'); card.style.borderBottomLeftRadius = ''; card.style.borderBottomRightRadius = ''; }
                    if (warning) warning.remove();
                    const btn = card ? card.querySelector('.btn-bayar') : null;
                    if (btn) btn.remove();
                }
            });
        }

        function updateTimers() {
            const now = Math.floor(Date.now() / 1000);
            let minLeft = Infinity;
            let anyActive = false;

            pendingOrders.forEach(o => {
                if (expired[o.id]) return;
                const left = o.expiry - now;
                const el   = document.getElementById('timer-' + o.id);

                if (left <= 0) {
                    if (el) el.textContent = '00:00:00';
                    if (!expired[o.id]) {
                        expired[o.id] = true;
                        expireOrder(o.id);
                    }
                } else {
                    if (el) el.textContent = fmtTime(left);
                    if (left < minLeft) minLeft = left;
                    anyActive = true;
                }
            });

        }

        function scrollToPending() {
            const pending = document.querySelector('.order-card:not(.paid):not(.failed)');
            if (pending) pending.scrollIntoView({behavior: 'smooth', block: 'center'});
        }

        const pollingIntervals = {};

        function markPaid(dbId) {
            clearInterval(pollingIntervals[dbId]);
            delete pollingIntervals[dbId];
            fetch('api/update_status.php', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({db_id: dbId})
            }).then(() => location.reload());
        }

        function startPolling(dbId, orderId) {
            if (pollingIntervals[dbId]) return;
            pollingIntervals[dbId] = setInterval(() => {
                fetch('api/check_payment.php', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({db_id: dbId, order_id: orderId})
                })
                .then(r => r.json())
                .then(d => {
                    if (d.status === 'paid') {
                        clearInterval(pollingIntervals[dbId]);
                        delete pollingIntervals[dbId];
                        location.reload();
                    } else if (d.status === 'failed') {
                        clearInterval(pollingIntervals[dbId]);
                        delete pollingIntervals[dbId];
                        location.reload();
                    }
                })
                .catch(() => {});
            }, 3000);
        }

        function bayar(dbId, orderId, harga) {
            fetch('api/get_snap_token.php', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({db_id: dbId, order_id: orderId, harga: harga})
            })
            .then(r => r.json())
            .then(data => {
                if (!data.url) { alert('Gagal membuat halaman pembayaran. Silakan coba lagi.'); return; }
                // Buka halaman pembayaran DOKU di tab baru, lalu mulai polling
                window.open(data.url, '_blank');
                startPolling(dbId, orderId);
            })
            .catch(() => alert('Koneksi ke server gagal.'));
        }

        if (pendingOrders.length > 0) {
            updateTimers();
            setInterval(updateTimers, 1000);
        }

        /* ── Preview Modal ── */
        function openPreview(fileName, isPending, dbId, ordId, harga) {
            var overlay  = document.getElementById('pvOverlay');
            var wrapper  = document.getElementById('pvFrameWrapper');
            var frame    = document.getElementById('pvFrame');
            var loading  = document.getElementById('pvLoading');
            var bayarBtn = document.getElementById('pvBayarBtn');

            document.getElementById('pvTitle').textContent = fileName;

            // Reset
            wrapper.style.display = 'none';
            frame.src             = '';
            loading.style.display = 'flex';
            loading.innerHTML     = '<div class="pv-spinner"></div><span>Memuat preview...</span>';

            if (isPending) {
                bayarBtn.style.display = 'inline-block';
                bayarBtn.onclick = function() { closePreview(); bayar(dbId, ordId, harga); };
            } else {
                bayarBtn.style.display = 'none';
            }

            overlay.classList.add('open');
            document.body.style.overflow = 'hidden';

            // Ambil viewer URL via AJAX — file path tidak pernah ada di HTML
            fetch('api/get_preview_url.php', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({db_id: dbId, vw: window.innerWidth})
            })
            .then(function(r) { return r.json(); })
            .then(function(data) {
                if (!data.ok) {
                    loading.innerHTML = '<span style="color:#f87171">Gagal memuat preview.</span>';
                    return;
                }
                frame.onload = function() {
                    loading.style.display = 'none';
                    wrapper.style.display = 'flex';
                };
                frame.src = data.viewer_url;
            })
            .catch(function() {
                loading.innerHTML = '<span style="color:#f87171">Gagal terhubung ke server.</span>';
            });
        }

        function closePreview() {
            var overlay = document.getElementById('pvOverlay');
            var wrapper = document.getElementById('pvFrameWrapper');
            var frame   = document.getElementById('pvFrame');
            overlay.classList.remove('open');
            wrapper.style.display = 'none';
            frame.src             = '';
            document.getElementById('pvLoading').style.display = 'flex';
            document.getElementById('pvLoading').innerHTML =
                '<div class="pv-spinner"></div><span>Memuat preview...</span>';
            document.body.style.overflow = '';
        }

        function handleOverlayClick(e) {
            if (e.target === document.getElementById('pvOverlay')) closePreview();
        }

        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape') closePreview();
        });

        function openMobileMenu()  { document.getElementById('mobileNav').classList.add('open'); document.body.style.overflow='hidden'; }
        function closeMobileMenu() { document.getElementById('mobileNav').classList.remove('open'); document.body.style.overflow=''; }

        /* ── Bug Report ── */
        let _bugFile = '', _bugOrder = '', _bugPaket = '', _bugCardId = null;

        function openBugReport(fileName, orderId, paket, cardId) {
            _bugFile  = fileName;
            _bugOrder = orderId;
            _bugPaket = paket;
            document.getElementById('bugSubtitle').textContent = fileName + ' · ' + paket;
            document.querySelectorAll('input[name=bugType]').forEach(r => r.checked = false);
            document.querySelectorAll('.bug-option').forEach(o => o.style.borderColor = '');
            document.getElementById('bugDesc').value = '';
            document.getElementById('bugDesc').classList.remove('show');
            document.getElementById('bugOverlay').classList.add('open');
            document.body.style.overflow = 'hidden';
        }

        function closeBugReport() {
            document.getElementById('bugOverlay').classList.remove('open');
            document.body.style.overflow = '';
        }

        function handleBugOverlay(e) {
            if (e.target === document.getElementById('bugOverlay')) closeBugReport();
        }

        function selectBug(el, showTextarea) {
            el.querySelector('input[type=radio]').checked = true;
            document.querySelectorAll('.bug-option').forEach(o => o.style.borderColor = '');
            el.style.borderColor = 'rgba(239,68,68,0.5)';
            const ta = document.getElementById('bugDesc');
            if (showTextarea) ta.classList.add('show');
            else { ta.classList.remove('show'); ta.value = ''; }
        }

        function kirimBug() {
            const selected = document.querySelector('input[name=bugType]:checked');
            if (!selected) { alert('Pilih jenis masalah terlebih dahulu.'); return; }

            const jenis = selected.value;
            const desc  = document.getElementById('bugDesc').value.trim();
            const btn   = document.getElementById('bugSendBtn');
            const spinner = document.getElementById('bugSpinner');
            const label   = document.getElementById('bugSendLabel');

            // Loading state
            btn.disabled      = true;
            spinner.style.display = 'inline-block';
            label.textContent = 'Mengirim...';

            fetch('api/report_bug.php', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                credentials: 'same-origin',
                body: JSON.stringify({
                    order_id:  _bugOrder,
                    file_name: _bugFile,
                    paket:     _bugPaket,
                    jenis:     jenis,
                    deskripsi: desc
                })
            })
            .then(r => r.json())
            .then(data => {
                if (data.ok) {
                    closeBugReport();
                    const bugBtn = document.getElementById('bugBtn-' + _bugCardId);
                    if (bugBtn) bugBtn.remove();
                } else {
                    throw new Error(data.error || 'Gagal');
                }
            })
            .catch(() => {
                spinner.style.display = 'none';
                label.textContent = 'Gagal, coba lagi';
                btn.style.background = '#dc2626';
                btn.disabled = false;
                setTimeout(() => {
                    btn.style.background = '';
                    label.textContent = 'Kirim Laporan';
                }, 2000);
            });
        }
        function toggleJasaDrop()  { document.getElementById('jasaDrop').classList.toggle('open'); }
        document.addEventListener('click', function(e) { var d = document.getElementById('jasaDrop'); if (d && !d.contains(e.target)) d.classList.remove('open'); });

        document.addEventListener('contextmenu', e => e.preventDefault());
        document.addEventListener('keydown', e => {
            if ((e.ctrlKey || e.metaKey) && ['c','a','u','s'].includes(e.key.toLowerCase()))
                e.preventDefault();
        });

        /* ── Guest Login ── */
        <?php if ($isGuest): ?>
        const _guestToken = <?= json_encode($guestToken) ?>;

        function toggleGuestLogin() {
            const wrap = document.getElementById('guestFormWrap');
            const isHidden = wrap.style.display === 'none';
            wrap.style.display = isHidden ? 'block' : 'none';
            if (isHidden) document.getElementById('guestPhoneInput').focus();
        }

        function setGuestError(msg) {
            document.getElementById('guestLoginError').textContent = msg || '';
        }

        function doGuestLogin() {
            const phone = document.getElementById('guestPhoneInput').value.trim();
            if (!phone) { setGuestError('Masukkan nomor telepon.'); return; }

            const btn = document.getElementById('guestSubmitBtn');
            btn.disabled    = true;
            btn.textContent = 'Memproses...';
            setGuestError('');

            requireOTP(phone, function () {
                // OTP terverifikasi, klaim order tamu ke nomor ini
                fetch('api/claim_guest.php', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ phone: phone, guest_token: _guestToken })
                })
                .then(function(r) { return r.json(); })
                .then(function(data) {
                    if (data.ok) {
                        location.reload();
                    } else {
                        setGuestError(data.error || 'Gagal login, coba lagi.');
                        btn.disabled    = false;
                        btn.textContent = 'Verifikasi';
                    }
                })
                .catch(function() {
                    setGuestError('Gagal terhubung ke server.');
                    btn.disabled    = false;
                    btn.textContent = 'Verifikasi';
                });
            });
        }

        document.addEventListener('keydown', function(e) {
            if (e.key === 'Enter' && document.getElementById('guestFormWrap').style.display !== 'none') {
                doGuestLogin();
            }
        });
        <?php endif; ?>
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





