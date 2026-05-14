<?php
date_default_timezone_set('Asia/Jakarta');
require_once __DIR__ . '/_guard.php';

// ── Maintenance mode toggle ───────────────────────────────────────────────────
$maintenanceFlag = __DIR__ . '/../config/maintenance.flag';
$isMaintenance   = file_exists($maintenanceFlag);

// Ambil flash message dari session (hasil redirect sebelumnya)
$maintenanceError   = $_SESSION['maint_error']   ?? '';
$maintenanceSuccess = $_SESSION['maint_success'] ?? '';
unset($_SESSION['maint_error'], $_SESSION['maint_success']);

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['toggle_maintenance'])) {
    $pwd = $_POST['maintenance_password'] ?? '';
    if (!password_verify($pwd, ADMIN_PASSWORD_HASH)) {
        $_SESSION['maint_error'] = 'Sandi salah.';
    } else {
        if ($isMaintenance) {
            @unlink($maintenanceFlag);
            $_SESSION['maint_success'] = 'Maintenance mode dinonaktifkan.';
        } else {
            $written = file_put_contents($maintenanceFlag, date('Y-m-d H:i:s'));
            if ($written !== false) {
                $_SESSION['maint_success'] = 'Maintenance mode diaktifkan.';
            } else {
                $_SESSION['maint_error'] = 'Gagal mengaktifkan. Periksa izin folder config/.';
            }
        }
    }
    // PRG: redirect supaya refresh tidak memicu resubmit form
    header('Location: dashboard.php');
    exit;
}

$db = getDB();

// ── Stats ──────────────────────────────────────────────────────────────────
$stats = [];
$stats['total']   = (int)$db->query("SELECT COUNT(*) FROM orders")->fetchColumn();
$stats['today']   = (int)$db->query("SELECT COUNT(*) FROM orders WHERE DATE(created_at) = CURDATE()")->fetchColumn();
$stats['paid']    = (int)$db->query("SELECT COUNT(*) FROM orders WHERE status = 'paid'")->fetchColumn();
$stats['pending'] = (int)$db->query("SELECT COUNT(*) FROM orders WHERE status = 'pending'")->fetchColumn();
$stats['failed']  = (int)$db->query("SELECT COUNT(*) FROM orders WHERE status = 'failed'")->fetchColumn();
$stats['revenue'] = (float)$db->query("SELECT COALESCE(SUM(harga),0) FROM orders WHERE status = 'paid'")->fetchColumn();

// ── Filters ────────────────────────────────────────────────────────────────
$filterStatus = $_GET['status'] ?? 'all';
$filterDate   = $_GET['date']   ?? 'all';

$allowed_status = ['all', 'paid', 'pending', 'failed'];
$allowed_date   = ['all', 'today', 'week'];
if (!in_array($filterStatus, $allowed_status)) $filterStatus = 'all';
if (!in_array($filterDate,   $allowed_date))   $filterDate   = 'all';

$where  = [];
$params = [];

if ($filterStatus !== 'all') {
    $where[]           = "status = :status";
    $params[':status'] = $filterStatus;
}
if ($filterDate === 'today') {
    $where[] = "DATE(created_at) = CURDATE()";
} elseif ($filterDate === 'week') {
    $where[] = "created_at >= DATE_SUB(NOW(), INTERVAL 7 DAY)";
}

$sql = "SELECT * FROM orders";
if ($where) $sql .= " WHERE " . implode(" AND ", $where);
$sql .= " ORDER BY created_at DESC";

$stmt = $db->prepare($sql);
$stmt->execute($params);
$orders = $stmt->fetchAll();

// ── Bug Reports ────────────────────────────────────────────────────────────
$bugs = $db->query("SELECT * FROM bug_reports ORDER BY created_at DESC")->fetchAll();

// ── Helpers ────────────────────────────────────────────────────────────────
function fmtRupiah(float $n): string {
    return 'Rp ' . number_format($n, 0, ',', '.');
}

function tglAdmin(string $dt): string {
    $bulan = ['','Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
    $ts = strtotime($dt);
    return date('d',$ts).' '.$bulan[(int)date('n',$ts)].' '.date('Y',$ts).' '.date('H:i',$ts);
}
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard Admin — ADK</title>
    <style>
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

        body {
            font-family: 'Segoe UI', sans-serif;
            background: linear-gradient(135deg, #0e0b28, #1a1043);
            min-height: 100vh;
            color: white;
            display: flex;
        }

        /* ===== SIDEBAR ===== */
        .sidebar {
            width: 220px;
            flex-shrink: 0;
            background: rgba(255,255,255,0.04);
            border-right: 1px solid rgba(124,58,237,0.2);
            display: flex;
            flex-direction: column;
            padding: 28px 16px;
            position: sticky;
            top: 0;
            height: 100vh;
            overflow-y: auto;
        }

        .sidebar-brand {
            font-size: 15px;
            font-weight: 800;
            color: white;
            letter-spacing: 0.4px;
            padding: 0 8px 24px;
            border-bottom: 1px solid rgba(124,58,237,0.15);
            margin-bottom: 20px;
        }

        .sidebar-brand span {
            display: block;
            font-size: 10px;
            font-weight: 500;
            color: #6b7280;
            margin-top: 3px;
        }

        .sidebar-nav {
            list-style: none;
            display: flex;
            flex-direction: column;
            gap: 4px;
            flex: 1;
        }

        .sidebar-nav a {
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 10px 12px;
            border-radius: 10px;
            text-decoration: none;
            color: #9ca3af;
            font-size: 14px;
            font-weight: 600;
            transition: background 0.15s, color 0.15s;
        }

        .sidebar-nav a:hover,
        .sidebar-nav a.active {
            background: rgba(124,58,237,0.18);
            color: white;
        }

        .nav-icon { font-size: 16px; }

        .btn-logout {
            display: block;
            width: 100%;
            padding: 10px 12px;
            background: rgba(239,68,68,0.1);
            border: 1px solid rgba(239,68,68,0.25);
            border-radius: 10px;
            color: #f87171;
            font-size: 13px;
            font-weight: 700;
            cursor: pointer;
            text-decoration: none;
            text-align: center;
            transition: background 0.15s;
            margin-top: 12px;
            font-family: inherit;
        }

        .btn-logout:hover { background: rgba(239,68,68,0.2); }

        /* ===== MAIN ===== */
        .main {
            flex: 1;
            min-width: 0;
            padding: 32px 28px;
            overflow-x: hidden;
        }

        .page-title {
            font-size: 22px;
            font-weight: 800;
            margin-bottom: 4px;
        }

        .page-sub {
            font-size: 13px;
            color: #6b7280;
            margin-bottom: 28px;
        }

        /* ===== STATS GRID ===== */
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(155px, 1fr));
            gap: 14px;
            margin-bottom: 36px;
        }

        .stat-card {
            background: rgba(255,255,255,0.05);
            border: 1px solid rgba(124,58,237,0.2);
            border-radius: 16px;
            padding: 18px 20px;
        }

        .stat-label {
            font-size: 11px;
            font-weight: 700;
            color: #6b7280;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 10px;
        }

        .stat-value {
            font-size: 28px;
            font-weight: 800;
            line-height: 1;
            color: white;
        }

        .stat-value.green  { color: #34d399; }
        .stat-value.yellow { color: #fbbf24; }
        .stat-value.red    { color: #f87171; }
        .stat-value.purple { color: #a78bfa; }
        .stat-value.sm     { font-size: 18px; }

        /* ===== SECTION HEADER ===== */
        .section-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 12px;
            margin-bottom: 16px;
        }

        .section-title {
            font-size: 16px;
            font-weight: 800;
        }

        /* ===== FILTER BAR ===== */
        .filter-bar {
            display: flex;
            align-items: center;
            gap: 8px;
            flex-wrap: wrap;
        }

        .filter-tabs {
            display: flex;
            gap: 4px;
            background: rgba(255,255,255,0.05);
            border-radius: 10px;
            padding: 4px;
        }

        .filter-tab {
            padding: 6px 14px;
            border-radius: 7px;
            font-size: 12px;
            font-weight: 700;
            cursor: pointer;
            text-decoration: none;
            color: #9ca3af;
            transition: background 0.15s, color 0.15s;
        }

        .filter-tab:hover { color: white; }
        .filter-tab.active { background: #7c3aed; color: white; }

        .filter-select {
            padding: 7px 12px;
            background: rgba(255,255,255,0.07);
            border: 1px solid rgba(124,58,237,0.3);
            border-radius: 9px;
            color: white;
            font-size: 12px;
            font-weight: 600;
            outline: none;
            cursor: pointer;
            font-family: inherit;
        }

        .filter-select option { background: #1a1043; }

        /* ===== TABLE ===== */
        .table-wrap {
            background: rgba(255,255,255,0.03);
            border: 1px solid rgba(124,58,237,0.18);
            border-radius: 16px;
            overflow: hidden;
            margin-bottom: 44px;
        }

        .table-scroll { overflow-x: auto; }

        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }

        thead tr { background: rgba(124,58,237,0.12); }

        th {
            padding: 12px 16px;
            text-align: left;
            font-size: 11px;
            font-weight: 700;
            color: #9ca3af;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            white-space: nowrap;
        }

        td {
            padding: 12px 16px;
            color: #d1d5db;
            border-top: 1px solid rgba(255,255,255,0.05);
            vertical-align: middle;
        }

        tbody tr:hover { background: rgba(124,58,237,0.06); }

        .order-id {
            font-family: monospace;
            font-size: 11px;
            color: #a78bfa;
        }

        .phone-cell { color: #e5e7eb; font-weight: 600; }

        .badge {
            display: inline-block;
            font-size: 11px;
            font-weight: 700;
            padding: 3px 10px;
            border-radius: 20px;
            white-space: nowrap;
        }

        .badge-paid    { color: #34d399; background: rgba(52,211,153,0.12); }
        .badge-pending { color: #fbbf24; background: rgba(251,191,36,0.12); }
        .badge-failed  { color: #f87171; background: rgba(239,68,68,0.1); }

        .action-btn {
            display: inline-block;
            padding: 5px 10px;
            border-radius: 7px;
            font-size: 11px;
            font-weight: 700;
            text-decoration: none;
            color: #a78bfa;
            background: rgba(124,58,237,0.15);
            border: 1px solid rgba(124,58,237,0.3);
            transition: background 0.15s;
            white-space: nowrap;
        }

        .action-btn:hover { background: rgba(124,58,237,0.3); color: white; }

        .action-btn.green {
            color: #34d399;
            background: rgba(52,211,153,0.1);
            border-color: rgba(52,211,153,0.3);
        }

        .action-btn.green:hover { background: rgba(52,211,153,0.2); }

        .action-btn.disabled {
            color: #374151;
            background: rgba(255,255,255,0.03);
            border-color: rgba(255,255,255,0.06);
            cursor: not-allowed;
            pointer-events: none;
        }

        .actions-cell { display: flex; gap: 6px; flex-wrap: wrap; align-items: center; }

        .table-empty {
            text-align: center;
            padding: 48px 20px;
            color: #4b5563;
            font-size: 14px;
        }

        /* ===== MOBILE ===== */
        .menu-toggle {
            display: none;
            position: fixed;
            top: 14px;
            left: 14px;
            z-index: 300;
            background: rgba(124,58,237,0.85);
            border: none;
            border-radius: 8px;
            width: 38px;
            height: 38px;
            cursor: pointer;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            gap: 5px;
        }

        .menu-toggle span {
            display: block;
            width: 20px;
            height: 2px;
            background: white;
            border-radius: 2px;
        }

        .sidebar-overlay {
            display: none;
            position: fixed;
            inset: 0;
            background: rgba(0,0,0,0.6);
            z-index: 199;
        }

        .sidebar-overlay.open { display: block; }

        @media (max-width: 768px) {
            body { display: block; }

            .sidebar {
                position: fixed;
                left: -240px;
                top: 0;
                height: 100vh;
                z-index: 200;
                transition: left 0.25s ease;
            }

            .sidebar.open { left: 0; }

            .menu-toggle { display: flex; }

            .main { padding: 68px 16px 32px; }

            .stats-grid { grid-template-columns: repeat(2, 1fr); }
        }

        /* ===== MAINTENANCE (SIDEBAR KECIL) ===== */
        .maint-sidebar {
            display: flex; align-items: center; gap: 7px;
            padding: 7px 8px; margin-top: 10px;
            border-radius: 8px; cursor: pointer;
            border: 1px solid rgba(255,255,255,0.06);
            background: rgba(255,255,255,0.03);
        }
        .maint-sidebar:hover { background: rgba(255,255,255,0.06); }
        .maint-dot {
            width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0;
        }
        .maint-dot.on  { background: #ef4444; box-shadow: 0 0 5px #ef4444; }
        .maint-dot.off { background: #34d399; }
        .maint-label { font-size: 11px; color: #6b7280; flex: 1; }
        .maint-label.on { color: #f87171; }
        .maint-gear { font-size: 11px; color: #4b5563; }

        /* ===== MODAL ===== */
        .maint-overlay {
            display: none; position: fixed; inset: 0;
            background: rgba(0,0,0,0.65); z-index: 500;
            align-items: center; justify-content: center;
        }
        .maint-overlay.show { display: flex; }
        .maint-modal {
            background: #1a1043; border: 1px solid rgba(124,58,237,0.35);
            border-radius: 18px; padding: 32px 28px; width: 100%; max-width: 360px;
            position: relative;
        }
        .maint-modal h3 { font-size: 15px; font-weight: 800; margin-bottom: 6px; }
        .maint-modal p  { font-size: 12px; color: #9ca3af; margin-bottom: 18px; }
        .maint-modal input[type=password] {
            width: 100%; padding: 10px 14px;
            background: rgba(255,255,255,0.07);
            border: 1px solid rgba(124,58,237,0.35);
            border-radius: 9px; color: white; font-size: 14px;
            outline: none; margin-bottom: 14px; font-family: inherit;
        }
        .maint-modal input[type=password]:focus { border-color: #7c3aed; }
        .maint-modal-actions { display: flex; gap: 10px; }
        .maint-modal .btn-confirm {
            flex: 1; padding: 10px; border: none; border-radius: 9px;
            font-size: 13px; font-weight: 700; cursor: pointer;
        }
        .maint-modal .btn-confirm.activate   { background: #ef4444; color: white; }
        .maint-modal .btn-confirm.deactivate { background: rgba(52,211,153,0.15); color: #34d399; border: 1px solid rgba(52,211,153,0.3); }
        .maint-modal .btn-cancel {
            padding: 10px 16px; border: 1px solid rgba(255,255,255,0.1);
            background: transparent; border-radius: 9px;
            color: #6b7280; font-size: 13px; cursor: pointer;
        }
        .maint-error { font-size: 12px; color: #f87171; margin-bottom: 10px; display: none; }
        .maint-error.show { display: block; }
    </style>
</head>
<body>

<button class="menu-toggle" id="menuToggle" onclick="toggleSidebar()">
    <span></span><span></span><span></span>
</button>
<div class="sidebar-overlay" id="sidebarOverlay" onclick="toggleSidebar()"></div>

<!-- SIDEBAR -->
<nav class="sidebar" id="sidebar">
    <div class="sidebar-brand">
        ADK PHOTOCOPY
        <span>Panel Admin</span>
        <div class="maint-sidebar" onclick="openMaintModal()" title="Mode Maintenance">
            <span class="maint-dot <?= $isMaintenance ? 'on' : 'off' ?>"></span>
            <span class="maint-label <?= $isMaintenance ? 'on' : '' ?>"><?= $isMaintenance ? 'Maintenance ON' : 'Maintenance' ?></span>
            <span class="maint-gear">⚙</span>
        </div>
    </div>

    <ul class="sidebar-nav">
        <li>
            <a href="<?= BASE_PATH ?>/admin/dashboard.php" class="active">
                <span class="nav-icon">📊</span> Dashboard
            </a>
        </li>
        <li>
            <a href="<?= BASE_PATH ?>/history.php">
                <span class="nav-icon">📋</span> Semua Order
            </a>
        </li>
        <li>
            <a href="#bug-section">
                <span class="nav-icon">🐛</span> Bug Reports
            </a>
        </li>
        <li>
            <a href="<?= BASE_PATH ?>/admin/harga.php">
                <span class="nav-icon">💰</span> Harga Paket
            </a>
        </li>
        <li>
            <a href="<?= BASE_PATH ?>/admin/cleanup.php">
                <span class="nav-icon">🗑</span> Cleanup File
            </a>
        </li>
        <li>
            <a href="<?= BASE_PATH ?>/admin/logs.php">
                <span class="nav-icon">📄</span> Log Error
            </a>
        </li>
        <li>
            <a href="<?= BASE_PATH ?>/admin/export.php">
                <span class="nav-icon">📥</span> Export CSV
            </a>
        </li>
    </ul>

    <a href="<?= BASE_PATH ?>/admin/logout.php" class="btn-logout">Keluar</a>
</nav>

<!-- MAIN -->
<main class="main">
    <div class="page-title">Dashboard</div>
    <div class="page-sub">
        <?= date('l, d F Y') ?> &nbsp;·&nbsp;
        <?php
            $todayLabel = $stats['today'] === 0 ? 'Belum ada order hari ini' : $stats['today'].' order hari ini';
            echo htmlspecialchars($todayLabel);
        ?>
    </div>

    <!-- STATS -->
    <div class="stats-grid">
        <div class="stat-card">
            <div class="stat-label">Total Order</div>
            <div class="stat-value"><?= $stats['total'] ?></div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Hari Ini</div>
            <div class="stat-value purple"><?= $stats['today'] ?></div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Paid</div>
            <div class="stat-value green"><?= $stats['paid'] ?></div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Pending</div>
            <div class="stat-value yellow"><?= $stats['pending'] ?></div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Failed</div>
            <div class="stat-value red"><?= $stats['failed'] ?></div>
        </div>
        <div class="stat-card">
            <div class="stat-label">Total Pendapatan</div>
            <div class="stat-value green sm"><?= fmtRupiah($stats['revenue']) ?></div>
        </div>
    </div>

    <!-- ORDERS TABLE -->
    <div class="section-header">
        <div class="section-title">Semua Order</div>
        <div class="filter-bar">
            <div class="filter-tabs">
                <?php foreach (['all' => 'Semua', 'paid' => 'Paid', 'pending' => 'Pending', 'failed' => 'Failed'] as $k => $label): ?>
                <a href="?status=<?= $k ?>&date=<?= $filterDate ?>"
                   class="filter-tab <?= $filterStatus === $k ? 'active' : '' ?>">
                    <?= $label ?>
                </a>
                <?php endforeach; ?>
            </div>
            <form method="GET" id="dateForm">
                <input type="hidden" name="status" value="<?= htmlspecialchars($filterStatus) ?>">
                <select name="date" class="filter-select" onchange="document.getElementById('dateForm').submit()">
                    <option value="all"   <?= $filterDate === 'all'   ? 'selected' : '' ?>>Semua Waktu</option>
                    <option value="today" <?= $filterDate === 'today' ? 'selected' : '' ?>>Hari Ini</option>
                    <option value="week"  <?= $filterDate === 'week'  ? 'selected' : '' ?>>7 Hari Terakhir</option>
                </select>
            </form>
        </div>
    </div>

    <div class="table-wrap">
        <div class="table-scroll">
            <table>
                <thead>
                    <tr>
                        <th>#</th>
                        <th>ID Transaksi</th>
                        <th>Phone / Tamu</th>
                        <th>Paket</th>
                        <th>Harga</th>
                        <th>Status</th>
                        <th>Tanggal</th>
                        <th>File</th>
                    </tr>
                </thead>
                <tbody>
                <?php if (empty($orders)): ?>
                    <tr><td colspan="8" class="table-empty">Tidak ada data untuk filter ini.</td></tr>
                <?php else: ?>
                    <?php foreach ($orders as $i => $o):
                        $isGuest = !empty($o['guest_token']) && empty($o['phone']);
                        $paketLabel = match($o['paket'] ?? '') {
                            'paket1' => 'Paket 1',
                            'paket2' => 'Paket 2',
                            'paket3' => 'Paket 3',
                            default  => htmlspecialchars($o['paket'] ?? '-')
                        };
                        $badgeClass = match($o['status']) {
                            'paid'    => 'badge-paid',
                            'pending' => 'badge-pending',
                            default   => 'badge-failed'
                        };
                        $badgeLabel = match($o['status']) {
                            'paid'    => 'Paid',
                            'pending' => 'Pending',
                            default   => 'Failed'
                        };
                        $inputFile  = $o['file_input']      ?? '';
                        $outputFile = $o['file_output']     ?? '';
                        $pdfFile    = $o['file_output_pdf'] ?? '';
                    ?>
                    <tr>
                        <td style="color:#4b5563"><?= $i + 1 ?></td>
                        <td><span class="order-id"><?= htmlspecialchars($o['order_id']) ?></span></td>
                        <td class="phone-cell">
                            <?php if ($isGuest): ?>
                                <span style="color:#fbbf24;font-size:11px;font-weight:700">TAMU</span>
                            <?php else: ?>
                                <?= htmlspecialchars($o['phone']) ?>
                            <?php endif; ?>
                        </td>
                        <td><?= $paketLabel ?></td>
                        <td><?= fmtRupiah((float)$o['harga']) ?></td>
                        <td><span class="badge <?= $badgeClass ?>"><?= $badgeLabel ?></span></td>
                        <td style="white-space:nowrap;color:#9ca3af;font-size:12px"><?= tglAdmin($o['created_at']) ?></td>
                        <td>
                            <div class="actions-cell">
                                <?php if ($inputFile): ?>
                                <a href="<?= BASE_PATH ?>/upload/<?= htmlspecialchars(basename($inputFile)) ?>"
                                   class="action-btn" download title="Download file upload asli">Input</a>
                                <?php else: ?>
                                <span class="action-btn disabled">Input</span>
                                <?php endif; ?>

                                <?php if ($outputFile): ?>
                                <a href="<?= BASE_PATH ?>/<?= htmlspecialchars($outputFile) ?>"
                                   class="action-btn green" download title="Download hasil .docx">Docx</a>
                                <?php else: ?>
                                <span class="action-btn disabled">Docx</span>
                                <?php endif; ?>

                                <?php if ($pdfFile): ?>
                                <a href="<?= BASE_PATH ?>/<?= htmlspecialchars($pdfFile) ?>"
                                   class="action-btn green" target="_blank" title="Lihat PDF hasil">PDF</a>
                                <?php else: ?>
                                <span class="action-btn disabled">PDF</span>
                                <?php endif; ?>
                            </div>
                        </td>
                    </tr>
                    <?php endforeach; ?>
                <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>

    <!-- BUG REPORTS -->
    <div id="bug-section" class="section-header">
        <div class="section-title">Laporan Bug</div>
        <?php if (!empty($bugs)): ?>
        <span style="font-size:12px;color:#f87171;font-weight:700"><?= count($bugs) ?> laporan</span>
        <?php endif; ?>
    </div>

    <div class="table-wrap">
        <div class="table-scroll">
            <table>
                <thead>
                    <tr>
                        <th>Order ID</th>
                        <th>Jenis Masalah</th>
                        <th>Phone</th>
                        <th>Paket</th>
                        <th>Deskripsi</th>
                        <th>Tanggal</th>
                    </tr>
                </thead>
                <tbody>
                <?php if (empty($bugs)): ?>
                    <tr><td colspan="6" class="table-empty">Belum ada laporan bug.</td></tr>
                <?php else: ?>
                    <?php foreach ($bugs as $b): ?>
                    <tr>
                        <td><span class="order-id"><?= htmlspecialchars($b['order_id']) ?></span></td>
                        <td style="color:#f87171;font-weight:600"><?= htmlspecialchars($b['jenis']) ?></td>
                        <td class="phone-cell"><?= htmlspecialchars($b['phone'] ?: '-') ?></td>
                        <td><?= htmlspecialchars($b['paket'] ?: '-') ?></td>
                        <td style="color:#9ca3af;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap"
                            title="<?= htmlspecialchars($b['deskripsi'] ?? '') ?>">
                            <?= htmlspecialchars($b['deskripsi'] ?: '-') ?>
                        </td>
                        <td style="white-space:nowrap;color:#9ca3af;font-size:12px"><?= tglAdmin($b['created_at']) ?></td>
                    </tr>
                    <?php endforeach; ?>
                <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>
</main>

<!-- MAINTENANCE MODAL -->
<div class="maint-overlay" id="maintOverlay">
    <div class="maint-modal">
        <h3><?= $isMaintenance ? '🔴 Nonaktifkan Maintenance' : '⚙ Aktifkan Maintenance' ?></h3>
        <p><?= $isMaintenance
            ? 'Server akan kembali normal dan order baru dapat masuk.'
            : 'Order baru akan ditolak. Order lama tetap bisa bayar &amp; unduh.' ?></p>
        <?php if ($maintenanceError): ?>
        <div class="maint-error show"><?= htmlspecialchars($maintenanceError) ?></div>
        <?php else: ?>
        <div class="maint-error" id="maintErrMsg"></div>
        <?php endif; ?>
        <form method="POST" id="maintForm">
            <input type="password" name="maintenance_password" id="maintPwd"
                   placeholder="Sandi admin" autocomplete="current-password" required>
            <div class="maint-modal-actions">
                <button type="submit" name="toggle_maintenance"
                        class="btn-confirm <?= $isMaintenance ? 'deactivate' : 'activate' ?>">
                    <?= $isMaintenance ? 'Nonaktifkan' : 'Aktifkan' ?>
                </button>
                <button type="button" class="btn-cancel" onclick="closeMaintModal()">Batal</button>
            </div>
        </form>
    </div>
</div>

<script>
function toggleSidebar() {
    document.getElementById('sidebar').classList.toggle('open');
    document.getElementById('sidebarOverlay').classList.toggle('open');
}
function openMaintModal() {
    document.getElementById('maintOverlay').classList.add('show');
    setTimeout(function(){ var p = document.getElementById('maintPwd'); if(p) p.focus(); }, 100);
}
function closeMaintModal() {
    document.getElementById('maintOverlay').classList.remove('show');
}
document.getElementById('maintOverlay').addEventListener('click', function(e){
    if (e.target === this) closeMaintModal();
});
<?php if ($maintenanceError): ?>
openMaintModal();
<?php endif; ?>
<?php if ($maintenanceSuccess): ?>
(function(){
    var t = document.createElement('div');
    t.textContent = <?= json_encode($maintenanceSuccess) ?>;
    t.style.cssText = 'position:fixed;top:20px;right:20px;background:#065f46;color:#6ee7b7;padding:12px 20px;border-radius:10px;font-size:14px;font-weight:600;z-index:9999;box-shadow:0 4px 20px rgba(0,0,0,.4)';
    document.body.appendChild(t);
    setTimeout(function(){ t.remove(); }, 3000);
})();
<?php endif; ?>
</script>
</body>
</html>
