<?php
date_default_timezone_set('Asia/Jakarta');
require_once __DIR__ . '/_guard.php';

$db = getDB();

// ── Filter params ──────────────────────────────────────────────────────────
$filterStatus = $_GET['status'] ?? 'all';
$filterFrom   = $_GET['from']   ?? '';
$filterTo     = $_GET['to']     ?? '';
$doExport     = isset($_GET['download']);

$allowed = ['all','paid','pending','failed'];
if (!in_array($filterStatus, $allowed)) $filterStatus = 'all';

// Bangun query
$where  = [];
$params = [];
if ($filterStatus !== 'all') {
    $where[]           = "status = :status";
    $params[':status'] = $filterStatus;
}
if ($filterFrom) {
    $where[]          = "DATE(created_at) >= :from";
    $params[':from']  = $filterFrom;
}
if ($filterTo) {
    $where[]        = "DATE(created_at) <= :to";
    $params[':to']  = $filterTo;
}

$sql = "SELECT id, order_id, phone, paket, harga, status, created_at FROM orders";
if ($where) $sql .= " WHERE " . implode(" AND ", $where);
$sql .= " ORDER BY created_at DESC";

$stmt = $db->prepare($sql);
$stmt->execute($params);
$orders = $stmt->fetchAll();

// ── Download CSV ───────────────────────────────────────────────────────────
if ($doExport) {
    $filename = 'adk-orders-' . date('Ymd-His') . '.csv';
    header('Content-Type: text/csv; charset=UTF-8');
    header('Content-Disposition: attachment; filename="' . $filename . '"');
    header('Cache-Control: no-cache');

    $out = fopen('php://output', 'w');
    // BOM agar Excel baca UTF-8 dengan benar
    fwrite($out, "\xEF\xBB\xBF");
    fputcsv($out, ['ID', 'Order ID', 'No. HP', 'Paket', 'Harga', 'Status', 'Tanggal'], ';');
    foreach ($orders as $o) {
        fputcsv($out, [
            $o['id'],
            $o['order_id'],
            $o['phone'] ?? '(tamu)',
            $o['paket'],
            $o['harga'],
            $o['status'],
            $o['created_at'],
        ], ';');
    }
    fclose($out);
    exit;
}

// ── Halaman filter ─────────────────────────────────────────────────────────
function fmtRupiah(float $n): string { return 'Rp ' . number_format($n, 0, ',', '.'); }
$totalRevenue = array_sum(array_map(fn($o) => $o['status']==='paid' ? $o['harga'] : 0, $orders));
?>
<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Export CSV — ADK Admin</title>
<style>
  body { font-family: sans-serif; background: #0d0b1e; color: #e5e7eb; margin: 0; padding: 24px; }
  h1   { color: white; font-size: 22px; margin-bottom: 4px; }
  .sub { color: #6b7280; font-size: 13px; margin-bottom: 24px; }
  .filter-card {
    background: rgba(255,255,255,.04); border: 1px solid rgba(124,58,237,.2);
    border-radius: 16px; padding: 20px 22px; margin-bottom: 24px;
    display: flex; gap: 16px; flex-wrap: wrap; align-items: flex-end;
  }
  .fg { display: flex; flex-direction: column; gap: 5px; }
  .fg label { font-size: 11px; font-weight: 700; color: #6b7280; text-transform: uppercase; letter-spacing: .4px; }
  .fg select, .fg input[type=date] {
    padding: 8px 12px; background: rgba(255,255,255,.07);
    border: 1px solid rgba(124,58,237,.35); border-radius: 9px;
    color: white; font-size: 13px; outline: none; font-family: inherit;
  }
  .fg select option { background: #1a1043; }
  .btn-filter {
    padding: 9px 20px; background: rgba(124,58,237,.3); border: 1px solid rgba(124,58,237,.4);
    border-radius: 9px; color: white; font-size: 13px; font-weight: 700; cursor: pointer; text-decoration: none;
  }
  .btn-filter:hover { background: rgba(124,58,237,.5); }
  .btn-download {
    padding: 9px 20px; background: #7c3aed; border: none;
    border-radius: 9px; color: white; font-size: 13px; font-weight: 700; cursor: pointer; text-decoration: none; display: inline-block;
  }
  .btn-download:hover { background: #6d28d9; }
  .summary { font-size: 13px; color: #9ca3af; margin-bottom: 16px; }
  .summary strong { color: white; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th { text-align: left; padding: 8px 12px; color: #6b7280; font-size: 11px; text-transform: uppercase; border-bottom: 1px solid rgba(255,255,255,.07); }
  td { padding: 8px 12px; border-bottom: 1px solid rgba(255,255,255,.04); }
  tr:hover td { background: rgba(255,255,255,.02); }
  .badge-paid    { color: #34d399; }
  .badge-pending { color: #fbbf24; }
  .badge-failed  { color: #f87171; }
  .empty { color: #6b7280; font-size: 14px; padding: 40px 0; text-align: center; }
  .btn-back { display: inline-block; margin-bottom: 20px; color: #a78bfa; font-size: 13px; text-decoration: none; }
  .btn-back:hover { color: white; }
</style>
</head>
<body>

<a href="dashboard.php" class="btn-back">← Dashboard</a>
<h1>Export CSV</h1>
<p class="sub">Filter order lalu unduh sebagai file CSV yang bisa dibuka di Excel.</p>

<form method="GET">
    <div class="filter-card">
        <div class="fg">
            <label>Status</label>
            <select name="status">
                <?php foreach (['all'=>'Semua','paid'=>'Paid','pending'=>'Pending','failed'=>'Failed'] as $k=>$v): ?>
                <option value="<?= $k ?>" <?= $filterStatus===$k?'selected':'' ?>><?= $v ?></option>
                <?php endforeach; ?>
            </select>
        </div>
        <div class="fg">
            <label>Dari Tanggal</label>
            <input type="date" name="from" value="<?= htmlspecialchars($filterFrom) ?>">
        </div>
        <div class="fg">
            <label>Sampai Tanggal</label>
            <input type="date" name="to" value="<?= htmlspecialchars($filterTo) ?>">
        </div>
        <button type="submit" class="btn-filter">Terapkan Filter</button>
    </div>
</form>

<?php if (!empty($orders)): ?>
<div class="summary">
    Menampilkan <strong><?= count($orders) ?> order</strong>
    · Total paid: <strong><?= fmtRupiah($totalRevenue) ?></strong>
    &nbsp;
    <a class="btn-download" href="?status=<?= $filterStatus ?>&from=<?= urlencode($filterFrom) ?>&to=<?= urlencode($filterTo) ?>&download=1">
        ⬇ Download CSV
    </a>
</div>

<table>
    <thead>
        <tr>
            <th>#</th><th>Order ID</th><th>No. HP</th><th>Paket</th><th>Harga</th><th>Status</th><th>Tanggal</th>
        </tr>
    </thead>
    <tbody>
    <?php foreach ($orders as $o): ?>
    <tr>
        <td><?= $o['id'] ?></td>
        <td style="font-family:monospace;font-size:11px"><?= htmlspecialchars($o['order_id']) ?></td>
        <td><?= htmlspecialchars($o['phone'] ?? '<em style="color:#4b5563">tamu</em>') ?></td>
        <td><?= htmlspecialchars($o['paket']) ?></td>
        <td><?= fmtRupiah($o['harga']) ?></td>
        <td class="badge-<?= $o['status'] ?>"><?= $o['status'] ?></td>
        <td><?= htmlspecialchars($o['created_at']) ?></td>
    </tr>
    <?php endforeach; ?>
    </tbody>
</table>
<?php else: ?>
<div class="empty">Tidak ada order yang cocok dengan filter.</div>
<?php endif; ?>

</body>
</html>
