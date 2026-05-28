<?php
date_default_timezone_set('Asia/Jakarta');
require_once __DIR__ . '/_guard.php';

$db        = getDB();
$rootDir   = __DIR__ . '/../';
$deleted   = [];
$errors    = [];

// ── Hapus file jika diminta ──────────────────────────────
if ($_SERVER['REQUEST_METHOD'] === 'POST' && !empty($_POST['delete'])) {
    foreach ((array)$_POST['delete'] as $relPath) {
        $relPath = ltrim(str_replace(['..', '\\'], ['', '/'], $relPath), '/');
        if (!preg_match('#^(upload|hasil)/#', $relPath)) continue;
        $abs = $rootDir . $relPath;
        if (file_exists($abs)) {
            if (unlink($abs)) $deleted[] = $relPath;
            else              $errors[]  = $relPath;
        }
    }
}

// ── Ambil order yang sudah selesai (paid/failed) > 7 hari ─
$cutoff = date('Y-m-d H:i:s', strtotime('-7 days'));
$stmt   = $db->prepare(
    "SELECT id, file_input, file_output, file_output_pdf, status, created_at
     FROM orders
     WHERE status IN ('paid','failed') AND created_at < :cutoff
     ORDER BY created_at ASC"
);
$stmt->execute([':cutoff' => $cutoff]);
$oldOrders = $stmt->fetchAll();

// Kumpulkan file-file kandidat hapus
$candidates = [];
foreach ($oldOrders as $o) {
    $files = [];
    if ($o['file_input'])      $files[] = 'upload/' . $o['file_input'];
    if ($o['file_output'])     $files[] = $o['file_output'];
    if ($o['file_output_pdf']) $files[] = $o['file_output_pdf'];
    foreach ($files as $f) {
        $abs = $rootDir . $f;
        if (file_exists($abs)) {
            $candidates[] = [
                'rel'     => $f,
                'size'    => filesize($abs),
                'order'   => $o['id'],
                'status'  => $o['status'],
                'tanggal' => $o['created_at'],
            ];
        }
    }
}

function fmtBytes(int $b): string {
    if ($b >= 1048576) return round($b / 1048576, 1) . ' MB';
    if ($b >= 1024)    return round($b / 1024, 1) . ' KB';
    return $b . ' B';
}

$totalSize = array_sum(array_column($candidates, 'size'));
?>
<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Cleanup File — ADK Admin</title>
<link rel="icon" type="image/jpeg" href="../favicon.jpg">
<style>
  body { font-family: sans-serif; background: #0d0b1e; color: #e5e7eb; margin: 0; padding: 24px; }
  h1   { color: white; font-size: 22px; margin-bottom: 4px; }
  .sub { color: #6b7280; font-size: 13px; margin-bottom: 24px; }
  .alert-ok  { background: rgba(52,211,153,.1); border: 1px solid rgba(52,211,153,.3); border-radius: 8px; padding: 10px 16px; margin-bottom: 16px; font-size: 13px; color: #34d399; }
  .alert-err { background: rgba(239,68,68,.1);  border: 1px solid rgba(239,68,68,.3);  border-radius: 8px; padding: 10px 16px; margin-bottom: 16px; font-size: 13px; color: #f87171; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th    { text-align: left; padding: 8px 12px; color: #6b7280; font-size: 11px; text-transform: uppercase; border-bottom: 1px solid rgba(255,255,255,.07); }
  td    { padding: 8px 12px; border-bottom: 1px solid rgba(255,255,255,.05); vertical-align: middle; }
  tr:hover td { background: rgba(255,255,255,.03); }
  input[type=checkbox] { accent-color: #7c3aed; width: 15px; height: 15px; }
  .badge-paid   { color: #34d399; }
  .badge-failed { color: #f87171; }
  .btn-hapus {
    padding: 9px 20px; background: #ef4444; border: none; border-radius: 8px;
    color: white; font-size: 13px; font-weight: 700; cursor: pointer;
  }
  .btn-hapus:hover { background: #dc2626; }
  .btn-back  {
    display: inline-block; margin-bottom: 20px;
    color: #a78bfa; font-size: 13px; text-decoration: none;
  }
  .btn-back:hover { color: white; }
  .summary { margin-bottom: 16px; font-size: 13px; color: #9ca3af; }
  .summary strong { color: white; }
  .empty { color: #6b7280; font-size: 14px; padding: 40px 0; text-align: center; }
  .actions { display: flex; gap: 12px; align-items: center; margin-top: 16px; }
  label.sel-all { font-size: 12px; color: #9ca3af; cursor: pointer; }
</style>
</head>
<body>

<a href="dashboard.php" class="btn-back">← Dashboard</a>
<h1>Cleanup File Lama</h1>
<p class="sub">Order dengan status paid/failed yang dibuat lebih dari 7 hari lalu.</p>

<?php if ($deleted): ?>
<div class="alert-ok">Berhasil dihapus: <?= implode(', ', array_map('htmlspecialchars', $deleted)) ?></div>
<?php endif; ?>
<?php if ($errors): ?>
<div class="alert-err">Gagal menghapus: <?= implode(', ', array_map('htmlspecialchars', $errors)) ?></div>
<?php endif; ?>

<?php if (empty($candidates)): ?>
<div class="empty">Tidak ada file yang perlu dihapus.</div>
<?php else: ?>

<div class="summary">
    Ditemukan <strong><?= count($candidates) ?></strong> file
    · Total ukuran: <strong><?= fmtBytes($totalSize) ?></strong>
</div>

<form method="POST">
    <table>
        <thead>
            <tr>
                <th><input type="checkbox" id="selAll" onchange="toggleAll(this)"> Pilih</th>
                <th>File</th>
                <th>Ukuran</th>
                <th>Order #</th>
                <th>Status</th>
                <th>Tanggal Order</th>
            </tr>
        </thead>
        <tbody>
        <?php foreach ($candidates as $c): ?>
        <tr>
            <td><input type="checkbox" name="delete[]" value="<?= htmlspecialchars($c['rel']) ?>" class="cb-file"></td>
            <td style="word-break:break-all;max-width:300px"><?= htmlspecialchars($c['rel']) ?></td>
            <td><?= fmtBytes($c['size']) ?></td>
            <td>#<?= $c['order'] ?></td>
            <td class="badge-<?= $c['status'] ?>"><?= $c['status'] ?></td>
            <td><?= htmlspecialchars($c['tanggal']) ?></td>
        </tr>
        <?php endforeach; ?>
        </tbody>
    </table>

    <div class="actions">
        <button type="submit" class="btn-hapus"
                onclick="return confirm('Yakin ingin menghapus file yang dipilih? Tindakan ini tidak bisa dibatalkan.')">
            Hapus File Terpilih
        </button>
        <label class="sel-all" for="selAll">Pilih semua</label>
    </div>
</form>

<script>
function toggleAll(master) {
    document.querySelectorAll('.cb-file').forEach(cb => cb.checked = master.checked);
}
</script>

<?php endif; ?>
</body>
</html>
