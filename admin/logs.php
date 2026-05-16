<?php
date_default_timezone_set('Asia/Jakarta');
require_once __DIR__ . '/_guard.php';

$logPath = __DIR__ . '/../logs/error.log';
$cleared = false;

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['clear_log'])) {
    file_put_contents($logPath, '');
    $cleared = true;
}

$logExists  = file_exists($logPath);
$logContent = $logExists ? file_get_contents($logPath) : '';
$logSize    = $logExists ? filesize($logPath) : 0;
$logMtime   = $logExists && $logSize > 0 ? date('d M Y H:i:s', filemtime($logPath)) : null;

function fmtBytes(int $b): string {
    if ($b >= 1048576) return round($b / 1048576, 1) . ' MB';
    if ($b >= 1024)    return round($b / 1024, 1) . ' KB';
    return $b . ' B';
}

// Balik urutan: entri terbaru di atas (dipisah oleh "---\n")
$entries = [];
if ($logContent) {
    $entries = array_filter(array_map('trim', explode('---', $logContent)));
    $entries = array_reverse(array_values($entries));
}
?>
<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Log Error — ADK Admin</title>
<link rel="icon" type="image/png" href="../favicon.png">
<style>
  body { font-family: sans-serif; background: #0d0b1e; color: #e5e7eb; margin: 0; padding: 24px; }
  h1   { color: white; font-size: 22px; margin-bottom: 4px; }
  .sub { color: #6b7280; font-size: 13px; margin-bottom: 24px; }
  .alert-ok  { background: rgba(52,211,153,.1); border: 1px solid rgba(52,211,153,.3); border-radius: 8px; padding: 10px 16px; margin-bottom: 16px; font-size: 13px; color: #34d399; }
  .meta { font-size: 12px; color: #6b7280; margin-bottom: 16px; display: flex; gap: 20px; flex-wrap: wrap; }
  .meta span { display: flex; align-items: center; gap: 6px; }
  .actions { display: flex; gap: 10px; margin-bottom: 20px; }
  .btn-clear {
    padding: 8px 18px; background: rgba(239,68,68,.15); border: 1px solid rgba(239,68,68,.3);
    border-radius: 8px; color: #f87171; font-size: 13px; font-weight: 700; cursor: pointer;
  }
  .btn-clear:hover { background: rgba(239,68,68,.25); }
  .btn-back { display: inline-block; margin-bottom: 20px; color: #a78bfa; font-size: 13px; text-decoration: none; }
  .btn-back:hover { color: white; }
  .entry {
    background: rgba(255,255,255,.03); border: 1px solid rgba(255,255,255,.07);
    border-radius: 10px; padding: 14px 16px; margin-bottom: 12px;
    font-family: monospace; font-size: 12px; white-space: pre-wrap; word-break: break-all;
    line-height: 1.6; color: #d1d5db;
  }
  .entry-head { color: #fbbf24; font-weight: 700; margin-bottom: 6px; }
  .empty { color: #6b7280; font-size: 14px; padding: 40px 0; text-align: center; }
  .count { background: rgba(124,58,237,.2); color: #a78bfa; font-size: 11px; padding: 2px 8px; border-radius: 20px; margin-left: 8px; font-weight: 700; }
</style>
</head>
<body>

<a href="dashboard.php" class="btn-back">← Dashboard</a>
<h1>Log Error <span class="count"><?= count($entries) ?> entri</span></h1>
<p class="sub">Log dari proses dokumen yang gagal. Entri terbaru ditampilkan di atas.</p>

<?php if ($cleared): ?>
<div class="alert-ok">Log berhasil dikosongkan.</div>
<?php endif; ?>

<?php if ($logMtime): ?>
<div class="meta">
    <span>📁 Ukuran: <?= fmtBytes($logSize) ?></span>
    <span>🕐 Diperbarui: <?= $logMtime ?></span>
</div>
<?php endif; ?>

<?php if (!empty($entries)): ?>
<div class="actions">
    <form method="POST" onsubmit="return confirm('Yakin ingin mengosongkan log? Tindakan ini tidak bisa dibatalkan.')">
        <button type="submit" name="clear_log" class="btn-clear">Kosongkan Log</button>
    </form>
</div>

<?php foreach ($entries as $entry): ?>
<?php
    $lines = explode("\n", trim($entry), 2);
    $head  = $lines[0] ?? '';
    $body  = $lines[1] ?? '';
?>
<div class="entry">
    <div class="entry-head"><?= htmlspecialchars($head) ?></div><?php if ($body): ?><?= htmlspecialchars($body) ?><?php endif; ?>
</div>
<?php endforeach; ?>

<?php else: ?>
<div class="empty">Tidak ada entri log.</div>
<?php endif; ?>

</body>
</html>
