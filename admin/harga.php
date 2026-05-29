<?php
date_default_timezone_set('Asia/Jakarta');
require_once __DIR__ . '/_guard.php';

$configPath = __DIR__ . '/../config/harga.json';
$harga      = json_decode(file_get_contents($configPath), true) ?: ['paket1'=>5000,'paket2'=>10000,'paket3'=>10000,'paket4'=>15000];
$success    = false;
$error      = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $p1 = (int)($_POST['paket1'] ?? 0);
    $p2 = (int)($_POST['paket2'] ?? 0);
    $p3 = (int)($_POST['paket3'] ?? 0);
    $p4 = (int)($_POST['paket4'] ?? 0);

    if ($p1 < 1000 || $p2 < 1000 || $p3 < 1000 || $p4 < 1000) {
        $error = 'Harga minimal Rp 1.000 per paket.';
    } else {
        $harga   = ['paket1' => $p1, 'paket2' => $p2, 'paket3' => $p3, 'paket4' => $p4];
        $written = file_put_contents($configPath, json_encode($harga, JSON_PRETTY_PRINT));
        if ($written !== false) {
            $success = true;
        } else {
            $error = 'Gagal menyimpan file harga. Pastikan direktori config/ dapat ditulis oleh server (chmod 775 atau chown www-data).';
        }
    }
}

function fmt(int $n): string { return 'Rp ' . number_format($n, 0, ',', '.'); }
?>
<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Harga Paket — ADK Admin</title>
<link rel="icon" type="image/png" href="../favicon.png">
<style>
  body { font-family: sans-serif; background: #0d0b1e; color: #e5e7eb; margin: 0; padding: 24px; }
  h1   { color: white; font-size: 22px; margin-bottom: 4px; }
  .sub { color: #6b7280; font-size: 13px; margin-bottom: 28px; }
  .alert-ok  { background: rgba(52,211,153,.1); border: 1px solid rgba(52,211,153,.3); border-radius: 8px; padding: 10px 16px; margin-bottom: 20px; font-size: 13px; color: #34d399; }
  .alert-err { background: rgba(239,68,68,.1);  border: 1px solid rgba(239,68,68,.3);  border-radius: 8px; padding: 10px 16px; margin-bottom: 20px; font-size: 13px; color: #f87171; }
  .cards { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 16px; margin-bottom: 28px; }
  .card {
    background: rgba(255,255,255,.05); border: 1px solid rgba(124,58,237,.2);
    border-radius: 16px; padding: 20px 22px;
  }
  .card-label { font-size: 11px; font-weight: 700; color: #6b7280; text-transform: uppercase; letter-spacing: .5px; margin-bottom: 6px; }
  .card-desc  { font-size: 12px; color: #9ca3af; margin-bottom: 14px; line-height: 1.5; }
  label { display: block; font-size: 12px; color: #9ca3af; margin-bottom: 5px; }
  input[type=number] {
    width: 100%; padding: 9px 13px;
    background: rgba(255,255,255,.07); border: 1px solid rgba(124,58,237,.35);
    border-radius: 9px; color: white; font-size: 15px; font-weight: 700;
    outline: none; font-family: inherit;
  }
  input[type=number]:focus { border-color: #7c3aed; }
  .hint { font-size: 11px; color: #6b7280; margin-top: 5px; }
  .btn-save {
    padding: 11px 28px; background: #7c3aed; border: none; border-radius: 10px;
    color: white; font-size: 14px; font-weight: 700; cursor: pointer;
  }
  .btn-save:hover { background: #6d28d9; }
  .btn-back { display: inline-block; margin-bottom: 20px; color: #a78bfa; font-size: 13px; text-decoration: none; }
  .btn-back:hover { color: white; }
</style>
</head>
<body>

<a href="dashboard.php" class="btn-back">← Dashboard</a>
<h1>Harga Paket</h1>
<p class="sub">Perubahan harga berlaku untuk order baru. Order yang sudah ada tidak terpengaruh.</p>

<?php if ($success): ?>
<div class="alert-ok">Harga berhasil disimpan.</div>
<?php endif; ?>
<?php if ($error): ?>
<div class="alert-err"><?= htmlspecialchars($error) ?></div>
<?php endif; ?>

<form method="POST">
    <div class="cards">
        <div class="card">
            <div class="card-label">Paket 1 — Angka</div>
            <div class="card-desc">Penomoran halaman angka biasa (1, 2, 3 …). Tanpa cover tersembunyi.</div>
            <label>Harga (Rp)</label>
            <input type="number" name="paket1" value="<?= $harga['paket1'] ?>" min="1000" step="500" required>
            <div class="hint">Saat ini: <?= fmt($harga['paket1']) ?></div>
        </div>
        <div class="card">
            <div class="card-label">Paket 2 — Romawi + Angka</div>
            <div class="card-desc">Penomoran halaman romawi dan angka dengan posisi bebas pilih.</div>
            <label>Harga (Rp)</label>
            <input type="number" name="paket2" value="<?= $harga['paket2'] ?>" min="1000" step="500" required>
            <div class="hint">Saat ini: <?= fmt($harga['paket2']) ?></div>
        </div>
        <div class="card">
            <div class="card-label">Paket 3 — Skripsi ⭐</div>
            <div class="card-desc">Romawi + angka, posisi tetap sesuai standar skripsi. Paling populer.</div>
            <label>Harga (Rp)</label>
            <input type="number" name="paket3" value="<?= $harga['paket3'] ?>" min="1000" step="500" required>
            <div class="hint">Saat ini: <?= fmt($harga['paket3']) ?></div>
        </div>
        <div class="card">
            <div class="card-label">Paket 4 — Custom</div>
            <div class="card-desc">Romawi + angka, 3 posisi dapat dipilih bebas (custom). Segera hadir.</div>
            <label>Harga (Rp)</label>
            <input type="number" name="paket4" value="<?= $harga['paket4'] ?? 15000 ?>" min="1000" step="500" required>
            <div class="hint">Saat ini: <?= fmt($harga['paket4'] ?? 15000) ?></div>
        </div>
    </div>
    <button type="submit" class="btn-save">Simpan Harga</button>
</form>

</body>
</html>
