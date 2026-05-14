<?php
date_default_timezone_set('Asia/Jakarta');
require_once __DIR__ . '/_guard.php';

$testDir = __DIR__ . '/../temp/test';
if (!is_dir($testDir)) @mkdir($testDir, 0775, true);

// ── Auto-bersihkan file test lama (>2 jam) ────────────────────────────────
foreach (glob($testDir . '/*.docx') ?: [] as $f) {
    if (filemtime($f) < time() - 7200) @unlink($f);
}

// ── Handle download ───────────────────────────────────────────────────────
if ($_SERVER['REQUEST_METHOD'] === 'GET' && isset($_GET['dl'])) {
    $safe = preg_replace('/[^a-zA-Z0-9_-]/', '', basename($_GET['dl']));
    $path = $testDir . '/' . $safe . '.docx';
    if ($safe && file_exists($path)) {
        header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        header('Content-Disposition: attachment; filename="ADK_Test_' . date('Ymd_His') . '.docx"');
        header('Content-Length: ' . filesize($path));
        readfile($path);
        @unlink($path);
        exit;
    }
    http_response_code(404);
    die('File tidak ditemukan atau sudah dihapus.');
}

// ── Handle delete file test ───────────────────────────────────────────────
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['del'])) {
    $safe = preg_replace('/[^a-zA-Z0-9_-]/', '', basename($_POST['del']));
    @unlink($testDir . '/' . $safe . '.docx');
    header('Location: test.php');
    exit;
}

// ── Handle proses ─────────────────────────────────────────────────────────
$result  = null;
$errMsg  = null;
$outKey  = null;

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['file'])) {
    do {
        if ($_FILES['file']['error'] !== UPLOAD_ERR_OK) {
            $errMsg = 'Upload gagal (kode ' . $_FILES['file']['error'] . ').'; break;
        }
        $ext = strtolower(pathinfo($_FILES['file']['name'], PATHINFO_EXTENSION));
        if ($ext !== 'docx') { $errMsg = 'Hanya file .docx yang didukung.'; break; }
        if ($_FILES['file']['size'] > 30 * 1024 * 1024) { $errMsg = 'Ukuran file melebihi 30 MB.'; break; }

        $ts      = date('YmdHis') . '_' . bin2hex(random_bytes(4));
        $inPath  = $testDir . '/' . $ts . '_in.docx';
        $outKey  = $ts . '_out';
        $outPath = $testDir . '/' . $outKey . '.docx';

        if (!move_uploaded_file($_FILES['file']['tmp_name'], $inPath)) {
            $errMsg = 'Gagal menyimpan file.'; $outKey = null; break;
        }

        $paket  = in_array($_POST['paket'] ?? '', ['paket1','paket2','paket3'])
                  ? $_POST['paket'] : 'paket3';
        $font   = $_POST['font']         ?? 'Times New Roman';
        $size   = $_POST['size']         ?? '12 pt';
        $hidden = $_POST['hidden_cover'] ?? 'Ya';
        $posisi = $_POST['posisi']       ?? 'Tengah Bawah';

        $python = PYTHON_EXE;
        $script = PYTHON_SCRIPT_DIR . '/main.py';

        $cmd = "\"$python\" \"$script\" "
            . escapeshellarg($inPath)  . ' '
            . escapeshellarg($outPath) . ' '
            . escapeshellarg($paket)   . ' '
            . escapeshellarg($font)    . ' '
            . escapeshellarg($size)    . ' '
            . escapeshellarg($hidden)  . ' '
            . escapeshellarg($posisi)  . ' 2>&1';

        $tStart = microtime(true);
        exec($cmd, $rawOut, $exitCode);
        $elapsed = round(microtime(true) - $tStart, 2);

        @unlink($inPath);

        $parsed = null;
        foreach (array_reverse($rawOut) as $line) {
            $t = trim($line);
            if (strlen($t) > 0 && $t[0] === '{') {
                $parsed = json_decode($t, true);
                break;
            }
        }

        if ($exitCode !== 0 || !file_exists($outPath)) {
            $errMsg  = $parsed['message'] ?? ('Python exit code ' . $exitCode);
            $errCode = $parsed['code']    ?? '';
            if ($errCode) $errMsg = "[{$errCode}] {$errMsg}";
            $outKey  = null;
            break;
        }

        $result = [
            'paket'    => $paket,
            'elapsed'  => $elapsed,
            'outKey'   => $outKey,
            'outSize'  => round(filesize($outPath) / 1024, 1),
            'sections' => $parsed['total_sections'] ?? '?',
            'bab'      => $parsed['detected_bab']   ?? [],
            'secs'     => $parsed['sections']        ?? [],
            'raw'      => implode("\n", $rawOut),
        ];

    } while (false);
}

// ── File test yang masih ada ──────────────────────────────────────────────
$orphans = [];
foreach (glob($testDir . '/*_out.docx') ?: [] as $f) {
    $key = basename($f, '.docx');
    $orphans[] = [
        'key'  => $key,
        'size' => round(filesize($f) / 1024, 1),
        'age'  => round((time() - filemtime($f)) / 60),
    ];
}
usort($orphans, fn($a,$b) => strcmp($b['key'], $a['key']));
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Test Dokumen — ADK Admin</title>
    <style>
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

        body {
            font-family: 'Segoe UI', sans-serif;
            background: linear-gradient(135deg, #0e0b28, #1a1043);
            min-height: 100vh;
            color: white;
            padding: 32px 24px;
        }

        .top-bar {
            display: flex;
            align-items: center;
            gap: 14px;
            margin-bottom: 28px;
        }

        .back-link {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 7px 14px;
            background: rgba(124,58,237,0.15);
            border: 1px solid rgba(124,58,237,0.3);
            border-radius: 9px;
            color: #a78bfa;
            text-decoration: none;
            font-size: 13px;
            font-weight: 700;
            transition: background 0.15s;
        }
        .back-link:hover { background: rgba(124,58,237,0.28); color: white; }

        .page-title { font-size: 20px; font-weight: 800; }
        .page-sub   { font-size: 12px; color: #6b7280; margin-top: 3px; }

        .layout {
            display: grid;
            grid-template-columns: 380px 1fr;
            gap: 22px;
            align-items: start;
        }

        @media (max-width: 860px) {
            .layout { grid-template-columns: 1fr; }
        }

        /* ── Card ── */
        .card {
            background: rgba(255,255,255,0.04);
            border: 1px solid rgba(124,58,237,0.2);
            border-radius: 18px;
            padding: 24px;
        }

        .card-title {
            font-size: 14px;
            font-weight: 800;
            margin-bottom: 18px;
            color: #e5e7eb;
        }

        /* ── Form ── */
        .field { margin-bottom: 16px; }

        .field label {
            display: block;
            font-size: 11px;
            font-weight: 700;
            color: #9ca3af;
            text-transform: uppercase;
            letter-spacing: 0.4px;
            margin-bottom: 7px;
        }

        .field input[type=file],
        .field select {
            width: 100%;
            padding: 10px 13px;
            background: rgba(255,255,255,0.06);
            border: 1px solid rgba(124,58,237,0.3);
            border-radius: 10px;
            color: white;
            font-size: 13px;
            outline: none;
            font-family: inherit;
            cursor: pointer;
            transition: border-color 0.15s;
        }

        .field select option { background: #1a1043; }
        .field input[type=file]:focus,
        .field select:focus { border-color: #7c3aed; }

        .field-row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }

        .paket-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 8px;
        }

        .paket-btn {
            padding: 10px 8px;
            border: 2px solid rgba(124,58,237,0.25);
            border-radius: 10px;
            background: rgba(255,255,255,0.03);
            color: #9ca3af;
            font-size: 12px;
            font-weight: 700;
            cursor: pointer;
            text-align: center;
            transition: all 0.15s;
        }

        .paket-btn:hover { border-color: rgba(124,58,237,0.5); color: white; }
        .paket-btn.selected {
            border-color: #7c3aed;
            background: rgba(124,58,237,0.2);
            color: white;
        }

        .paket-btn .pb-label { font-size: 11px; color: #6b7280; font-weight: 500; margin-top: 3px; }
        .paket-btn.selected .pb-label { color: #c4b5fd; }

        .btn-proses {
            width: 100%;
            padding: 13px;
            background: #7c3aed;
            border: none;
            border-radius: 11px;
            color: white;
            font-size: 14px;
            font-weight: 800;
            cursor: pointer;
            margin-top: 4px;
            transition: background 0.15s, opacity 0.15s;
            font-family: inherit;
        }
        .btn-proses:hover  { background: #6d28d9; }
        .btn-proses:active { opacity: 0.85; }

        /* ── Result ── */
        .result-empty {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            min-height: 220px;
            color: #374151;
            font-size: 13px;
            gap: 10px;
        }

        .result-icon { font-size: 40px; }

        .badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 700;
        }
        .badge-ok  { background: rgba(52,211,153,0.12); color: #34d399; }
        .badge-err { background: rgba(239,68,68,0.1);  color: #f87171; }

        .result-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 10px;
            margin-bottom: 18px;
        }

        .result-stats {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(100px,1fr));
            gap: 10px;
            margin-bottom: 18px;
        }

        .stat-mini {
            background: rgba(255,255,255,0.04);
            border: 1px solid rgba(124,58,237,0.15);
            border-radius: 10px;
            padding: 10px 12px;
        }

        .stat-mini-label { font-size: 10px; color: #6b7280; font-weight: 700; text-transform: uppercase; margin-bottom: 5px; }
        .stat-mini-val   { font-size: 20px; font-weight: 800; color: #a78bfa; }

        .bab-list {
            list-style: none;
            display: flex;
            flex-direction: column;
            gap: 6px;
            margin-bottom: 18px;
        }

        .bab-item {
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 8px 12px;
            background: rgba(255,255,255,0.03);
            border: 1px solid rgba(255,255,255,0.06);
            border-radius: 8px;
            font-size: 12px;
            color: #d1d5db;
        }

        .bab-num {
            flex-shrink: 0;
            width: 22px; height: 22px;
            border-radius: 50%;
            background: rgba(124,58,237,0.25);
            color: #a78bfa;
            font-size: 10px;
            font-weight: 800;
            display: flex; align-items: center; justify-content: center;
        }

        .btn-download {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            width: 100%;
            padding: 13px;
            background: rgba(52,211,153,0.1);
            border: 1px solid rgba(52,211,153,0.35);
            border-radius: 11px;
            color: #34d399;
            font-size: 14px;
            font-weight: 800;
            text-decoration: none;
            transition: background 0.15s;
        }
        .btn-download:hover { background: rgba(52,211,153,0.2); color: white; }

        .error-box {
            background: rgba(239,68,68,0.07);
            border: 1px solid rgba(239,68,68,0.25);
            border-radius: 12px;
            padding: 16px;
        }

        .error-title { font-size: 13px; font-weight: 700; color: #f87171; margin-bottom: 8px; }
        .error-msg   { font-size: 12px; color: #fca5a5; line-height: 1.55; }

        /* ── Debug log ── */
        .debug-toggle {
            background: none;
            border: none;
            color: #4b5563;
            font-size: 11px;
            cursor: pointer;
            padding: 0;
            margin-top: 12px;
            font-family: inherit;
            text-decoration: underline;
        }
        .debug-toggle:hover { color: #9ca3af; }

        .debug-log {
            display: none;
            margin-top: 8px;
            background: rgba(0,0,0,0.3);
            border-radius: 8px;
            padding: 12px;
            font-family: monospace;
            font-size: 11px;
            color: #6b7280;
            white-space: pre-wrap;
            word-break: break-all;
            max-height: 200px;
            overflow-y: auto;
        }

        /* ── Orphan files ── */
        .orphan-table { width: 100%; border-collapse: collapse; font-size: 12px; }
        .orphan-table td { padding: 8px 10px; border-top: 1px solid rgba(255,255,255,0.05); color: #9ca3af; }
        .orphan-table td:first-child { font-family: monospace; color: #6b7280; font-size: 10px; }
        .orphan-table tr:hover td { background: rgba(255,255,255,0.02); }

        .btn-del {
            padding: 3px 9px;
            background: rgba(239,68,68,0.1);
            border: 1px solid rgba(239,68,68,0.2);
            border-radius: 6px;
            color: #f87171;
            font-size: 10px;
            font-weight: 700;
            cursor: pointer;
            font-family: inherit;
        }
        .btn-dl-sm {
            padding: 3px 9px;
            background: rgba(124,58,237,0.15);
            border: 1px solid rgba(124,58,237,0.3);
            border-radius: 6px;
            color: #a78bfa;
            font-size: 10px;
            font-weight: 700;
            text-decoration: none;
        }

        /* ── Loading overlay ── */
        .loading-overlay {
            display: none;
            position: fixed;
            inset: 0;
            background: rgba(10,8,30,0.85);
            z-index: 999;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            gap: 18px;
        }

        .loading-overlay.show { display: flex; }

        .spinner {
            width: 44px; height: 44px;
            border: 3px solid rgba(124,58,237,0.2);
            border-top-color: #7c3aed;
            border-radius: 50%;
            animation: spin 0.7s linear infinite;
        }

        @keyframes spin { to { transform: rotate(360deg); } }

        .loading-text { font-size: 14px; font-weight: 700; color: #a78bfa; }
        .loading-sub  { font-size: 12px; color: #6b7280; }
    </style>
</head>
<body>

<div class="loading-overlay" id="loadingOverlay">
    <div class="spinner"></div>
    <div class="loading-text">Memproses dokumen…</div>
    <div class="loading-sub">Mohon tunggu, ini bisa memakan waktu beberapa detik.</div>
</div>

<div class="top-bar">
    <a href="<?= BASE_PATH ?>/admin/dashboard.php" class="back-link">← Dashboard</a>
    <div>
        <div class="page-title">Test Dokumen</div>
        <div class="page-sub">Proses file tanpa order & payment — hasil langsung diunduh</div>
    </div>
</div>

<div class="layout">

    <!-- ── KIRI: FORM ── -->
    <div>
        <div class="card">
            <div class="card-title">Upload &amp; Konfigurasi</div>

            <form method="POST" enctype="multipart/form-data" id="testForm">
                <input type="hidden" name="paket" id="paketInput" value="paket3">

                <!-- Paket selector -->
                <div class="field">
                    <label>Paket</label>
                    <div class="paket-grid">
                        <div class="paket-btn" id="pb1" onclick="selectPaket('paket1',this)">
                            Paket 1
                            <div class="pb-label">Full Angka</div>
                        </div>
                        <div class="paket-btn" id="pb2" onclick="selectPaket('paket2',this)">
                            Paket 2
                            <div class="pb-label">Romawi + Angka</div>
                        </div>
                        <div class="paket-btn selected" id="pb3" onclick="selectPaket('paket3',this)">
                            Paket 3
                            <div class="pb-label">Populer Skripsi</div>
                        </div>
                    </div>
                </div>

                <!-- File -->
                <div class="field">
                    <label>File Dokumen (.docx)</label>
                    <input type="file" name="file" accept=".docx" required>
                </div>

                <!-- Font + Size -->
                <div class="field-row">
                    <div class="field">
                        <label>Font</label>
                        <select name="font">
                            <option value="Times New Roman">Times New Roman</option>
                            <option value="Arial">Arial</option>
                            <option value="Calibri">Calibri</option>
                            <option value="Georgia">Georgia</option>
                        </select>
                    </div>
                    <div class="field">
                        <label>Ukuran Font</label>
                        <select name="size">
                            <option value="10 pt">10 pt</option>
                            <option value="11 pt">11 pt</option>
                            <option value="12 pt" selected>12 pt</option>
                        </select>
                    </div>
                </div>

                <!-- Cover -->
                <div class="field">
                    <label>Cover</label>
                    <select name="hidden_cover">
                        <option value="Ya">Cover tanpa nomor</option>
                        <option value="Tidak">Cover tampilkan nomor</option>
                    </select>
                </div>

                <!-- Posisi (hanya paket 1 & 2) -->
                <div class="field" id="posisiField" style="display:none">
                    <label>Posisi Nomor</label>
                    <select name="posisi">
                        <option value="Tengah Bawah">Tengah Bawah</option>
                        <option value="Kanan Bawah">Kanan Bawah</option>
                        <option value="Kiri Bawah">Kiri Bawah</option>
                        <option value="Tengah Atas">Tengah Atas</option>
                        <option value="Kanan Atas">Kanan Atas</option>
                        <option value="Kiri Atas">Kiri Atas</option>
                    </select>
                </div>

                <button type="submit" class="btn-proses">⚡ Proses Sekarang</button>
            </form>
        </div>

        <?php if (!empty($orphans)): ?>
        <!-- Sisa file test -->
        <div class="card" style="margin-top:18px">
            <div class="card-title" style="margin-bottom:12px">File Test Tersimpan</div>
            <table class="orphan-table">
                <?php foreach ($orphans as $o): ?>
                <tr>
                    <td><?= htmlspecialchars(substr($o['key'], 0, 20)) ?>…</td>
                    <td><?= $o['size'] ?> KB</td>
                    <td><?= $o['age'] ?> menit lalu</td>
                    <td>
                        <a href="?dl=<?= urlencode($o['key']) ?>" class="btn-dl-sm">Unduh</a>
                    </td>
                    <td>
                        <form method="POST" style="display:inline"
                              onsubmit="return confirm('Hapus file ini?')">
                            <input type="hidden" name="del" value="<?= htmlspecialchars($o['key']) ?>">
                            <button type="submit" class="btn-del">Hapus</button>
                        </form>
                    </td>
                </tr>
                <?php endforeach; ?>
            </table>
        </div>
        <?php endif; ?>
    </div>

    <!-- ── KANAN: HASIL ── -->
    <div class="card">
        <div class="card-title">Hasil Proses</div>

        <?php if ($errMsg): ?>

        <div class="result-header">
            <span class="badge badge-err">Gagal</span>
        </div>
        <div class="error-box">
            <div class="error-title">Error</div>
            <div class="error-msg"><?= htmlspecialchars($errMsg) ?></div>
        </div>

        <?php elseif ($result): ?>

        <div class="result-header">
            <span class="badge badge-ok">Berhasil</span>
            <span style="font-size:12px;color:#6b7280">
                <?= htmlspecialchars(strtoupper($result['paket'])) ?> ·
                <?= $result['elapsed'] ?> detik ·
                <?= $result['outSize'] ?> KB
            </span>
        </div>

        <div class="result-stats">
            <div class="stat-mini">
                <div class="stat-mini-label">Sections</div>
                <div class="stat-mini-val"><?= $result['sections'] ?></div>
            </div>
            <div class="stat-mini">
                <div class="stat-mini-label">BAB Terdeteksi</div>
                <div class="stat-mini-val"><?= count($result['bab']) ?></div>
            </div>
            <div class="stat-mini">
                <div class="stat-mini-label">Waktu</div>
                <div class="stat-mini-val" style="font-size:15px"><?= $result['elapsed'] ?>s</div>
            </div>
        </div>

        <?php if (!empty($result['bab'])): ?>
        <div style="font-size:11px;color:#6b7280;font-weight:700;text-transform:uppercase;margin-bottom:8px;letter-spacing:.4px">
            BAB yang Terdeteksi
        </div>
        <ul class="bab-list">
            <?php foreach ($result['bab'] as $i => $b): ?>
            <li class="bab-item">
                <span class="bab-num"><?= $i + 1 ?></span>
                <?= htmlspecialchars($b) ?>
            </li>
            <?php endforeach; ?>
        </ul>
        <?php endif; ?>

        <a href="?dl=<?= urlencode($result['outKey']) ?>" class="btn-download">
            ⬇ Unduh Hasil (.docx)
        </a>

        <button class="debug-toggle" onclick="toggleDebug(this)">Lihat log Python ▾</button>
        <pre class="debug-log" id="debugLog"><?= htmlspecialchars($result['raw']) ?></pre>

        <?php else: ?>

        <div class="result-empty">
            <div class="result-icon">📄</div>
            <div>Belum ada hasil. Upload file dan klik <strong>Proses Sekarang</strong>.</div>
        </div>

        <?php endif; ?>
    </div>

</div>

<script>
function selectPaket(val, el) {
    document.querySelectorAll('.paket-btn').forEach(b => b.classList.remove('selected'));
    el.classList.add('selected');
    document.getElementById('paketInput').value = val;
    document.getElementById('posisiField').style.display =
        (val === 'paket1' || val === 'paket2') ? 'block' : 'none';
}

document.getElementById('testForm').addEventListener('submit', function() {
    document.getElementById('loadingOverlay').classList.add('show');
});

function toggleDebug(btn) {
    var el = document.getElementById('debugLog');
    var open = el.style.display === 'block';
    el.style.display = open ? 'none' : 'block';
    btn.textContent = open ? 'Lihat log Python ▾' : 'Sembunyikan log ▴';
}
</script>
</body>
</html>
