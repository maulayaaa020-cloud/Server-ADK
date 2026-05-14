<?php
date_default_timezone_set('Asia/Jakarta');
require_once __DIR__ . '/_guard.php';

$error  = '';
$result = null;

// ── Download hasil test ───────────────────────────────────────────────────────
if (isset($_GET['dl'])) {
    $dl = ltrim(str_replace(['..', '\\'], ['', '/'], $_GET['dl']), '/');
    if (preg_match('#^hasil/test_[a-zA-Z0-9_.]+\.docx$#', $dl)) {
        $abs = __DIR__ . '/../' . $dl;
        if (file_exists($abs)) {
            header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            header('Content-Disposition: attachment; filename="' . basename($abs) . '"');
            header('Content-Length: ' . filesize($abs));
            readfile($abs);
            @unlink($abs);
            exit;
        }
    }
    http_response_code(404);
    exit('File tidak ditemukan.');
}

// ── Proses upload ─────────────────────────────────────────────────────────────
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['file'])) {
    $namaFile = $_FILES['file']['name'] ?? '';
    $tmpFile  = $_FILES['file']['tmp_name'] ?? '';
    $ext      = strtolower(pathinfo($namaFile, PATHINFO_EXTENSION));

    if (!in_array($ext, ['doc', 'docx'])) {
        $error = 'Format harus .doc atau .docx.';
    } elseif (($_FILES['file']['size'] ?? 0) > 30 * 1024 * 1024) {
        $error = 'File maksimal 30MB.';
    } else {
        $uploadDir = __DIR__ . '/../upload';
        if (!is_dir($uploadDir)) mkdir($uploadDir, 0775, true);

        $namaBaru  = 'test_' . time() . '_' . bin2hex(random_bytes(4)) . '.' . $ext;
        $inputFull = $uploadDir . '/' . $namaBaru;

        if (!move_uploaded_file($tmpFile, $inputFull)) {
            $error = 'Gagal menyimpan file upload.';
        } else {
            $hasilDir   = __DIR__ . '/../hasil';
            if (!is_dir($hasilDir)) mkdir($hasilDir, 0775, true);

            $outputName = 'test_' . date('YmdHis') . '_' . bin2hex(random_bytes(4)) . '.docx';
            $outputFull = $hasilDir . '/' . $outputName;

            $font   = in_array($_POST['font'] ?? '', ['Times New Roman','Arial','Calibri','Cambria'])
                        ? $_POST['font'] : 'Times New Roman';
            $size   = in_array($_POST['size'] ?? '', ['10 pt','11 pt','12 pt'])
                        ? $_POST['size'] : '12 pt';
            $hidden = ($_POST['hidden_cover'] ?? 'Ya') === 'Tidak' ? 'Tidak' : 'Ya';

            $python = PYTHON_EXE;
            $script = PYTHON_SCRIPT_DIR . '/main.py';
            $cmd = "\"$python\" \"$script\" " .
                escapeshellarg($inputFull)  . ' ' .
                escapeshellarg($outputFull) . ' ' .
                escapeshellarg('paket3')    . ' ' .
                escapeshellarg($font)       . ' ' .
                escapeshellarg($size)       . ' ' .
                escapeshellarg($hidden)     . ' ' .
                escapeshellarg('Tengah Bawah') . ' 2>&1';

            exec($cmd, $out, $status);
            @unlink($inputFull);

            if ($status !== 0 || !file_exists($outputFull)) {
                $detail = htmlspecialchars(implode("\n", $out));
                $error  = "Proses gagal (exit {$status}):\n{$detail}";
            } else {
                $jsonData = [];
                foreach (array_reverse($out) as $line) {
                    $dec = json_decode(trim($line), true);
                    if ($dec) { $jsonData = $dec; break; }
                }
                $result = [
                    'rel'      => 'hasil/' . $outputName,
                    'sections' => $jsonData['total_sections'] ?? '?',
                    'bab'      => $jsonData['detected_bab']   ?? [],
                    'size'     => round(filesize($outputFull) / 1024, 1),
                ];
            }
        }
    }
}
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
    color: #e5e7eb;
    padding: 32px 24px;
}
.container { max-width: 560px; margin: 0 auto; }
h1  { font-size: 20px; font-weight: 800; color: white; margin-bottom: 4px; }
.sub { font-size: 13px; color: #6b7280; margin-bottom: 28px; }
a.back { color: #a78bfa; font-size: 13px; text-decoration: none; display: inline-block; margin-bottom: 20px; }
a.back:hover { color: white; }

.card {
    background: rgba(255,255,255,0.05);
    border: 1px solid rgba(124,58,237,0.25);
    border-radius: 16px;
    padding: 28px;
}

label {
    display: block;
    font-size: 11px;
    font-weight: 700;
    color: #9ca3af;
    text-transform: uppercase;
    letter-spacing: .5px;
    margin-bottom: 7px;
}

.field { margin-bottom: 18px; }

input[type=file],
select {
    width: 100%;
    padding: 10px 14px;
    background: rgba(255,255,255,0.07);
    border: 1px solid rgba(124,58,237,0.3);
    border-radius: 10px;
    color: white;
    font-size: 14px;
    outline: none;
    transition: border-color .2s;
}
input[type=file]:focus,
select:focus { border-color: #7c3aed; }
select option { background: #1a1043; }

.row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }

.badge-paket {
    display: inline-block;
    padding: 3px 10px;
    background: rgba(124,58,237,.2);
    border: 1px solid rgba(124,58,237,.4);
    border-radius: 20px;
    font-size: 11px;
    font-weight: 700;
    color: #a78bfa;
    margin-bottom: 20px;
}

.btn-submit {
    width: 100%;
    padding: 12px;
    background: #7c3aed;
    border: none;
    border-radius: 12px;
    color: white;
    font-size: 15px;
    font-weight: 700;
    cursor: pointer;
    transition: background .2s;
    margin-top: 4px;
}
.btn-submit:hover { background: #6d28d9; }
.btn-submit:disabled { background: #4b5563; cursor: not-allowed; }

.alert-err {
    background: rgba(239,68,68,.1);
    border: 1px solid rgba(239,68,68,.3);
    border-radius: 10px;
    padding: 14px 16px;
    font-size: 13px;
    color: #f87171;
    margin-bottom: 20px;
    white-space: pre-wrap;
    word-break: break-all;
}

.result-card {
    background: rgba(52,211,153,.07);
    border: 1px solid rgba(52,211,153,.25);
    border-radius: 14px;
    padding: 22px 24px;
    margin-bottom: 20px;
}
.result-title { font-size: 15px; font-weight: 700; color: #34d399; margin-bottom: 14px; }
.result-row   { display: flex; justify-content: space-between; font-size: 13px; color: #9ca3af; margin-bottom: 7px; }
.result-row span:last-child { color: white; font-weight: 600; }

.bab-list { margin-top: 10px; }
.bab-item {
    font-size: 12px;
    color: #a78bfa;
    background: rgba(124,58,237,.1);
    border-radius: 6px;
    padding: 3px 9px;
    display: inline-block;
    margin: 3px 3px 0 0;
}

.btn-dl {
    display: block;
    width: 100%;
    padding: 13px;
    background: linear-gradient(135deg, #059669, #047857);
    border: none;
    border-radius: 12px;
    color: white;
    font-size: 15px;
    font-weight: 700;
    text-align: center;
    text-decoration: none;
    transition: opacity .2s;
}
.btn-dl:hover { opacity: .85; }

.spinner {
    display: none;
    text-align: center;
    padding: 12px 0;
    font-size: 13px;
    color: #9ca3af;
}
</style>
</head>
<body>
<div class="container">

    <a href="dashboard.php" class="back">← Dashboard</a>
    <h1>Test Dokumen</h1>
    <p class="sub">Upload file .docx, proses paket 3, langsung unduh — tanpa payment.</p>

    <?php if ($error): ?>
    <div class="alert-err"><?= $error ?></div>
    <?php endif; ?>

    <?php if ($result): ?>
    <div class="result-card">
        <div class="result-title">✅ Berhasil diproses</div>
        <div class="result-row"><span>Total section</span><span><?= $result['sections'] ?></span></div>
        <div class="result-row"><span>Ukuran output</span><span><?= $result['size'] ?> KB</span></div>
        <?php if ($result['bab']): ?>
        <div class="result-row" style="flex-direction:column;gap:6px">
            <span>BAB terdeteksi</span>
            <div class="bab-list">
                <?php foreach ($result['bab'] as $b): ?>
                <span class="bab-item"><?= htmlspecialchars($b) ?></span>
                <?php endforeach; ?>
            </div>
        </div>
        <?php endif; ?>
    </div>
    <a href="test.php?dl=<?= urlencode($result['rel']) ?>" class="btn-dl">⬇ Unduh Hasil</a>
    <p style="font-size:11px;color:#6b7280;text-align:center;margin-top:10px">
        File otomatis dihapus setelah diunduh.
    </p>
    <?php else: ?>

    <div class="card">
        <div class="badge-paket">Paket 3 — Romawi + Angka (Skripsi)</div>

        <form method="POST" enctype="multipart/form-data"
              onsubmit="this.querySelector('.btn-submit').disabled=true;
                        this.querySelector('.spinner').style.display='block';">

            <div class="field">
                <label>File Dokumen (.docx)</label>
                <input type="file" name="file" accept=".doc,.docx" required>
            </div>

            <div class="row2">
                <div class="field">
                    <label>Font</label>
                    <select name="font">
                        <option value="Times New Roman">Times New Roman</option>
                        <option value="Arial">Arial</option>
                        <option value="Calibri">Calibri</option>
                        <option value="Cambria">Cambria</option>
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

            <div class="field">
                <label>Cover</label>
                <select name="hidden_cover">
                    <option value="Ya">Cover tanpa nomor</option>
                    <option value="Tidak">Cover tampil nomor romawi</option>
                </select>
            </div>

            <button type="submit" class="btn-submit">Proses Sekarang</button>
            <div class="spinner">⏳ Memproses dokumen, harap tunggu…</div>
        </form>
    </div>

    <?php endif; ?>
</div>
</body>
</html>
