<?php
date_default_timezone_set('Asia/Jakarta');
session_start();
require_once __DIR__ . '/includes/config.php';

// Validasi parameter
$jobId = preg_replace('/[^a-zA-Z0-9_\-]/', '', $_GET['job']   ?? '');
$dbId  = (int)($_GET['db_id'] ?? $_SESSION['order_db_id'] ?? 0);

if (!$jobId) {
    header("Location: jasa.html");
    exit;
}
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Memproses File – ADK PHOTOCOPY</title>
    <link rel="icon" type="image/png" href="favicon.png">
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            background: #0d0b1e;
            color: white;
            font-family: 'Segoe UI', sans-serif;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .card {
            text-align: center;
            padding: 48px 40px;
            background: rgba(255,255,255,0.04);
            border: 1px solid rgba(124,58,237,0.2);
            border-radius: 20px;
            max-width: 460px;
            width: 90%;
        }
        .spinner {
            width: 56px; height: 56px;
            border: 4px solid rgba(167,139,250,0.2);
            border-top-color: #a78bfa;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
            margin: 0 auto 24px;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
        h2 { font-size: 1.3rem; color: #e5e7eb; margin-bottom: 10px; }
        p  { color: #9ca3af; font-size: 0.95rem; line-height: 1.6; }
        .step {
            margin-top: 24px;
            font-size: 0.88rem;
            color: #6b7280;
        }
        .step span { color: #a78bfa; font-weight: 600; }

        /* Error state */
        .err-icon { font-size: 2.8rem; margin-bottom: 16px; }
        .err-msg  {
            background: rgba(239,68,68,0.1);
            border: 1px solid rgba(239,68,68,0.3);
            border-radius: 10px;
            padding: 14px 16px;
            color: #fca5a5;
            font-size: 0.9rem;
            line-height: 1.6;
            margin: 16px 0;
        }
        .btn-back {
            display: inline-block;
            margin-top: 16px;
            padding: 10px 24px;
            background: rgba(124,58,237,0.2);
            border: 1px solid rgba(124,58,237,0.4);
            border-radius: 10px;
            color: #a78bfa;
            text-decoration: none;
            font-size: 0.9rem;
            transition: 0.2s;
        }
        .btn-back:hover { background: rgba(124,58,237,0.35); }

        /* Queue state */
        .queue-badge {
            display: inline-block;
            background: rgba(251,191,36,0.15);
            border: 1px solid rgba(251,191,36,0.3);
            color: #fbbf24;
            border-radius: 8px;
            padding: 6px 16px;
            font-size: 0.85rem;
            margin-bottom: 12px;
        }
    </style>
</head>
<body>
<div class="card" id="card">
    <div class="spinner" id="spinner"></div>
    <h2 id="title">Sedang memproses file kamu...</h2>
    <p id="desc">Mohon tunggu sebentar, jangan tutup halaman ini.</p>
    <div class="step" id="step">Estimasi: <span>10–30 detik</span></div>
</div>

<script>
const JOB      = <?= json_encode($jobId) ?>;
const DB_ID    = <?= (int)$dbId ?>;
const BASE     = <?= json_encode(rtrim(defined('BASE_PATH') ? BASE_PATH : '', '/')) ?>;
// Timer penghitung waktu (tampilan saja)
const startTime = Date.now();
const timerEl   = document.getElementById('step');
setInterval(() => {
    const secs = Math.round((Date.now() - startTime) / 1000);
    timerEl.innerHTML = 'Sudah menunggu: <span>' + secs + ' detik</span>';
}, 1000);

function showError(msg) {
    document.getElementById('card').innerHTML =
        '<div class="err-icon">&#9888;&#65039;</div>' +
        '<h2 style="color:#f87171">Gagal memproses file</h2>' +
        '<div class="err-msg">' + msg + '</div>' +
        '<a href="' + BASE + '/jasa.html" class="btn-back">&larr; Upload ulang</a>';
}

// Satu AJAX call ke run_job.php — PHP jalankan Python, browser tunggu hasilnya.
fetch(BASE + '/api/run_job.php?job=' + encodeURIComponent(JOB) + '&db_id=' + DB_ID, {
    method: 'GET',
    signal: AbortSignal.timeout ? AbortSignal.timeout(300000) : undefined
})
.then(r => {
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return r.json();
})
.then(data => {
    if (data.status === 'done') {
        document.getElementById('spinner').style.borderTopColor = '#34d399';
        document.getElementById('title').textContent = 'Selesai!';
        document.getElementById('desc').textContent  = 'File kamu sudah siap diunduh.';
        timerEl.innerHTML = '';
        setTimeout(() => { window.location.href = BASE + '/history.php'; }, 600);
    } else {
        showError(data.message || 'Gagal memproses file. Coba upload ulang.');
    }
})
.catch(err => {
    showError('Koneksi terputus atau proses timeout. Cek history — mungkin file sudah selesai. Jika belum, coba upload ulang.');
});
</script>
</body>
</html>