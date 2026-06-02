<?php
date_default_timezone_set('Asia/Jakarta');
session_start();
require_once __DIR__ . '/includes/config.php';

// Validasi parameter
$jobId   = preg_replace('/[^a-zA-Z0-9_\-]/', '', $_GET['job']  ?? '');
$dbId    = (int)($_GET['db_id'] ?? $_SESSION['order_db_id'] ?? 0);
$jobType = ($_GET['type'] ?? '') === 'daftar_isi' ? 'daftar_isi' : 'penomoran';

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
    <script>(function(){var t=localStorage.getItem('adkTheme')||'light';document.documentElement.setAttribute('data-theme',t);})();</script>
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
        h2 { font-size: 1.5rem; color: #e5e7eb; margin-bottom: 10px; }
        p  { color: #9ca3af; font-size: 1.05rem; line-height: 1.6; }
        .step {
            margin-top: 24px;
            font-size: 1rem;
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
            font-size: 1rem;
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
            font-size: 1rem;
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
            font-size: 1rem;
            margin-bottom: 12px;
        }

        html[data-theme="light"] body {
            background: linear-gradient(135deg, #f0f7ff, #e1f0fb);
            color: #1a2332;
        }
        html[data-theme="light"] .card {
            background: #ffffff;
            border-color: #e2e8f0;
            box-shadow: 0 4px 24px rgba(0,0,0,0.06);
        }
        html[data-theme="light"] .spinner {
            border-color: rgba(99,102,241,0.15);
            border-top-color: #6d28d9;
        }
        html[data-theme="light"] h2 { color: #1a2332; }
        html[data-theme="light"] p  { color: #64748b; }
        html[data-theme="light"] .step { color: #64748b; }
        html[data-theme="light"] .step span { color: #6d28d9; }
        html[data-theme="light"] .err-msg {
            background: rgba(239,68,68,0.07);
            border-color: rgba(239,68,68,0.25);
            color: #dc2626;
        }
        html[data-theme="light"] .btn-back {
            background: rgba(99,102,241,0.08);
            border-color: rgba(99,102,241,0.25);
            color: #6d28d9;
        }
        html[data-theme="light"] .btn-back:hover { background: rgba(99,102,241,0.15); }
        html[data-theme="light"] .queue-badge {
            background: rgba(245,158,11,0.1);
            border-color: rgba(245,158,11,0.25);
            color: #b45309;
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
const JOB_TYPE = <?= json_encode($jobType) ?>;
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

const endpoint = JOB_TYPE === 'daftar_isi' ? '/api/run_job_daftar_isi.php' : '/api/run_job.php';
fetch(BASE + endpoint + '?job=' + encodeURIComponent(JOB) + '&db_id=' + DB_ID, {
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
    <script src="theme.js?v=1"></script>
</body>
</html>