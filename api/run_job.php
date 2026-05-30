<?php
/**
 * api/run_job.php
 * Menjalankan Python secara synchronous, dipanggil via AJAX dari menunggu.php.
 * User melihat spinner, browser menunggu response AJAX ini (bukan page load).
 */
date_default_timezone_set('Asia/Jakarta');
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';
require_once __DIR__ . '/../includes/log.php';

header('Content-Type: application/json');

// Beri waktu hingga 5 menit untuk file besar
set_time_limit(300);
ignore_user_abort(false);

$jobId = preg_replace('/[^a-zA-Z0-9_\-]/', '', $_GET['job']   ?? '');
$dbId  = (int)($_GET['db_id'] ?? 0);

if (!$jobId) {
    echo json_encode(['status' => 'error', 'message' => 'Job tidak valid.']);
    exit;
}

// Baca metadata job
$logDir   = __DIR__ . '/../logs';
$metaFile = $logDir . '/job_' . $jobId . '_meta.json';

if (!file_exists($metaFile)) {
    echo json_encode(['status' => 'error', 'message' => 'Job tidak ditemukan.']);
    exit;
}

$meta       = json_decode(file_get_contents($metaFile), true) ?: [];
$input_full = $meta['input_full']  ?? '';
$output_full= $meta['output_full'] ?? '';
$outputRel  = $meta['output_rel']  ?? '';
$dbIdFull   = $meta['db_id']       ?? $dbId;

if (!$input_full || !file_exists($input_full)) {
    echo json_encode(['status' => 'error', 'message' => 'File input tidak ditemukan.']);
    exit;
}

// ── Buat folder hasil jika belum ada ──────────────────────────────────────────
$hasilDir = dirname($output_full);
if (!is_dir($hasilDir)) mkdir($hasilDir, 0775, true);

// ── Build command ─────────────────────────────────────────────────────────────
$python = PYTHON_EXE;
$script = PYTHON_SCRIPT_DIR . '/main.py';

$args = implode(' ', [
    escapeshellarg($input_full),
    escapeshellarg($output_full),
    escapeshellarg($meta['paket']       ?? 'paket3'),
    escapeshellarg($meta['font']        ?? 'Times New Roman'),
    escapeshellarg($meta['size']        ?? '12 pt'),
    escapeshellarg($meta['hidden']      ?? 'Ya'),
    escapeshellarg($meta['posisi']      ?? 'Tengah Bawah'),
    escapeshellarg($meta['pos_bab']     ?? 'Tengah Bawah'),
    escapeshellarg($meta['pos_isi']     ?? 'Kanan Atas'),
    escapeshellarg($meta['dimulai']     ?? 'i'),
    escapeshellarg($meta['semb_dafus']  ?? 'Tidak'),
    escapeshellarg($meta['semb_lamprn'] ?? 'Tidak'),
    escapeshellarg((string)($meta['num_cover'] ?? 1)),
]);

$cmd = '"' . $python . '" "' . $script . '" ' . $args . ' 2>&1';

adk_log('processing', 'Job mulai', ['job' => $jobId, 'db_id' => $dbIdFull]);

// ── Jalankan Python (synchronous — browser menunggu via AJAX) ─────────────────
exec($cmd, $out, $exitCode);
$rawOutput = implode('', $out);
$result    = @json_decode($rawOutput, true);

// ── Tangani hasil ──────────────────────────────────────────────────────────────
if ($exitCode !== 0 || !file_exists($output_full)) {
    $errCode = $result['code']    ?? 'PROCESSING_ERROR';
    $errMsg  = $result['message'] ?? $rawOutput;

    adk_log('error', 'Job gagal', ['job' => $jobId, 'code' => $errCode, 'exit' => $exitCode, 'msg' => substr($errMsg, 0, 200)]);

    // Simpan log error detail
    $errorLogDir = $logDir;
    if (!is_dir($errorLogDir)) mkdir($errorLogDir, 0775, true);
    file_put_contents($errorLogDir . '/error.log',
        date('Y-m-d H:i:s') . " | job={$jobId} | order=" . ($meta['order_id'] ?? '?') . " | exit={$exitCode}\n" . $rawOutput . "\n---\n",
        FILE_APPEND | LOCK_EX
    );

    // Jika Python sempat menulis file sebelum crash, simpan ke DB supaya admin bisa download.
    // Status TIDAK diubah — user yang sudah bayar tetap berstatus 'paid'.
    if ($dbIdFull && $outputRel && file_exists($output_full)) {
        try {
            $db = getDB();
            $db->prepare(
                "UPDATE orders SET file_output = COALESCE(file_output, :out) WHERE id = :id"
            )->execute([':out' => $outputRel, ':id' => $dbIdFull]);
        } catch (Exception $e) {
            adk_log('error', 'Simpan partial file_output gagal', ['job' => $jobId, 'err' => $e->getMessage()]);
        }
    }

    $pesanMap = [
        'FORMAT_NOT_SUPPORTED' => 'Format <b>.doc</b> tidak didukung. Buka di Microsoft Word → Simpan Sebagai → .docx, lalu upload ulang.',
        'FILE_TOO_LARGE'       => 'File terlalu besar. Coba hapus gambar yang tidak perlu atau kompres dulu.',
        'INVALID_DOCX'         => 'File rusak atau bukan Word yang valid. Buka di Word → Simpan ulang → upload lagi.',
        'MACRO_DETECTED'       => 'File mengandung macro. Simpan ulang sebagai .docx biasa di Word.',
        'FILE_READ_ERROR'      => 'File tidak bisa dibaca. Buka di Word → Simpan Sebagai .docx baru → upload lagi.',
        'PROCESSING_ERROR'     => 'Sistem kesulitan memproses file ini. Hubungi admin jika terus terjadi.',
        'FILE_SAVE_ERROR'      => 'Gagal menyimpan hasil. Coba ulangi.',
    ];
    $detail = $pesanMap[$errCode] ?? 'Gagal memproses dokumen. Coba upload ulang atau hubungi admin.';

    echo json_encode(['status' => 'failed', 'message' => $detail]);
    exit;
}

// ── Sukses: update DB ─────────────────────────────────────────────────────────
if ($dbIdFull) {
    try {
        $db = getDB();
        $db->prepare("UPDATE orders SET file_output = :out WHERE id = :id AND file_output IS NULL")
           ->execute([':out' => $outputRel, ':id' => $dbIdFull]);
    } catch (Exception $e) {
        adk_log('error', 'Update file_output gagal', ['job' => $jobId, 'err' => $e->getMessage()]);
    }
}

// Update session
$_SESSION['file_output'] = $outputRel;

adk_log('processing', 'Job selesai', ['job' => $jobId, 'db_id' => $dbIdFull]);

echo json_encode(['status' => 'done']);