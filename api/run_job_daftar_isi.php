<?php
/**
 * api/run_job_daftar_isi.php
 * Menjalankan Python daftar_isi.py secara synchronous via AJAX dari menunggu.php.
 */
date_default_timezone_set('Asia/Jakarta');
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';
require_once __DIR__ . '/../includes/log.php';

header('Content-Type: application/json');

set_time_limit(300);
ignore_user_abort(false);

$jobId = preg_replace('/[^a-zA-Z0-9_\-]/', '', $_GET['job']  ?? '');
$dbId  = (int)($_GET['db_id'] ?? 0);

if (!$jobId) {
    echo json_encode(['status' => 'error', 'message' => 'Job tidak valid.']);
    exit;
}

$logDir   = __DIR__ . '/../logs';
$metaFile = $logDir . '/job_' . $jobId . '_meta.json';

if (!file_exists($metaFile)) {
    echo json_encode(['status' => 'error', 'message' => 'Job tidak ditemukan.']);
    exit;
}

$meta        = json_decode(file_get_contents($metaFile), true) ?: [];
$input_full  = $meta['input_full']  ?? '';
$output_full = $meta['output_full'] ?? '';
$outputRel   = $meta['output_rel']  ?? '';
$dbIdFull    = $meta['db_id']       ?? $dbId;

if (!$input_full || !file_exists($input_full)) {
    echo json_encode(['status' => 'error', 'message' => 'File input tidak ditemukan.']);
    exit;
}

$hasilDir = dirname($output_full);
if (!is_dir($hasilDir)) mkdir($hasilDir, 0775, true);

$python = PYTHON_EXE;
$script = PYTHON_SCRIPT_DIR . '/daftar_isi.py';

$args = implode(' ', [
    escapeshellarg($input_full),
    escapeshellarg($output_full),
    escapeshellarg($meta['kedalaman']    ?? 'H1'),
    escapeshellarg($meta['format_titik'] ?? 'titik'),
    escapeshellarg($meta['font']         ?? 'Times New Roman'),
    escapeshellarg($meta['size']         ?? '12'),
    escapeshellarg($meta['space']        ?? '1.0'),
]);

$cmd = '"' . $python . '" "' . $script . '" ' . $args . ' 2>&1';

adk_log('daftar_isi', 'Job mulai', ['job' => $jobId, 'db_id' => $dbIdFull]);

exec($cmd, $out, $exitCode);
$rawOutput = implode('', $out);
$result    = @json_decode($rawOutput, true);

if ($exitCode !== 0 || !file_exists($output_full)) {
    $errCode = $result['code']    ?? 'PROCESSING_ERROR';
    $errMsg  = $result['message'] ?? $rawOutput;

    adk_log('error', 'Job daftar isi gagal', ['job' => $jobId, 'code' => $errCode, 'exit' => $exitCode, 'msg' => substr($errMsg, 0, 200)]);

    file_put_contents($logDir . '/error_daftar_isi.log',
        date('Y-m-d H:i:s') . " | job={$jobId} | order=" . ($meta['order_id'] ?? '?') . " | exit={$exitCode}\n" . $rawOutput . "\n---\n",
        FILE_APPEND | LOCK_EX
    );

    $pesanMap = [
        'FORMAT_NOT_SUPPORTED' => 'Format <b>.doc</b> tidak didukung. Buka di Microsoft Word → Simpan Sebagai → .docx, lalu upload ulang.',
        'FILE_TOO_LARGE'       => 'File terlalu besar. Coba hapus gambar yang tidak perlu.',
        'INVALID_DOCX'         => 'File rusak atau bukan Word yang valid. Buka di Word → Simpan ulang → upload lagi.',
        'MACRO_DETECTED'       => 'File mengandung macro. Simpan ulang sebagai .docx biasa di Word.',
        'FILE_READ_ERROR'      => 'File tidak bisa dibaca. Buka di Word → Simpan Sebagai .docx baru → upload lagi.',
        'NO_HEADINGS_FOUND'    => 'Tidak ditemukan heading di dokumen. Pastikan menggunakan style Heading 1/2/3 dari Word.',
        'PROCESSING_ERROR'     => 'Sistem kesulitan memproses file ini. Hubungi admin jika terus terjadi.',
        'FILE_SAVE_ERROR'      => 'Gagal menyimpan hasil. Coba ulangi.',
    ];
    $detail = $pesanMap[$errCode] ?? 'Gagal membuat daftar isi. Coba upload ulang atau hubungi admin.';

    echo json_encode(['status' => 'failed', 'message' => $detail]);
    exit;
}

if ($dbIdFull) {
    try {
        $db = getDB();
        $db->prepare("UPDATE daftar_isi_orders SET file_output = :out, status = 'selesai' WHERE id = :id")
           ->execute([':out' => $outputRel, ':id' => $dbIdFull]);
    } catch (Exception $e) {
        adk_log('error', 'Update file_output daftar isi gagal', ['job' => $jobId, 'err' => $e->getMessage()]);
    }
}

$_SESSION['di_file_output'] = $outputRel;

adk_log('daftar_isi', 'Job selesai', ['job' => $jobId, 'db_id' => $dbIdFull]);

echo json_encode(['status' => 'done']);