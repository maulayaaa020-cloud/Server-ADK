<?php
/**
 * api/job_status.php — Polling endpoint untuk cek status pemrosesan file.
 * Dipanggil oleh menunggu.php setiap 2 detik.
 *
 * GET params:
 *   job   = job ID (string, alphanumeric + dash/underscore)
 *   db_id = order DB id (int) — untuk update file_output saat selesai
 */
date_default_timezone_set('Asia/Jakarta');
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';
require_once __DIR__ . '/../includes/log.php';

header('Content-Type: application/json');

$jobId  = preg_replace('/[^a-zA-Z0-9_\-]/', '', $_GET['job']   ?? '');
$dbId   = (int)($_GET['db_id'] ?? 0);

if (!$jobId) {
    echo json_encode(['status' => 'error', 'message' => 'Job ID tidak valid.']);
    exit;
}

$logDir  = dirname(__DIR__) . '/logs';
$jobFile = $logDir . '/job_' . $jobId . '.json';

// Belum selesai — file belum ada
if (!file_exists($jobFile)) {
    echo json_encode(['status' => 'processing']);
    exit;
}

// Baca hasil dari Python
$raw    = file_get_contents($jobFile);
$result = @json_decode($raw, true);

// File ada tapi belum lengkap (Python masih nulis)
if (!$result || !isset($result['status'])) {
    echo json_encode(['status' => 'processing']);
    exit;
}

// ── Selesai ──────────────────────────────────────────────────────────────────
if ($result['status'] === 'success') {
    $outputRel = $_SESSION['file_output'] ?? null;

    // Coba baca output path dari job file (disimpan sebelum eksekusi)
    $metaFile = $logDir . '/job_' . $jobId . '_meta.json';
    if (file_exists($metaFile)) {
        $meta      = @json_decode(file_get_contents($metaFile), true) ?: [];
        $outputRel = $meta['output_rel'] ?? $outputRel;
        $dbIdFull  = $meta['db_id']      ?? $dbId;
    } else {
        $dbIdFull = $dbId;
    }

    // Update DB: tandai file_output dan status
    if ($dbIdFull) {
        try {
            $db = getDB();
            if ($outputRel) {
                $db->prepare("UPDATE orders SET file_output = :out WHERE id = :id AND file_output IS NULL")
                   ->execute([':out' => $outputRel, ':id' => $dbIdFull]);
            }
        } catch (Exception $e) {
            adk_log('error', 'job_status DB update gagal', [
                'job' => $jobId, 'db_id' => $dbIdFull, 'err' => $e->getMessage()
            ]);
        }
    }

    // Bersihkan file job agar tidak menumpuk (simpan 24 jam, hapus via cron atau manual)
    // @unlink($jobFile); // dinonaktifkan — berguna untuk debug

    adk_log('processing', 'Job selesai', ['job' => $jobId, 'db_id' => $dbIdFull]);
    echo json_encode(['status' => 'done']);
    exit;
}

// ── Gagal ─────────────────────────────────────────────────────────────────────
$errCode = $result['code']    ?? '';
$errMsg  = $result['message'] ?? 'Terjadi kesalahan.';

adk_log('error', 'Job gagal', ['job' => $jobId, 'code' => $errCode, 'msg' => $errMsg]);

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

echo json_encode(['status' => 'failed', 'code' => $errCode, 'message' => $detail]);