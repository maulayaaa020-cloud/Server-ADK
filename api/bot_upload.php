<?php
require_once __DIR__ . '/../includes/config.php';

header('Content-Type: application/json');
header('Access-Control-Allow-Origin: *');

$apiKey = $_SERVER['HTTP_X_BOT_KEY'] ?? '';
if ($apiKey !== 'adkivia-bot-2026') {
    http_response_code(401);
    echo json_encode(['error' => 'Unauthorized']);
    exit;
}

$uploadDir = __DIR__ . '/../uploads/bot/';
if (!is_dir($uploadDir)) {
    mkdir($uploadDir, 0755, true);
}

if (empty($_FILES['file'])) {
    http_response_code(400);
    echo json_encode(['error' => 'File tidak ditemukan']);
    exit;
}

$file     = $_FILES['file'];
$origName = basename($file['name']);
$ext      = strtolower(pathinfo($origName, PATHINFO_EXTENSION));

if (!in_array($ext, ['docx', 'doc'])) {
    http_response_code(400);
    echo json_encode(['error' => 'Hanya file .docx/.doc yang diterima']);
    exit;
}

$uploadId = uniqid('bot_', true);
$filename  = $uploadId . '.' . $ext;
$destPath  = $uploadDir . $filename;

if (!move_uploaded_file($file['tmp_name'], $destPath)) {
    http_response_code(500);
    echo json_encode(['error' => 'Gagal menyimpan file']);
    exit;
}

echo json_encode([
    'success'    => true,
    'upload_id'  => $uploadId,
    'filename'   => $filename,
    'orig_name'  => $origName,
    'size'       => $file['size'],
]);
