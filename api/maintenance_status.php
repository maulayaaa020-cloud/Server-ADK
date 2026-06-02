<?php
session_start();
header('Content-Type: application/json');
header('Cache-Control: no-store, no-cache, must-revalidate');
header('Pragma: no-cache');
header('Expires: 0');

$configDir = __DIR__ . '/../config/';
$isAdmin   = !empty($_SESSION['adk_admin']);

$serviceKeys = ['penomoran', 'daftar_isi', 'daftar_tabel', 'daftar_gambar', 'daftar_lampiran', 'lamaran_cv', 'ketik_surat'];
$services = [];
foreach ($serviceKeys as $key) {
    $flag = $configDir . 'maintenance_' . $key . '.flag';
    $services[$key] = file_exists($flag); // status nyata, tidak dipengaruhi admin
}

// backward compat
$isMaintenance = $services['penomoran'];

echo json_encode(['maintenance' => $isMaintenance, 'services' => $services, 'isAdmin' => $isAdmin]);
