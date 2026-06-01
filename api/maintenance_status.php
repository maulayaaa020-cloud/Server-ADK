<?php
session_start();
header('Content-Type: application/json');
header('Cache-Control: no-store');

$configDir = __DIR__ . '/../config/';
$isAdmin   = !empty($_SESSION['adk_admin']);

$serviceKeys = ['penomoran', 'daftar_isi', 'daftar_tabel', 'daftar_gambar', 'daftar_lampiran', 'ketik'];
$services = [];
foreach ($serviceKeys as $key) {
    $flag = $configDir . 'maintenance_' . $key . '.flag';
    $services[$key] = file_exists($flag) && !$isAdmin;
}

// backward compat: jasa.html checks d.maintenance for penomoran service
$isMaintenance = $services['penomoran'];

echo json_encode(['maintenance' => $isMaintenance, 'services' => $services]);
