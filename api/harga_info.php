<?php
header('Content-Type: application/json');
header('Cache-Control: no-store');
$config = @json_decode(@file_get_contents(__DIR__ . '/../config/harga.json'), true);
echo json_encode($config ?: ['paket1' => 5000, 'paket2' => 10000, 'paket3' => 10000]);
