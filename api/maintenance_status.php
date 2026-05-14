<?php
header('Content-Type: application/json');
header('Cache-Control: no-store');
echo json_encode(['maintenance' => file_exists(__DIR__ . '/../config/maintenance.flag')]);
