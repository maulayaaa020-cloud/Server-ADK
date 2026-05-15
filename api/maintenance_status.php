<?php
session_start();
header('Content-Type: application/json');
header('Cache-Control: no-store');
$isMaintenance = file_exists(__DIR__ . '/../config/maintenance.flag') && empty($_SESSION['adk_admin']);
echo json_encode(['maintenance' => $isMaintenance]);
