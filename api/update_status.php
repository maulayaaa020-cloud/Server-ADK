<?php
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

header('Content-Type: application/json');
$data = json_decode(file_get_contents('php://input'), true);
$dbId = (int)($data['db_id'] ?? 0);

if (!$dbId) { echo json_encode(['ok' => false]); exit; }

try {
    $db = getDB();
    $db->prepare("UPDATE orders SET status = 'paid' WHERE id = :id")->execute([':id' => $dbId]);
    echo json_encode(['ok' => true]);
} catch (Exception $e) {
    echo json_encode(['ok' => false, 'error' => $e->getMessage()]);
}
