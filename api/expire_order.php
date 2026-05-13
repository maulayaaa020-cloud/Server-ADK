<?php
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

header('Content-Type: application/json');
$data = json_decode(file_get_contents('php://input'), true);
$id   = (int)($data['id'] ?? 0);

if (!$id) { echo json_encode(['ok' => false]); exit; }

try {
    $db  = getDB();
    $row = $db->prepare("SELECT status, file_output FROM orders WHERE id = :id");
    $row->execute([':id' => $id]);
    $order = $row->fetch();

    if (!$order || $order['status'] !== 'pending') {
        echo json_encode(['ok' => false, 'reason' => 'not_pending']);
        exit;
    }

    if ($order['file_output']) {
        $path = __DIR__ . '/../' . $order['file_output'];
        if (file_exists($path)) unlink($path);
    }

    $db->prepare("UPDATE orders SET status = 'failed', file_output = NULL WHERE id = :id")
       ->execute([':id' => $id]);

    echo json_encode(['ok' => true]);
} catch (Exception $e) {
    echo json_encode(['ok' => false, 'error' => $e->getMessage()]);
}
