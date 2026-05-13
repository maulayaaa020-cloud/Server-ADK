<?php
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

header('Content-Type: application/json');
$data  = json_decode(file_get_contents('php://input'), true);
$phone = trim($data['phone'] ?? '');
$token = trim($data['token'] ?? '');

if (!$phone || !$token) { echo json_encode(['ok' => false]); exit; }

try {
    $db   = getDB();
    $stmt = $db->prepare(
        "SELECT id FROM trusted_devices
         WHERE phone = :p AND token = :t AND expires_at > NOW() LIMIT 1"
    );
    $stmt->execute([':p' => $phone, ':t' => $token]);
    $row = $stmt->fetch();

    if ($row) {
        $db->prepare("UPDATE trusted_devices SET expires_at = NOW() + INTERVAL 30 DAY WHERE id = :id")
           ->execute([':id' => $row['id']]);
        echo json_encode(['ok' => true]);
    } else {
        echo json_encode(['ok' => false]);
    }

} catch (Exception $e) {
    echo json_encode(['ok' => false]);
}
