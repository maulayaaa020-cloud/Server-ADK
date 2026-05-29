<?php
ob_start();
ini_set('display_errors', 0);
session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

$isAdmin    = !empty($_SESSION['adk_admin']);
$email      = $_SESSION['email'] ?? $_SESSION['cek_email'] ?? null;
$guestToken = $_SESSION['guest_token'] ?? $_COOKIE['adk_guest'] ?? null;

if (!$isAdmin && !$email && !$guestToken) {
    ob_end_clean(); http_response_code(403); exit;
}

$db_id = (int)($_GET['id']   ?? 0);
$page  = (int)($_GET['page'] ?? 1);
if (!$db_id || $page < 1) {
    ob_end_clean(); http_response_code(400); exit;
}

try {
    $db = getDB();
    if ($isAdmin) {
        $stmt = $db->prepare("SELECT id FROM orders WHERE id = ?");
        $stmt->execute([$db_id]);
    } elseif ($email) {
        $stmt = $db->prepare("SELECT id FROM orders WHERE id = ? AND phone = ?");
        $stmt->execute([$db_id, $email]);
    } else {
        $stmt = $db->prepare("SELECT id FROM orders WHERE id = ? AND guest_token = ?");
        $stmt->execute([$db_id, $guestToken]);
    }
    if (!$stmt->fetch()) {
        ob_end_clean(); http_response_code(403); exit;
    }
} catch (Exception $e) {
    ob_end_clean(); http_response_code(500); exit;
}

$cacheDir = __DIR__ . '/../preview_cache/' . $db_id;
$pages    = glob($cacheDir . '/page-*.png');
if (!is_array($pages)) $pages = [];
natsort($pages);
$pages = array_values($pages);

if (!isset($pages[$page - 1]) || !file_exists($pages[$page - 1])) {
    ob_end_clean(); http_response_code(404); exit;
}

$imgPath = $pages[$page - 1];
ob_end_clean();
header('Content-Type: image/png');
header('Content-Length: ' . filesize($imgPath));
header('Cache-Control: private, no-store, no-cache');
header('X-Robots-Tag: noindex');
readfile($imgPath);
