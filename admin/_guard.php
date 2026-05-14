<?php
if (session_status() === PHP_SESSION_NONE) session_start();
require_once __DIR__ . '/../includes/config.php';
require_once __DIR__ . '/../includes/db.php';

if (empty($_SESSION['adk_admin'])) {
    header('Location: ' . BASE_PATH . '/admin/login.php');
    exit;
}
