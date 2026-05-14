<?php
session_start();
require_once __DIR__ . '/../includes/config.php';
session_unset();
session_destroy();
header('Location: ' . BASE_PATH . '/index.html');
exit;
