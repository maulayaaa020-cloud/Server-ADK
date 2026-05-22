<?php
/**
 * Centralized logging helper.
 * Semua event penting (order, processing, error, payment) dicatat di sini.
 */

function adk_log(string $category, string $message, array $context = []): void
{
    $logDir = dirname(__DIR__) . '/logs';
    if (!is_dir($logDir)) {
        @mkdir($logDir, 0775, true);
    }

    $ts   = date('Y-m-d H:i:s');
    $ctx  = !empty($context) ? ' | ' . json_encode($context, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES) : '';
    $line = "[$ts] $message$ctx" . PHP_EOL;

    $file = $logDir . '/' . preg_replace('/[^a-z0-9_]/', '', $category) . '.log';
    @file_put_contents($file, $line, FILE_APPEND | LOCK_EX);
}