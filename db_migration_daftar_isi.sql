-- Migration: Buat tabel daftar_isi_orders
-- Jalankan sekali di phpMyAdmin atau mysql CLI

CREATE TABLE IF NOT EXISTS `daftar_isi_orders` (
    `id`           INT          NOT NULL AUTO_INCREMENT,
    `order_id`     VARCHAR(64)  NOT NULL UNIQUE,
    `phone`        VARCHAR(20)  DEFAULT NULL,
    `guest_token`  VARCHAR(64)  DEFAULT NULL,
    `file_input`   VARCHAR(255) NOT NULL,
    `file_output`  VARCHAR(255) DEFAULT NULL,
    `kedalaman`    VARCHAR(20)  NOT NULL,
    `format_titik` VARCHAR(20)  NOT NULL,
    `harga`        INT          NOT NULL DEFAULT 10000,
    `status`       VARCHAR(20)  NOT NULL DEFAULT 'pending',
    `created_at`   DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (`id`),
    INDEX `idx_order_id`    (`order_id`),
    INDEX `idx_phone`       (`phone`),
    INDEX `idx_guest_token` (`guest_token`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
