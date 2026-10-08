<?php
/**
 * Plugin Name: Daily Dispatch
 * Description: Export processed WooCommerce orders, shelving lists, and CityPak courier files by order status.
 * Version: 1.0.0
 * Author: Jumpbooks
 */

if (!defined('ABSPATH')) {
    exit;
}

define('JOP_PLUGIN_FILE', __FILE__);
define('JOP_PLUGIN_DIR', plugin_dir_path(__FILE__));
define('JOP_PLUGIN_URL', plugin_dir_url(__FILE__));

if (file_exists(JOP_PLUGIN_DIR . 'vendor/autoload.php')) {
    require_once JOP_PLUGIN_DIR . 'vendor/autoload.php';
}

require_once JOP_PLUGIN_DIR . 'includes/class-jumpbooks-order-processor.php';

add_action('plugins_loaded', function () {
    if (!class_exists('WooCommerce')) {
        add_action('admin_notices', function () {
            echo '<div class="notice notice-error"><p><strong>Daily Dispatch</strong> requires WooCommerce to be active.</p></div>';
        });
        return;
    }

    if (!class_exists('\PhpOffice\PhpSpreadsheet\Spreadsheet')) {
        add_action('admin_notices', function () {
            echo '<div class="notice notice-error"><p><strong>Daily Dispatch</strong> requires PhpSpreadsheet. Please run <code>composer install</code> before uploading the plugin.</p></div>';
        });
        return;
    }

    Jumpbooks_Order_Processor::instance();
});