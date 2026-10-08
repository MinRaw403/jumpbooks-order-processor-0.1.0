<?php

/*
 * ============================================================================
 * DAILY DISPATCH - HUMAN-FRIENDLY MAP
 * ============================================================================
 *
 * What this plugin does:
 *
 *   WooCommerce orders
 *          |
 *          v
 *   Daily Dispatch
 *          |
 *          +--> Orders sheet
 *          +--> Processed Orders sheet
 *          +--> Shelving List
 *          +--> Shortened Orders reference sheet
 *          +--> CityPak courier export
 *
 * In plain English:
 * We take WooCommerce orders, tidy them up, find where the books live in the
 * warehouse, prepare the courier file, and try to save everyone from doing
 * spreadsheet gymnastics before the first coffee of the day.
 *
 * Important rules:
 * - Gift Notes are kept.
 * - Multiple copies of a title become X2, X3, etc.
 * - Shelving quantities use the real number of books required.
 * - CityPak gets full book titles when they fit.
 * - Very long CityPak descriptions become "N BOOKS".
 * - Any shortened CityPak order is recorded in "Shortened Orders" with a
 *   ready-to-copy message containing the order number and full title list.
 * - The Shelving List gets Excel conditional-formatting color scales.
 *
 * Maintenance rule:
 * If you change the order of the generated worksheets, check the XLSX
 * conditional-formatting code near the bottom of this class. Excel is very
 * particular about its XML. Think "librarian with a ruler", but less forgiving.
 * ============================================================================
 */

if (!defined('ABSPATH')) {
    exit;
}

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

class Jumpbooks_Order_Processor
{
    private static $instance = null;

    private $option_shelving_file = 'jop_shelving_file_path';
    private $option_shelving_name = 'jop_shelving_file_name';
    private $option_citypak_file = 'jop_citypak_template_path';
    private $option_citypak_name = 'jop_citypak_template_name';

    /*
     * CityPak Item Description limit.
     * If the titles fit inside 80 characters, we keep them.
     * If they do not, CityPak gets "N BOOKS" and Sheet 4 keeps the full story.
     */
    private $citypak_description_max_chars = 80;

    /*
     * Shelf colors are assigned per Shelf No.
     * Same shelf = same color, so warehouse picking is less "Where on earth is G-12?"
     */
    private $shelf_color_palette = [
        'FFDDEBF7',
        'FFE2F0D9',
        'FFFFF2CC',
        'FFFCE4D6',
        'FFE4DFEC',
        'FFD9EAD3',
        'FFF4CCCC',
        'FFD9E1F2',
        'FFEADCF8',
        'FFD0E0E3',
        'FFFCE5CD',
    ];

    // ------------------------------------------------------------------------
    // 1) PLUGIN BOOTSTRAP & WORDPRESS ADMIN
    // ------------------------------------------------------------------------

    /*
     * Singleton entry point.
     * Translation: one processor per request is plenty; we do not need clones
     * running around creating competing spreadsheets.
     */
    public static function instance()
    {
        if (self::$instance === null) {
            self::$instance = new self();
        }

        return self::$instance;
    }

    /*
     * Wire the plugin into WordPress.
     * These hooks connect the admin menu, settings save, and both export buttons.
     */
    private function __construct()
    {
        add_action('admin_menu', [$this, 'register_admin_menu']);
        add_action('admin_post_jop_save_settings', [$this, 'save_settings']);
        add_action('admin_post_jop_export_processed', [$this, 'export_processed']);
        add_action('admin_post_jop_export_citypak', [$this, 'export_citypak']);
    }

    /*
     * Add "Daily Dispatch" to the WordPress admin menu.
     * This is the front door for warehouse spreadsheet magic.
     */
    public function register_admin_menu()
    {
        add_menu_page(
            'Daily Dispatch',
            'Daily Dispatch',
            'manage_woocommerce',
            'jumpbooks-order-processor',
            [$this, 'render_export_page'],
            'dashicons-media-spreadsheet',
            56
        );

        add_submenu_page(
            'jumpbooks-order-processor',
            'Export Orders',
            'Export Orders',
            'manage_woocommerce',
            'jumpbooks-order-processor',
            [$this, 'render_export_page']
        );

        add_submenu_page(
            'jumpbooks-order-processor',
            'Settings',
            'Settings',
            'manage_woocommerce',
            'jumpbooks-order-processor-settings',
            [$this, 'render_settings_page']
        );
    }

    /*
     * Return the private working folder inside WordPress uploads.
     * If it does not exist yet, create it quietly and carry on.
     */
    private function get_storage_dir()
    {
        $upload_dir = wp_upload_dir();
        $dir = trailingslashit($upload_dir['basedir']) . 'jumpbooks-order-processor/';

        if (!file_exists($dir)) {
            wp_mkdir_p($dir);
        }

        return $dir;
    }

    /*
     * Security checkpoint.
     * Only WooCommerce managers with a valid nonce get past this bouncer.
     */
    private function verify_admin_request($nonce_action, $nonce_name)
    {
        if (!current_user_can('manage_woocommerce')) {
            wp_die('You do not have permission to use this export.');
        }

        check_admin_referer($nonce_action, $nonce_name);
    }

    /*
     * Admin-page styling only.
     * No order logic lives here; this just stops the plugin from looking like 2008.
     */
    private function admin_styles()
    {
        ?>
        <style>
            .jop-wrap {
                max-width: 1080px;
            }

            .jop-hero {
                background: linear-gradient(135deg, #0f172a 0%, #1d4ed8 100%);
                color: #fff;
                border-radius: 18px;
                padding: 28px 32px;
                margin: 22px 0;
                box-shadow: 0 12px 30px rgba(15, 23, 42, 0.18);
            }

            .jop-hero h1 {
                color: #fff;
                margin: 0 0 8px;
                font-size: 30px;
                font-weight: 800;
            }

            .jop-hero p {
                margin: 0;
                color: #dbeafe;
                font-size: 14px;
            }

            .jop-grid {
                display: grid;
                grid-template-columns: repeat(2, minmax(0, 1fr));
                gap: 18px;
                margin-top: 18px;
            }

            .jop-card {
                background: #fff;
                border: 1px solid #e5e7eb;
                border-radius: 16px;
                padding: 22px;
                box-shadow: 0 6px 20px rgba(15, 23, 42, 0.06);
            }

            .jop-card h2 {
                margin-top: 0;
                font-size: 18px;
                font-weight: 700;
                color: #111827;
            }

            .jop-card p {
                color: #6b7280;
            }

            .jop-status {
                display: inline-flex;
                align-items: center;
                padding: 5px 10px;
                border-radius: 999px;
                background: #eef2ff;
                color: #3730a3;
                font-weight: 600;
                font-size: 12px;
            }

            .jop-status-missing {
                background: #fee2e2;
                color: #991b1b;
            }

            .jop-file-box {
                background: #f8fafc;
                border: 1px dashed #cbd5e1;
                border-radius: 14px;
                padding: 16px;
                margin-top: 8px;
            }

            .jop-file-box input[type="file"] {
                width: 100%;
            }

            .jop-actions {
                display: flex;
                gap: 12px;
                flex-wrap: wrap;
                margin-top: 18px;
            }

            .jop-btn {
                border: 0;
                border-radius: 12px;
                padding: 11px 18px;
                font-weight: 700;
                cursor: pointer;
                text-decoration: none;
                display: inline-flex;
                align-items: center;
                gap: 8px;
                min-height: 42px;
            }

            .jop-btn-primary {
                background: #2563eb;
                color: #fff;
            }

            .jop-btn-primary:hover {
                background: #1d4ed8;
                color: #fff;
            }

            .jop-btn-dark {
                background: #111827;
                color: #fff;
            }

            .jop-btn-dark:hover {
                background: #000;
                color: #fff;
            }

            .jop-select {
                min-width: 320px;
                height: 42px;
                border-radius: 10px;
                border-color: #cbd5e1;
                padding: 0 12px;
                font-weight: 600;
            }

            .jop-note {
                background: #fffbeb;
                border: 1px solid #fde68a;
                color: #92400e;
                padding: 12px 14px;
                border-radius: 12px;
                margin-top: 14px;
            }

            @media (max-width: 900px) {
                .jop-grid {
                    grid-template-columns: 1fr;
                }

                .jop-select {
                    width: 100%;
                    min-width: 0;
                }
            }
        </style>
        <?php
    }

    // ------------------------------------------------------------------------
    // 2) SETTINGS & SAVED REFERENCE FILES
    // ------------------------------------------------------------------------

    /*
     * Render the Settings page where the monthly shelving file and CityPak
     * template are uploaded and remembered.
     */
    public function render_settings_page()
    {
        if (!current_user_can('manage_woocommerce')) {
            return;
        }

        $shelving_name = get_option($this->option_shelving_name, '');
        $citypak_name = get_option($this->option_citypak_name, '');

        $this->admin_styles();
        ?>
        <div class="wrap jop-wrap">
            <div class="jop-hero">
                <h1>Daily Dispatch Settings</h1>
                <p>Upload the monthly shelving sheet and fixed CityPak template used by the export tool.</p>
            </div>

            <?php if (isset($_GET['jop_saved'])) : ?>
                <div class="notice notice-success is-dismissible">
                    <p>Settings saved successfully.</p>
                </div>
            <?php endif; ?>

            <form method="post" action="<?php echo esc_url(admin_url('admin-post.php')); ?>" enctype="multipart/form-data">
                <?php wp_nonce_field('jop_save_settings', 'jop_nonce'); ?>
                <input type="hidden" name="action" value="jop_save_settings">

                <div class="jop-grid">
                    <div class="jop-card">
                        <h2>Monthly Shelving Sheet</h2>
                        <p>Upload this once a month. The plugin uses it to match SKU to Shelf No.</p>

                        <p>
                            Current file:
                            <?php if ($shelving_name) : ?>
                                <span class="jop-status"><?php echo esc_html($shelving_name); ?></span>
                            <?php else : ?>
                                <span class="jop-status jop-status-missing">Not uploaded</span>
                            <?php endif; ?>
                        </p>

                        <div class="jop-file-box">
                            <input type="file" name="jop_shelving_file" accept=".xlsx,.xlsm">
                        </div>

                        <p class="description">Expected format: SKU in column A and Shelf No in column C. If a sheet named "all" exists, it will be used.</p>
                    </div>

                    <div class="jop-card">
                        <h2>CityPak Template</h2>
                        <p>Upload your fixed CityPak template. You only need to replace it if CityPak changes their format.</p>

                        <p>
                            Current file:
                            <?php if ($citypak_name) : ?>
                                <span class="jop-status"><?php echo esc_html($citypak_name); ?></span>
                            <?php else : ?>
                                <span class="jop-status jop-status-missing">Not uploaded</span>
                            <?php endif; ?>
                        </p>

                        <div class="jop-file-box">
                            <input type="file" name="jop_citypak_file" accept=".xlsx,.xlsm">
                        </div>
                    </div>
                </div>

                <div class="jop-actions">
                    <button type="submit" class="jop-btn jop-btn-primary">Save Settings</button>
                    <a class="jop-btn jop-btn-dark" href="<?php echo esc_url(admin_url('admin.php?page=jumpbooks-order-processor')); ?>">Go to Export Orders</a>
                </div>
            </form>
        </div>
        <?php
    }

    /*
     * Save newly uploaded reference files.
     * Existing files stay in place unless the user actually uploads replacements.
     */
    public function save_settings()
    {
        $this->verify_admin_request('jop_save_settings', 'jop_nonce');

        $storage_dir = $this->get_storage_dir();

        if (!empty($_FILES['jop_shelving_file']['name'])) {
            $this->save_uploaded_file(
                $_FILES['jop_shelving_file'],
                $storage_dir . 'shelving-current.xlsx',
                $this->option_shelving_file,
                $this->option_shelving_name
            );
        }

        if (!empty($_FILES['jop_citypak_file']['name'])) {
            $this->save_uploaded_file(
                $_FILES['jop_citypak_file'],
                $storage_dir . 'citypak-template.xlsx',
                $this->option_citypak_file,
                $this->option_citypak_name
            );
        }

        wp_safe_redirect(admin_url('admin.php?page=jumpbooks-order-processor-settings&jop_saved=1'));
        exit;
    }

    /*
     * Validate and store one uploaded Excel file.
     * Only .xlsx and .xlsm are invited to this party.
     */
    private function save_uploaded_file($file, $destination, $option_path, $option_name)
    {
        $allowed_ext = ['xlsx', 'xlsm'];
        $ext = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));

        if (!in_array($ext, $allowed_ext, true)) {
            wp_die('Only .xlsx and .xlsm files are allowed.');
        }

        if (!is_uploaded_file($file['tmp_name'])) {
            wp_die('Upload failed. Please try again.');
        }

        if (!move_uploaded_file($file['tmp_name'], $destination)) {
            wp_die('Could not save uploaded file.');
        }

        update_option($option_path, $destination);
        update_option($option_name, sanitize_file_name($file['name']));
    }

    /*
     * Render the main Daily Dispatch screen:
     * choose Packing/Pickup, then download Processed Workbook or CityPak export.
     */
    public function render_export_page()
    {
        if (!current_user_can('manage_woocommerce')) {
            return;
        }

        $statuses = $this->get_export_status_choices();

        $shelving_name = get_option($this->option_shelving_name, '');
        $citypak_name = get_option($this->option_citypak_name, '');

        $this->admin_styles();
        ?>
        <div class="wrap jop-wrap">
            <div class="jop-hero">
                <h1>Daily Dispatch</h1>
                <p>Export WooCommerce orders directly into the same format as your desktop processor.</p>
            </div>

            <div class="jop-grid">
                <div class="jop-card">
                    <h2>Saved Files</h2>

                    <p>
                        Monthly Shelving Sheet:
                        <?php if ($shelving_name) : ?>
                            <span class="jop-status"><?php echo esc_html($shelving_name); ?></span>
                        <?php else : ?>
                            <span class="jop-status jop-status-missing">Not uploaded</span>
                        <?php endif; ?>
                    </p>

                    <p>
                        CityPak Template:
                        <?php if ($citypak_name) : ?>
                            <span class="jop-status"><?php echo esc_html($citypak_name); ?></span>
                        <?php else : ?>
                            <span class="jop-status jop-status-missing">Not uploaded</span>
                        <?php endif; ?>
                    </p>

                    <div class="jop-actions">
                        <a class="jop-btn jop-btn-dark" href="<?php echo esc_url(admin_url('admin.php?page=jumpbooks-order-processor-settings')); ?>">Open Settings</a>
                    </div>
                </div>

                <div class="jop-card">
                    <h2>Export Rules</h2>
                    <p>The exported processed workbook will contain:</p>
                    <ul style="margin-left:18px; color:#4b5563;">
                        <li>Sheet 1: Orders</li>
                        <li>Sheet 2: Processed Orders</li>
                        <li>Sheet 3: Shelving List</li>
                        <li>Sheet 4: Shortened Orders</li>
                    </ul>
                    <p>Gift Note is kept. Missing shelf numbers are moved to the bottom. Shortened CityPak orders get a copy/paste message in Sheet 4.</p>
                </div>
            </div>

            <div class="jop-card" style="margin-top:18px;">
                <h2>Export by Order Tab</h2>

                <?php if (empty($statuses)) : ?>
                    <div class="jop-note">
                        Packing and Pickup Orders statuses were not found. Please check the exact WooCommerce status names.
                    </div>
                <?php endif; ?>

                <form method="post" action="<?php echo esc_url(admin_url('admin-post.php')); ?>" style="margin-bottom:22px;">
                    <?php wp_nonce_field('jop_export_processed', 'jop_nonce'); ?>
                    <input type="hidden" name="action" value="jop_export_processed">

                    <p>
                        <label for="jop_status_processed"><strong>Choose tab/status</strong></label><br>
                        <select id="jop_status_processed" name="order_status" class="jop-select">
                            <?php foreach ($statuses as $key => $label) : ?>
                                <option value="<?php echo esc_attr($key); ?>"><?php echo esc_html($label); ?></option>
                            <?php endforeach; ?>
                        </select>
                    </p>

                    <button type="submit" class="jop-btn jop-btn-primary">Download Processed Workbook</button>
                </form>

                <hr>

                <form method="post" action="<?php echo esc_url(admin_url('admin-post.php')); ?>">
                    <?php wp_nonce_field('jop_export_citypak', 'jop_nonce'); ?>
                    <input type="hidden" name="action" value="jop_export_citypak">

                    <p>
                        <label for="jop_status_citypak"><strong>Choose tab/status</strong></label><br>
                        <select id="jop_status_citypak" name="order_status" class="jop-select">
                            <?php foreach ($statuses as $key => $label) : ?>
                                <option value="<?php echo esc_attr($key); ?>"><?php echo esc_html($label); ?></option>
                            <?php endforeach; ?>
                        </select>
                    </p>

                    <button type="submit" class="jop-btn jop-btn-dark">Download CityPak Export</button>
                </form>
            </div>
        </div>
        <?php
    }

    /*
     * Find the WooCommerce statuses that look like Packing or Pickup.
     * We search both labels and keys because custom-status plugins can be creative.
     */
    private function get_export_status_choices()
    {
        $all_statuses = wc_get_order_statuses();

        $choices = [];

        foreach ($all_statuses as $key => $label) {
            $label_lower = strtolower($label);
            $key_lower = strtolower($key);

            if (
                strpos($label_lower, 'packing') !== false ||
                strpos($key_lower, 'packing') !== false
            ) {
                $choices[$key] = $label;
            }
        }

        foreach ($all_statuses as $key => $label) {
            $label_lower = strtolower($label);
            $key_lower = strtolower($key);

            if (
                strpos($label_lower, 'pickup') !== false ||
                strpos($label_lower, 'pick up') !== false ||
                strpos($key_lower, 'pickup') !== false ||
                strpos($key_lower, 'pick-up') !== false
            ) {
                $choices[$key] = $label;
            }
        }

        return $choices;
    }

    // ------------------------------------------------------------------------
    // 3) EXPORT PIPELINES
    // ------------------------------------------------------------------------

    /*
     * Build the main processed workbook.
     *
     * Sheet 1 = Orders
     * Sheet 2 = Processed Orders
     * Sheet 3 = Shelving List
     * Sheet 4 = Shortened Orders
     *
     * One click in WordPress, four sheets out. Excel earns its keep today.
     */
    public function export_processed()
    {
        $this->verify_admin_request('jop_export_processed', 'jop_nonce');

        $status = isset($_POST['order_status']) ? sanitize_text_field(wp_unslash($_POST['order_status'])) : '';
        if (!$status) {
            wp_die('Order status is missing.');
        }

        try {
            $orders = $this->get_orders_by_status($status);

            if (empty($orders)) {
                wp_die(
                    'Daily Dispatch found 0 WooCommerce orders for the selected status: ' .
                    esc_html($status)
                );
            }

            $shelf_map = $this->load_shelf_map();

            $raw_rows = $this->build_raw_order_rows($orders);
            $processed_rows = $this->build_processed_rows($orders);
            $shelving_rows = $this->build_shelving_rows($orders, $shelf_map);

            $spreadsheet = new Spreadsheet();

            $orders_sheet = $spreadsheet->getActiveSheet();
            $orders_sheet->setTitle('Orders');
            $this->write_orders_sheet($orders_sheet, $raw_rows);

            $processed_sheet = $spreadsheet->createSheet();
            $processed_sheet->setTitle('Processed Orders');
            $this->write_processed_sheet($processed_sheet, $processed_rows);

            $shelving_sheet = $spreadsheet->createSheet();
            $shelving_sheet->setTitle('Shelving List');
            $this->write_shelving_sheet($shelving_sheet, $shelving_rows);

            // Create a copy/paste reference sheet only for orders whose
            // CityPak Item Description had to be shortened to "N BOOKS".
            $shortened_rows = $this->build_shortened_citypak_rows($processed_rows);

            $shortened_sheet = $spreadsheet->createSheet();
            $shortened_sheet->setTitle('Shortened Orders');
            $this->write_shortened_orders_sheet($shortened_sheet, $shortened_rows);

            $spreadsheet->setActiveSheetIndex(0);

            $filename = 'daily-dispatch-processed-' . $this->clean_filename_status($status) . '-' . date('Y-m-d-His') . '.xlsx';

            $shelving_last_row = max(2, count($shelving_rows) + 1);

            $this->download_spreadsheet(
                $spreadsheet,
                $filename,
                $shelving_last_row
            );
        } catch (\Throwable $e) {
            wp_die(
                'Daily Dispatch processed workbook export failed.<br><br><strong>Error:</strong> ' .
                esc_html($e->getMessage())
            );
        }
    }

    /*
     * Build the CityPak courier workbook from the selected WooCommerce status.
     * The saved CityPak template provides the structure; we provide the order data.
     */
    public function export_citypak()
    {
        $this->verify_admin_request('jop_export_citypak', 'jop_nonce');

        $status = isset($_POST['order_status']) ? sanitize_text_field(wp_unslash($_POST['order_status'])) : '';
        if (!$status) {
            wp_die('Order status is missing.');
        }

        try {
            $orders = $this->get_orders_by_status($status);

            if (empty($orders)) {
                wp_die(
                    'Daily Dispatch found 0 WooCommerce orders for the selected status: ' .
                    esc_html($status)
                );
            }

            $processed_rows = $this->build_processed_rows($orders);

            if (empty($processed_rows)) {
                wp_die(
                    'Daily Dispatch found orders for this status, but no processable order rows were created.'
                );
            }

            $template_path = get_option($this->option_citypak_file, '');
            if (!$template_path || !file_exists($template_path)) {
                wp_die('CityPak template is not uploaded. Please upload it in Daily Dispatch → Settings.');
            }

            $spreadsheet = IOFactory::load($template_path);
            $sheet = $spreadsheet->getActiveSheet();

            $this->write_citypak_sheet($sheet, $processed_rows);

            $filename = 'daily-dispatch-citypak-' . $this->clean_filename_status($status) . '-' . date('Y-m-d-His') . '.xlsx';

            $this->download_spreadsheet($spreadsheet, $filename);
        } catch (\Throwable $e) {
            wp_die(
                'Daily Dispatch CityPak export failed.<br><br><strong>Error:</strong> ' .
                esc_html($e->getMessage())
            );
        }
    }

    /*
     * Fetch WooCommerce orders for a status.
     * We try both "packing" and "wc-packing" styles because live sites sometimes
     * disagree about which hat a custom status should wear.
     */
    private function get_orders_by_status($status)
    {
        $original_status = trim((string) $status);

        if ($original_status === '') {
            return [];
        }

        // WooCommerce order-status keys are often returned as "wc-packing",
        // while wc_get_orders() commonly works with the clean slug "packing".
        $clean_status = preg_replace('/^wc-/', '', $original_status);

        $query_args = [
            'status' => [$clean_status],
            'limit' => -1,
            'orderby' => 'date',
            'order' => 'ASC',
            'return' => 'objects',
        ];

        $orders = wc_get_orders($query_args);

        // Fallback for custom status implementations that expect the full wc- key.
        if (empty($orders) && $clean_status !== $original_status) {
            $query_args['status'] = [$original_status];
            $orders = wc_get_orders($query_args);
        }

        return is_array($orders) ? $orders : [];
    }

    // ------------------------------------------------------------------------
    // 4) TURN WOOCOMMERCE ORDERS INTO CLEAN DATA ARRAYS
    // ------------------------------------------------------------------------

    /*
     * Convert WooCommerce line items into the raw Orders-sheet format.
     * One purchased line item = one row, just like a traditional order export.
     */
    private function build_raw_order_rows($orders)
    {
        $rows = [];

        foreach ($orders as $order) {
            $item_number = 1;

            foreach ($order->get_items() as $item) {
                $product = $item->get_product();
                $sku = $product ? trim((string) $product->get_sku()) : '';

                $rows[] = [
                    'Order Number' => $order->get_order_number(),
                    'Gift Note' => $this->get_gift_note($order),
                    'Phone (Billing)' => $order->get_billing_phone(),
                    'First Name (Shipping)' => $order->get_shipping_first_name(),
                    'Last Name (Shipping)' => $order->get_shipping_last_name(),
                    'Phone (Shipping)' => $this->get_shipping_phone($order),
                    'Address 1&2 (Shipping)' => trim($order->get_shipping_address_1() . ' ' . $order->get_shipping_address_2()),
                    'City (Shipping)' => $order->get_shipping_city(),
                    'Payment Method Title' => $order->get_payment_method_title(),
                    'Order Total Amount' => (float) $order->get_total(),
                    'Shipping Method Title' => $this->get_shipping_method_title($order),
                    'Item #' => $item_number,
                    'Item Name' => $item->get_name(),
                    'SKU' => $sku,
                    'Quantity' => (int) $item->get_quantity(),
                    'Item Cost' => (float) $order->get_line_subtotal($item, false, false),
                ];

                $item_number++;
            }
        }

        return $rows;
    }

    /*
     * Collapse each WooCommerce order into one processed row.
     * Multiple books become "Book A / Book B / Book C X2", while Book Count keeps
     * the real physical quantity for CityPak and Shortened Orders.
     */
    private function build_processed_rows($orders)
    {
        $rows = [];

        foreach ($orders as $order) {
            $items = [];
            $first_item_name = '';
            $book_count = 0;

            foreach ($order->get_items() as $item) {
                $title = $this->clean_item_title($item->get_name());

                if ($first_item_name === '') {
                    $first_item_name = $title;
                }

                $qty = (int) $item->get_quantity();
                $book_count += max(1, $qty);

                if ($qty > 1) {
                    $title .= ' X' . $qty;
                }

                $items[] = $title;
            }

            $billing_phone = $order->get_billing_phone();
            $shipping_phone = $this->get_shipping_phone($order);
            $phone = $shipping_phone ? $shipping_phone : $billing_phone;

            $full_name = trim($order->get_shipping_first_name() . ' ' . $order->get_shipping_last_name());
            if (!$full_name) {
                $full_name = trim($order->get_billing_first_name() . ' ' . $order->get_billing_last_name());
            }

            $shipping_address = trim($order->get_shipping_address_1() . ' ' . $order->get_shipping_address_2());
            if (!$shipping_address) {
                $shipping_address = trim($order->get_billing_address_1() . ' ' . $order->get_billing_address_2());
            }

            $city = $order->get_shipping_city();
            if (!$city) {
                $city = $order->get_billing_city();
            }

            $payment_title = $order->get_payment_method_title();
            $cod_value = $this->payment_to_cod_value($payment_title);
            $total = (float) $order->get_total();

            if ($cod_value === 'NO') {
                $total = 0;
            }

            $rows[] = [
                'Order Number' => $order->get_order_number(),
                'Gift Note' => $this->get_gift_note($order),
                'Phone (Billing)' => $phone,
                'Full Name' => $full_name,
                'Address 1&2 (Shipping)' => $shipping_address,
                'City (Shipping)' => $city,
                'Payment Method Title' => $cod_value,
                'Order Total Amount' => $total,
                'Shipping Method Title' => $this->get_shipping_method_title($order),
                'Item #' => 1,
                'Item Name' => implode(' / ', $items),

                // Internal helper value used by CityPak and Shortened Orders.
                // It is intentionally not included in the Processed Orders sheet headers.
                'Book Count' => $book_count,

                // Internal timestamp used only to keep Processed Orders newest-first.
                'Order Date Timestamp' => (int) $order->get_date_created()->getTimestamp(),
            ];
        }

        // Newest order first, oldest order last.
        // This affects only the Processed Orders pipeline.
        usort($rows, function ($a, $b) {
            $date_compare = (int) $b['Order Date Timestamp'] <=> (int) $a['Order Date Timestamp'];

            if ($date_compare !== 0) {
                return $date_compare;
            }

            return strnatcasecmp((string) $b['Order Number'], (string) $a['Order Number']);
        });

        return $rows;
    }

    /*
     * Count required copies by SKU and attach each SKU to its shelf.
     * Missing shelf numbers are not lost; they are politely sent to the bottom.
     */
    private function build_shelving_rows($orders, $shelf_map)
    {
        $sku_counts = [];
        $sku_titles = [];

        foreach ($orders as $order) {
            foreach ($order->get_items() as $item) {
                $product = $item->get_product();
                $sku = $product ? trim((string) $product->get_sku()) : '';

                if (!$sku) {
                    continue;
                }

                $qty = (int) $item->get_quantity();

                if (!isset($sku_counts[$sku])) {
                    $sku_counts[$sku] = 0;
                    $sku_titles[$sku] = $this->clean_item_title($item->get_name());
                }

                $sku_counts[$sku] += max(1, $qty);
            }
        }

        $rows = [];

        foreach ($sku_counts as $sku => $count) {
            $shelf_no = isset($shelf_map[$sku]) ? trim((string) $shelf_map[$sku]) : '';

            $rows[] = [
                'Item Name' => $sku_titles[$sku],
                'SKU' => $sku,
                'No of Books' => $count,
                'Shelf No' => $shelf_no,
            ];
        }

        usort($rows, function ($a, $b) {
            $shelf_a = trim((string) $a['Shelf No']);
            $shelf_b = trim((string) $b['Shelf No']);

            $a_blank = $shelf_a === '';
            $b_blank = $shelf_b === '';

            if ($a_blank && !$b_blank) {
                return 1;
            }

            if (!$a_blank && $b_blank) {
                return -1;
            }

            if ($a_blank && $b_blank) {
                return strnatcasecmp((string) $a['Item Name'], (string) $b['Item Name']);
            }

            $shelf_compare = strnatcasecmp($shelf_a, $shelf_b);

            if ($shelf_compare !== 0) {
                return $shelf_compare;
            }

            return strnatcasecmp((string) $a['Item Name'], (string) $b['Item Name']);
        });

        return $rows;
    }

    /*
     * Read the monthly shelving workbook and build:
     *     SKU => Shelf No
     * Prefer the sheet named "all"; otherwise use the first sheet.
     */
    private function load_shelf_map()
    {
        $path = get_option($this->option_shelving_file, '');

        if (!$path || !file_exists($path)) {
            wp_die('Monthly shelving sheet is not uploaded. Please upload it in Daily Dispatch → Settings.');
        }

        $spreadsheet = IOFactory::load($path);

        $sheet = $spreadsheet->getSheetByName('all');
        if (!$sheet) {
            $sheet = $spreadsheet->getSheet(0);
        }

        $highest_row = $sheet->getHighestDataRow();
        $map = [];

        for ($row = 1; $row <= $highest_row; $row++) {
            $sku = trim((string) $sheet->getCell('A' . $row)->getValue());
            $shelf = trim((string) $sheet->getCell('C' . $row)->getValue());

            if ($sku === '') {
                continue;
            }

            if (!isset($map[$sku])) {
                $map[$sku] = $shelf;
            }
        }

        return $map;
    }

    // ------------------------------------------------------------------------
    // 5) WRITE & FORMAT EXCEL WORKSHEETS
    // ------------------------------------------------------------------------

    /*
     * Write Sheet 1: Orders.
     * Gift Note stays here because surprises should survive the export process.
     */
    private function write_orders_sheet($sheet, $rows)
    {
        $headers = [
            'Order Number',
            'Gift Note',
            'Phone (Billing)',
            'First Name (Shipping)',
            'Last Name (Shipping)',
            'Phone (Shipping)',
            'Address 1&2 (Shipping)',
            'City (Shipping)',
            'Payment Method Title',
            'Order Total Amount',
            'Shipping Method Title',
            'Item #',
            'Item Name',
            'SKU',
            'Quantity',
            'Item Cost',
        ];

        $this->write_headers($sheet, $headers);

        $r = 2;

        foreach ($rows as $row) {
            $c = 1;

            foreach ($headers as $header) {
                $value = isset($row[$header]) ? $row[$header] : '';

                if (in_array($header, ['Order Total Amount', 'Item #', 'Quantity', 'Item Cost'], true)) {
                    $sheet->setCellValue([$c, $r], $value);
                } else {
                    $sheet->setCellValueExplicit([$c, $r], $value, DataType::TYPE_STRING);
                }

                $c++;
            }

            $r++;
        }

        $last_row = max(2, $r - 1);
        $this->format_orders_like_export($sheet, count($headers), $last_row);
    }

    /*
     * Write Sheet 2: Processed Orders.
     * One order = one row, with all titles grouped into the Item Name cell.
     */
    private function write_processed_sheet($sheet, $rows)
    {
        $headers = [
            'Order Number',
            'Gift Note',
            'Phone (Billing)',
            'Full Name',
            'Address 1&2 (Shipping)',
            'City (Shipping)',
            'Payment Method Title',
            'Order Total Amount',
            'Shipping Method Title',
            'Item #',
            'Item Name',
        ];

        $this->write_headers($sheet, $headers);

        $r = 2;

        foreach ($rows as $row) {
            $c = 1;

            foreach ($headers as $header) {
                $value = isset($row[$header]) ? $row[$header] : '';

                if (in_array($header, ['Order Total Amount', 'Item #'], true)) {
                    $sheet->setCellValue([$c, $r], $value);
                } else {
                    $sheet->setCellValueExplicit([$c, $r], $value, DataType::TYPE_STRING);
                }

                $c++;
            }

            $r++;
        }

        $last_row = max(2, $r - 1);
        $this->format_processed_like_export($sheet, count($headers), $last_row);
    }

    /*
     * Write Sheet 3: Shelving List.
     * Rows are colored by shelf so pickers can visually group nearby books.
     */
    private function write_shelving_sheet($sheet, $rows)
    {
        $headers = [
            'Item Name',
            'SKU',
            'No of Books',
            'Shelf No',
        ];

        $this->write_headers($sheet, $headers);

        $shelf_colors = [];
        $color_index = 0;
        $r = 2;

        foreach ($rows as $row) {
            $shelf_no = trim((string) $row['Shelf No']);

            if ($shelf_no === '') {
                $row_color = 'FFE7E6E6';
            } else {
                if (!isset($shelf_colors[$shelf_no])) {
                    $shelf_colors[$shelf_no] = $this->shelf_color_palette[$color_index % count($this->shelf_color_palette)];
                    $color_index++;
                }

                $row_color = $shelf_colors[$shelf_no];
            }

            $sheet->setCellValue('A' . $r, $row['Item Name']);
            $sheet->setCellValueExplicit('B' . $r, $row['SKU'], DataType::TYPE_STRING);
            $sheet->setCellValue('C' . $r, $row['No of Books']);
            $sheet->setCellValue('D' . $r, $shelf_no);

            $sheet->getStyle('A' . $r . ':D' . $r)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()
                ->setARGB($row_color);

            $sheet->getStyle('A' . $r . ':D' . $r)->getFont()
                ->setName('Arial')
                ->setSize(9);

            $sheet->getStyle('A' . $r . ':D' . $r)->getBorders()->getAllBorders()
                ->setBorderStyle(Border::BORDER_THIN)
                ->getColor()
                ->setARGB('FFE2F0D9');

            $sheet->getRowDimension($r)->setRowHeight(18);

            $r++;
        }

        $last_row = max(2, $r - 1);

        $sheet->freezePane('A2');
        $sheet->setAutoFilter('A1:D' . $last_row);

        $sheet->getStyle('A1:D1')->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('FF548235');

        $sheet->getStyle('A1:D1')->getFont()
            ->setName('Arial')
            ->setSize(11)
            ->setBold(true)
            ->getColor()
            ->setARGB('FFFFFFFF');

        $sheet->getStyle('A1:D1')->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_CENTER)
            ->setVertical(Alignment::VERTICAL_CENTER)
            ->setWrapText(true);

        $sheet->getStyle('A1:D' . $last_row)->getBorders()->getAllBorders()
            ->setBorderStyle(Border::BORDER_THIN)
            ->getColor()
            ->setARGB('FFE2F0D9');

        $sheet->getStyle('A2:D' . $last_row)->getAlignment()
            ->setWrapText(false)
            ->setShrinkToFit(true)
            ->setVertical(Alignment::VERTICAL_CENTER);

        $sheet->getStyle('C2:D' . $last_row)->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $sheet->getColumnDimension('A')->setWidth(55);
        $sheet->getColumnDimension('B')->setWidth(18);
        $sheet->getColumnDimension('C')->setWidth(14);
        $sheet->getColumnDimension('D')->setWidth(14);
        $sheet->getRowDimension(1)->setRowHeight(24);
    }

    /*
     * Fill the CityPak template.
     *
     * Column W is the important bit for book descriptions:
     * - keep full titles when they fit,
     * - reduce font size when helpful,
     * - fall back to "N BOOKS" only when the description exceeds the limit.
     */
    private function write_citypak_sheet($sheet, $processed_rows)
    {
        $highest_row = max(3, $sheet->getHighestRow());

        for ($row = 3; $row <= $highest_row; $row++) {
            for ($col = 1; $col <= 25; $col++) {
                $sheet->setCellValue([$col, $row], null);
            }
        }

        $sender_name = 'Jump Pvt Ltd';
        $sender_address_1 = 'No 276/1, High Level Road';
        $sender_address_2 = '';
        $sender_address_3 = '';
        $sender_city = 'Nugegoda';
        $sender_contact = 'Akmal Iqbal';
        $sender_phone_1 = '';
        $sender_phone_2 = '';
        $weight = 1000;
        $billing_type = 'SENDER ACCOUNT';
        $item_type = 'PARCEL';
        $pieces = 1;

        $row_num = 3;

        foreach ($processed_rows as $row) {
            $order_no = isset($row['Order Number']) ? $row['Order Number'] : '';
            $phone = isset($row['Phone (Billing)']) ? $row['Phone (Billing)'] : '';
            $name = isset($row['Full Name']) ? $row['Full Name'] : '';
            $address = isset($row['Address 1&2 (Shipping)']) ? $row['Address 1&2 (Shipping)'] : '';
            $city = isset($row['City (Shipping)']) ? $row['City (Shipping)'] : '';
            $cod = isset($row['Payment Method Title']) ? $row['Payment Method Title'] : '';
            $amount = isset($row['Order Total Amount']) ? (float) $row['Order Total Amount'] : 0;
            $items = isset($row['Item Name']) ? $row['Item Name'] : '';

            $sheet->setCellValueExplicit('B' . $row_num, $order_no, DataType::TYPE_STRING);
            $sheet->setCellValue('C' . $row_num, $sender_name);
            $sheet->setCellValue('D' . $row_num, $sender_address_1);
            $sheet->setCellValue('E' . $row_num, $sender_address_2);
            $sheet->setCellValue('F' . $row_num, $sender_address_3);
            $sheet->setCellValue('G' . $row_num, $sender_city);
            $sheet->setCellValue('H' . $row_num, $sender_contact);
            $sheet->setCellValueExplicit('I' . $row_num, $sender_phone_1, DataType::TYPE_STRING);
            $sheet->setCellValueExplicit('J' . $row_num, $sender_phone_2, DataType::TYPE_STRING);

            $sheet->setCellValue('K' . $row_num, $this->clean_citypak_text($name));
            $sheet->setCellValue('L' . $row_num, $this->clean_citypak_text($address));
            $sheet->setCellValue('M' . $row_num, '');
            $sheet->setCellValue('N' . $row_num, '');
            $sheet->setCellValue('O' . $row_num, $this->clean_citypak_text($city));
            $sheet->setCellValueExplicit('P' . $row_num, $this->clean_citypak_text($phone), DataType::TYPE_STRING);
            $sheet->setCellValue('Q' . $row_num, '');
            $sheet->setCellValue('R' . $row_num, '');

            $sheet->setCellValue('S' . $row_num, $weight);
            $sheet->setCellValue('T' . $row_num, $billing_type);
            $sheet->setCellValue('U' . $row_num, $item_type);
            $sheet->setCellValue('V' . $row_num, $pieces);

            $book_count = isset($row['Book Count'])
                ? max(1, (int) $row['Book Count'])
                : $this->count_books_from_item_description($items);

            /*
             * Keep the real titles whenever they fit inside the existing
             * CityPak Item Description character limit. Only descriptions
             * over the limit become "3 BOOKS", "4 BOOKS", etc.
             */
            $description = $this->make_citypak_description($items, $book_count);
            $sheet->setCellValue('W' . $row_num, $description);

            // Reduce font size for full-title descriptions that still fit.
            // Shortened "N BOOKS" descriptions stay at normal size.
            $description_was_shortened = $this->citypak_description_needs_shortening($items);
            $description_length = function_exists('mb_strlen')
                ? mb_strlen($description, 'UTF-8')
                : strlen($description);

            $description_font_size = 10;

            if (!$description_was_shortened) {
                if ($description_length > 55) {
                    $description_font_size = 9;
                }

                if ($description_length > 68) {
                    $description_font_size = 8;
                }

                if ($description_length > 76) {
                    $description_font_size = 7;
                }
            }

            $sheet->getStyle('W' . $row_num)->getFont()
                ->setSize($description_font_size);

            $sheet->getStyle('W' . $row_num)->getAlignment()
                ->setShrinkToFit(true)
                ->setWrapText(false)
                ->setVertical(Alignment::VERTICAL_CENTER);

            $sheet->setCellValue('X' . $row_num, $cod === 'YES' ? 'YES' : 'NO');
            $sheet->setCellValue('Y' . $row_num, $cod === 'YES' ? $amount : 0);

            $row_num++;
        }
    }

    /*
     * Tiny helper: write an array of column names into row 1.
     * Boring? Yes. Useful? Also yes.
     */
    private function write_headers($sheet, $headers)
    {
        $col = 1;

        foreach ($headers as $header) {
            $sheet->setCellValue([$col, 1], $header);
            $col++;
        }
    }

    /*
     * Make the Orders sheet readable: header color, borders, widths, filters,
     * row heights, and shrink-to-fit for long text.
     */
    private function format_orders_like_export($sheet, $max_col, $max_row)
    {
        $sheet->freezePane('A2');
        $sheet->setAutoFilter('A1:' . $this->column_letter($max_col) . $max_row);

        $header_range = 'A1:' . $this->column_letter($max_col) . '1';
        $data_range = 'A1:' . $this->column_letter($max_col) . $max_row;

        $sheet->getStyle($header_range)->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('FF4472C4');

        $sheet->getStyle($header_range)->getFont()
            ->setName('Arial')
            ->setBold(true)
            ->getColor()
            ->setARGB('FFFFFFFF');

        $sheet->getStyle($data_range)->getBorders()->getAllBorders()
            ->setBorderStyle(Border::BORDER_THIN)
            ->getColor()
            ->setARGB('FFD9E2F3');

        $sheet->getStyle($data_range)->getAlignment()
            ->setVertical(Alignment::VERTICAL_CENTER)
            ->setWrapText(false)
            ->setShrinkToFit(true);

        $sheet->getRowDimension(1)->setRowHeight(22);

        for ($r = 2; $r <= $max_row; $r++) {
            $sheet->getRowDimension($r)->setRowHeight(18);
        }

        $widths = [
            'A' => 16,
            'B' => 14,
            'C' => 18,
            'D' => 20,
            'E' => 20,
            'F' => 18,
            'G' => 45,
            'H' => 20,
            'I' => 24,
            'J' => 18,
            'K' => 28,
            'L' => 10,
            'M' => 55,
            'N' => 18,
            'O' => 12,
            'P' => 14,
        ];

        foreach ($widths as $col => $width) {
            $sheet->getColumnDimension($col)->setWidth($width);
        }
    }

    /*
     * Format the Processed Orders sheet so long customer/order data stays usable
     * without turning the workbook into a horizontal road trip.
     */
    private function format_processed_like_export($sheet, $max_col, $max_row)
    {
        $sheet->freezePane('A2');
        $sheet->setAutoFilter('A1:' . $this->column_letter($max_col) . $max_row);

        $header_range = 'A1:' . $this->column_letter($max_col) . '1';
        $data_range = 'A1:' . $this->column_letter($max_col) . $max_row;

        $sheet->getStyle($header_range)->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('FF5B9BD5');

        $sheet->getStyle($header_range)->getFont()
            ->setName('Arial')
            ->setBold(true)
            ->getColor()
            ->setARGB('FFFFFFFF');

        $sheet->getStyle($data_range)->getBorders()->getAllBorders()
            ->setBorderStyle(Border::BORDER_THIN)
            ->getColor()
            ->setARGB('FFDDEBF7');

        $sheet->getStyle($data_range)->getAlignment()
            ->setVertical(Alignment::VERTICAL_CENTER)
            ->setWrapText(false)
            ->setShrinkToFit(true);

        $sheet->getRowDimension(1)->setRowHeight(22);

        for ($r = 2; $r <= $max_row; $r++) {
            $sheet->getRowDimension($r)->setRowHeight(24);
        }

        $widths = [
            'A' => 16,
            'B' => 14,
            'C' => 18,
            'D' => 26,
            'E' => 45,
            'F' => 20,
            'G' => 18,
            'H' => 18,
            'I' => 28,
            'J' => 10,
            'K' => 65,
        ];

        foreach ($widths as $col => $width) {
            $sheet->getColumnDimension($col)->setWidth($width);
        }
    }

    // ------------------------------------------------------------------------
    // 6) SMALL ORDER-DATA HELPERS
    // ------------------------------------------------------------------------

    /*
     * Collect all shipping method names attached to an order.
     */
    private function get_shipping_method_title($order)
    {
        $methods = [];

        foreach ($order->get_shipping_methods() as $shipping_item) {
            $methods[] = $shipping_item->get_name();
        }

        return implode(', ', $methods);
    }

    /*
     * Try the common shipping-phone meta keys.
     * WooCommerce extensions do not always agree on naming, naturally.
     */
    private function get_shipping_phone($order)
    {
        return $this->get_order_meta_any($order, [
            '_shipping_phone',
            'shipping_phone',
            'Phone (Shipping)',
            '_shipping_phone_number',
        ]);
    }

    /*
     * Find the Gift Note without assuming every WooCommerce plugin uses the
     * same meta key. Spoiler: they absolutely do not.
     *
     * We check:
     * 1. Known order-level keys.
     * 2. Every order meta key whose name looks like gift + note/message.
     * 3. Order-item meta as a fallback, because some checkout/add-on plugins
     *    attach the note to a line item instead of the order itself.
     *
     * We intentionally do NOT use WooCommerce's generic customer note here,
     * because delivery instructions such as "call before arriving" are not
     * necessarily Gift Notes.
     */
    private function get_gift_note($order)
    {
        // First try the usual suspects. Fast and predictable when we know the key.
        $known_keys = [
            'Gift Note',
            'gift_note',
            '_gift_note',
            'gift-note',
            '_gift-note',
            'gift note',
            '_gift note',

            'Gift Message',
            'gift_message',
            '_gift_message',
            'gift-message',
            '_gift-message',
            'gift message',
            '_gift message',

            'gift_message_text',
            '_gift_message_text',
            'gift_note_text',
            '_gift_note_text',

            'billing_gift_note',
            '_billing_gift_note',
            'shipping_gift_note',
            '_shipping_gift_note',

            'billing_gift_message',
            '_billing_gift_message',
            'shipping_gift_message',
            '_shipping_gift_message',

            'order_gift_note',
            '_order_gift_note',
            'order_gift_message',
            '_order_gift_message',

            '_ywgc_gift_card_message',
            'ywgc_gift_card_message',
            'giftcard_message',
            'gift_card_message',
        ];

        foreach ($known_keys as $key) {
            $value = $order->get_meta($key, true);
            $value = $this->gift_note_value_to_text($value);

            if ($value !== '') {
                return $value;
            }
        }

        /*
         * Next, scan ALL order meta keys.
         *
         * This catches fields created by Checkout Field Editor and similar
         * plugins even when the saved key is something unexpected like:
         *     billing_custom_gift_message
         *     checkout_gift_note_1
         *     wc_order_gift_message
         */
        foreach ($order->get_meta_data() as $meta) {
            $data = $meta->get_data();

            $key = isset($data['key']) ? (string) $data['key'] : '';
            $value = isset($data['value']) ? $data['value'] : '';

            if (!$this->meta_key_looks_like_gift_note($key)) {
                continue;
            }

            $value = $this->gift_note_value_to_text($value);

            if ($value !== '') {
                return $value;
            }
        }

        /*
         * Last fallback: inspect line-item meta.
         * Some product-option / checkout plugins hide the Gift Note here.
         * WooCommerce likes a treasure hunt, apparently.
         */
        foreach ($order->get_items() as $item) {
            foreach ($item->get_meta_data() as $meta) {
                $data = $meta->get_data();

                $key = isset($data['key']) ? (string) $data['key'] : '';
                $value = isset($data['value']) ? $data['value'] : '';

                if (!$this->meta_key_looks_like_gift_note($key)) {
                    continue;
                }

                $value = $this->gift_note_value_to_text($value);

                if ($value !== '') {
                    return $value;
                }
            }
        }

        // Nothing found. Better an honest blank cell than inventing a love letter.
        return '';
    }

    /*
     * Decide whether an unknown meta key probably contains a Gift Note.
     *
     * We normalize punctuation/spaces/underscores first so keys such as
     * "_billing-gift_message" and "Billing Gift Message" are treated alike.
     */
    private function meta_key_looks_like_gift_note($key)
    {
        $normalized = strtolower((string) $key);
        $normalized = preg_replace('/[^a-z0-9]+/', ' ', $normalized);
        $normalized = trim(preg_replace('/\s+/', ' ', $normalized));

        if ($normalized === '') {
            return false;
        }

        $has_gift = strpos($normalized, 'gift') !== false;

        $has_note_or_message =
            strpos($normalized, 'note') !== false ||
            strpos($normalized, 'message') !== false;

        return $has_gift && $has_note_or_message;
    }

    /*
     * WooCommerce meta values are usually strings, but plugins sometimes save
     * arrays. This turns either form into something readable for Excel.
     */
    private function gift_note_value_to_text($value)
    {
        if (is_array($value)) {
            $parts = [];

            array_walk_recursive($value, function ($item) use (&$parts) {
                if (is_scalar($item)) {
                    $text = trim(wp_strip_all_tags((string) $item));

                    if ($text !== '') {
                        $parts[] = $text;
                    }
                }
            });

            return trim(implode(' / ', array_unique($parts)));
        }

        if (is_object($value)) {
            if (method_exists($value, '__toString')) {
                $value = (string) $value;
            } else {
                return '';
            }
        }

        if (!is_scalar($value)) {
            return '';
        }

        $value = html_entity_decode(
            wp_strip_all_tags((string) $value),
            ENT_QUOTES | ENT_HTML5,
            'UTF-8'
        );

        $value = preg_replace('/\s+/', ' ', $value);

        return trim($value);
    }

    /*
     * Remove anything after the first colon in a product title.
     * This mirrors the original processor's short-title behavior.
     */
    private function clean_item_title($title)
    {
        $title = (string) $title;
        $parts = explode(':', $title);
        return trim($parts[0]);
    }

    /*
     * Convert payment method into CityPak COD language:
     * Cash On Delivery = YES; everything else = NO.
     */
    private function payment_to_cod_value($payment_title)
    {
        $payment_title = trim((string) $payment_title);

        if (strcasecmp($payment_title, 'Cash On Delivery') === 0) {
            return 'YES';
        }

        return 'NO';
    }

    /*
     * Remove characters CityPak does not want.
     * Allowed: English letters, numbers, spaces, commas and slashes.
     */
    private function clean_citypak_text($value)
    {
        $text = (string) $value;
        $text = preg_replace('/[^A-Za-z0-9\s,\/]/', '', $text);
        $text = preg_replace('/\s+/', ' ', $text);
        return trim($text);
    }

    // ------------------------------------------------------------------------
    // 7) CITYPAK DESCRIPTION + SHORTENED-ORDER LOGIC
    // ------------------------------------------------------------------------

    /*
     * Decide whether the cleaned title list exceeds CityPak's description limit.
     * This is the fork in the road between "full titles" and "N BOOKS".
     */
    private function citypak_description_needs_shortening($item_name)
    {
        $cleaned = $this->clean_citypak_text($item_name);

        if ($cleaned === '') {
            return false;
        }

        $length = function_exists('mb_strlen')
            ? mb_strlen($cleaned, 'UTF-8')
            : strlen($cleaned);

        return $length > $this->citypak_description_max_chars;
    }

    /*
     * Produce the final CityPak description.
     * Full titles if they fit; otherwise a compact count such as "4 BOOKS".
     */
    private function make_citypak_description($item_name, $book_count = 1)
    {
        $cleaned = $this->clean_citypak_text($item_name);

        if ($cleaned === '') {
            return 'BOOKS';
        }

        if (!$this->citypak_description_needs_shortening($item_name)) {
            return $cleaned;
        }

        $book_count = max(1, (int) $book_count);

        return $book_count . ' BOOKS';
    }

    /*
     * Backup book counter for text such as:
     *     Book A / Book B X2 / Book C
     * which correctly becomes 4 books, not 3 titles.
     */
    private function count_books_from_item_description($item_name)
    {
        $parts = array_filter(array_map('trim', explode('/', (string) $item_name)));
        $count = 0;

        foreach ($parts as $part) {
            if (preg_match('/\\sX(\\d+)$/i', $part, $matches)) {
                $count += max(1, (int) $matches[1]);
            } else {
                $count++;
            }
        }

        return max(1, $count);
    }

    /*
     * Build Sheet 4's copy/paste messages ONLY for orders shortened in CityPak.
     *
     * Example:
     * #146030 13 BOOKS - Title A / Title B / Title C ...
     *
     * CityPak gets the short label; humans still get the full shopping list.
     */
    private function build_shortened_citypak_rows($processed_rows)
    {
        $rows = [];

        foreach ($processed_rows as $row) {
            $items = isset($row['Item Name']) ? (string) $row['Item Name'] : '';

            if (!$this->citypak_description_needs_shortening($items)) {
                continue;
            }

            $order_no = isset($row['Order Number'])
                ? trim((string) $row['Order Number'])
                : '';

            $book_count = isset($row['Book Count'])
                ? max(1, (int) $row['Book Count'])
                : $this->count_books_from_item_description($items);

            $display_order_no = ltrim($order_no, '#');

            /*
             * Use the unmodified processed Item Name here so the copy/paste
             * message keeps the actual titles and X2 / X3 quantities.
             */
            $message =
                '#' . $display_order_no .
                ' ' . $book_count . ' BOOKS - ' .
                $items;

            $rows[] = [
                'Order Number' => $order_no,
                'No of Books' => $book_count,
                'Message' => $message,
            ];
        }

        return $rows;
    }

    /*
     * Write Sheet 4: Shortened Orders.
     * Column C is intentionally wide and wrapped so the message is easy to copy
     * into WhatsApp, notes, customer-care tools, or wherever it needs to travel.
     */
    private function write_shortened_orders_sheet($sheet, $rows)
    {
        $headers = [
            'Order Number',
            'No of Books',
            'Message',
        ];

        $this->write_headers($sheet, $headers);

        $r = 2;

        foreach ($rows as $row) {
            $sheet->setCellValueExplicit(
                'A' . $r,
                (string) $row['Order Number'],
                DataType::TYPE_STRING
            );

            $sheet->setCellValue(
                'B' . $r,
                (int) $row['No of Books']
            );

            $sheet->setCellValueExplicit(
                'C' . $r,
                (string) $row['Message'],
                DataType::TYPE_STRING
            );

            $r++;
        }

        $last_row = max(2, $r - 1);

        $sheet->freezePane('A2');
        $sheet->setAutoFilter('A1:C' . $last_row);

        $sheet->getStyle('A1:C1')->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('FF7F6000');

        $sheet->getStyle('A1:C1')->getFont()
            ->setName('Arial')
            ->setSize(11)
            ->setBold(true)
            ->getColor()
            ->setARGB('FFFFFFFF');

        $sheet->getStyle('A1:C' . $last_row)->getBorders()->getAllBorders()
            ->setBorderStyle(Border::BORDER_THIN)
            ->getColor()
            ->setARGB('FFE7E6E6');

        $sheet->getStyle('A1:C' . $last_row)->getAlignment()
            ->setVertical(Alignment::VERTICAL_CENTER);

        $sheet->getStyle('C2:C' . $last_row)->getAlignment()
            ->setWrapText(false)
            ->setShrinkToFit(true);

        // Use 10 pt Arial for the message column. Excel can shrink the text
        // further automatically when a message is too long for the column.
        $sheet->getStyle('C2:C' . $last_row)->getFont()
            ->setName('Arial')
            ->setSize(10);

        $sheet->getStyle('B2:B' . $last_row)->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $sheet->getColumnDimension('A')->setWidth(18);
        $sheet->getColumnDimension('B')->setWidth(14);
        $sheet->getColumnDimension('C')->setWidth(120);

        $sheet->getRowDimension(1)->setRowHeight(24);

        if (empty($rows)) {
            $sheet->setCellValue(
                'A2',
                'No CityPak descriptions were shortened for this export.'
            );
            $sheet->mergeCells('A2:C2');
        }
    }

    /*
     * Return the first non-empty order meta value from a list of possible keys.
     * A small function doing a lot of peacekeeping between plugins.
     */
    private function get_order_meta_any($order, $keys)
    {
        foreach ($keys as $key) {
            $value = $order->get_meta($key);

            if ($value !== '' && $value !== null) {
                return $value;
            }
        }

        return '';
    }

    /*
     * Turn a WooCommerce status into a filename-safe label.
     */
    private function clean_filename_status($status)
    {
        $status = str_replace('wc-', '', $status);
        $status = preg_replace('/[^A-Za-z0-9_-]/', '-', $status);
        return trim($status, '-');
    }

    /*
     * Convert 1, 2, 3... into A, B, C...
     * Excel refuses to speak normal numbers for columns, so we translate.
     */
    private function column_letter($col_number)
    {
        $letter = '';

        while ($col_number > 0) {
            $mod = ($col_number - 1) % 26;
            $letter = chr(65 + $mod) . $letter;
            $col_number = (int)(($col_number - $mod) / 26);
        }

        return $letter;
    }

    // ------------------------------------------------------------------------
    // 8) LOW-LEVEL XLSX FINISHING & DOWNLOAD
    // ------------------------------------------------------------------------

    /*
     * Add Excel's REAL Conditional Formatting -> Color Scale to Shelving column C.
     *
     * Why the XML work?
     * The live/staging environments loaded incompatible PhpSpreadsheet conditional
     * formatting APIs, so this writes the standard XLSX rule directly.
     *
     * Also: Excel cares deeply about XML element order. Deeply.
     * Put this rule in the wrong place and Excel may "repair" the sheet into sadness.
     */
    private function apply_shelving_color_scale_to_xlsx($file_path, $last_row)
    {
        if (!class_exists('ZipArchive')) {
            throw new \RuntimeException(
                'The PHP ZipArchive extension is required to add the Excel color scale.'
            );
        }

        $last_row = max(2, (int) $last_row);

        $zip = new \ZipArchive();
        $opened = $zip->open($file_path);

        if ($opened !== true) {
            throw new \RuntimeException(
                'Could not open the generated Excel file for conditional formatting.'
            );
        }

        // The third worksheet created by Daily Dispatch is always Shelving List.
        $sheet_path = 'xl/worksheets/sheet3.xml';
        $xml = $zip->getFromName($sheet_path);

        if ($xml === false) {
            $zip->close();
            throw new \RuntimeException(
                'Could not find the Shelving List worksheet inside the generated Excel file.'
            );
        }

        $range = 'C2:C' . $last_row;

        /*
         * Real Excel:
         * Home -> Conditional Formatting -> Color Scales
         *
         * Low    = Red
         * Middle = Yellow
         * High   = Green
         */
        $conditional_xml =
            '<conditionalFormatting sqref="' . $range . '">' .
                '<cfRule type="colorScale" priority="1">' .
                    '<colorScale>' .
                        '<cfvo type="min"/>' .
                        '<cfvo type="percentile" val="50"/>' .
                        '<cfvo type="max"/>' .
                        '<color rgb="FFF8696B"/>' .
                        '<color rgb="FFFFEB84"/>' .
                        '<color rgb="FF63BE7B"/>' .
                    '</colorScale>' .
                '</cfRule>' .
            '</conditionalFormatting>';

        /*
         * IMPORTANT:
         * Excel requires conditionalFormatting to appear in the correct
         * worksheet XML order. It must be before printOptions/pageMargins/
         * pageSetup/headerFooter/tableParts.
         *
         * Putting it at the very end of the worksheet can cause Excel to
         * repair the worksheet and make it appear empty.
         */
        $insert_before_tags = [
            '<printOptions',
            '<pageMargins',
            '<pageSetup',
            '<headerFooter',
            '<rowBreaks',
            '<colBreaks',
            '<customProperties',
            '<cellWatches',
            '<ignoredErrors',
            '<smartTags',
            '<drawing',
            '<legacyDrawing',
            '<legacyDrawingHF',
            '<picture',
            '<oleObjects',
            '<controls',
            '<webPublishItems',
            '<tableParts',
            '<extLst',
            '</worksheet>',
        ];

        $insert_position = false;

        foreach ($insert_before_tags as $tag) {
            $position = strpos($xml, $tag);

            if ($position !== false) {
                if ($insert_position === false || $position < $insert_position) {
                    $insert_position = $position;
                }
            }
        }

        if ($insert_position === false) {
            $zip->close();
            throw new \RuntimeException(
                'Could not find a safe place to add the Excel conditional formatting rule.'
            );
        }

        $xml =
            substr($xml, 0, $insert_position) .
            $conditional_xml .
            substr($xml, $insert_position);

        if (!$zip->addFromString($sheet_path, $xml)) {
            $zip->close();
            throw new \RuntimeException(
                'Could not save the Excel conditional formatting rule.'
            );
        }

        $zip->close();
    }

    /*
     * Save the workbook to a temporary file, apply any final XLSX patching,
     * send it to the browser, then clean up after ourselves.
     * No abandoned temp files left behind like socks after laundry day.
     */
    private function download_spreadsheet($spreadsheet, $filename, $shelving_last_row = null)
    {
        $temp_file = wp_tempnam($filename);

        if (!$temp_file) {
            throw new \RuntimeException(
                'Could not create a temporary Excel file.'
            );
        }

        try {
            $writer = new Xlsx($spreadsheet);
            $writer->save($temp_file);

            // Only the processed workbook passes a shelving row count.
            // CityPak exports are not changed.
            if ($shelving_last_row !== null) {
                $this->apply_shelving_color_scale_to_xlsx(
                    $temp_file,
                    $shelving_last_row
                );
            }

            while (ob_get_level()) {
                ob_end_clean();
            }

            nocache_headers();

            header(
                'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            );
            header(
                'Content-Disposition: attachment; filename="' . $filename . '"'
            );
            header('Content-Length: ' . filesize($temp_file));
            header('Cache-Control: max-age=0');

            readfile($temp_file);
        } finally {
            if (file_exists($temp_file)) {
                @unlink($temp_file);
            }
        }

        exit;
    }
}