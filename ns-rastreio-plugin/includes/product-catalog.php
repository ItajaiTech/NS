<?php
defined('ABSPATH') || exit;

/** The RMA catalog is authoritative when installed in this WordPress database. */
function nsr_rma_products_available() {
    global $wpdb;
    static $available = null;
    if ($available === null) {
        $table = $wpdb->prefix . 'rma_produtos';
        $available = $wpdb->get_var($wpdb->prepare('SHOW TABLES LIKE %s', $wpdb->esc_like($table))) === $table;
        if ($available) {
            $columns = $wpdb->get_col("SHOW COLUMNS FROM {$table}");
            $available = !array_diff(array('sku', 'nome', 'gtin'), (array) $columns);
        }
    }
    return $available;
}

function nsr_catalog_sku_key($sku) {
    return strtoupper(trim((string) $sku));
}

/** SKU equality retains spaces, punctuation and leading zeros. No fuzzy linking. */
function nsr_get_catalog_products($skus) {
    global $wpdb;
    $skus = array_values(array_unique(array_filter(array_map('nsr_catalog_sku_key', $skus), function ($sku) { return $sku !== ''; })));
    $rma = nsr_rma_products_available();
    $table = $rma ? $wpdb->prefix . 'rma_produtos' : nsr_get_products_table_name();
    $fields = $rma ? 'sku, nome AS descricao, gtin' : "sku, descricao, '' AS gtin";
    $map = array();
    foreach (array_chunk($skus, 100) as $chunk) {
        $placeholders = implode(',', array_fill(0, count($chunk), '%s'));
        $rows = $wpdb->get_results($wpdb->prepare("SELECT {$fields} FROM {$table} WHERE sku IN ({$placeholders})", $chunk), ARRAY_A);
        foreach ((array) $rows as $row) {
            $key = nsr_catalog_sku_key($row['sku']);
            if (in_array($key, $chunk, true)) {
                $map[$key] = $row;
            }
        }
    }
    return $map;
}

/** Refresh product data for query results without changing invoice history. */
function nsr_enrich_product_records($rows) {
    if (!$rows) { return $rows; }
    $products = nsr_get_catalog_products(array_column($rows, 'sku'));
    foreach ($rows as &$row) {
        $key = nsr_catalog_sku_key($row['sku'] ?? '');
        if (isset($products[$key])) {
            if ($products[$key]['descricao'] !== '') {
                $row['descricao'] = $products[$key]['descricao'];
            }
            $row['gtin'] = $products[$key]['gtin'];
        } else {
            $row['gtin'] = '';
        }
    }
    unset($row);
    return $rows;
}
