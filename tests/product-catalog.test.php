<?php
define('ABSPATH', __DIR__);
define('ARRAY_A', 'ARRAY_A');
function nsr_get_products_table_name() { return 'wp_nsr_produtos'; }
function check($value, $message) { if (!$value) { throw new RuntimeException($message); } }
$rma = !in_array('--local', $argv, true);
$wpdb = new class($rma) {
    public $prefix = 'wp_';
    public $rma;
    public $queries = array();
    public function __construct($rma) { $this->rma = $rma; }
    public function esc_like($value) { return $value; }
    public function prepare($sql, $args) { return $sql; }
    public function get_var($sql) { return $this->rma ? 'wp_rma_produtos' : null; }
    public function get_col($sql) { return array('sku', 'nome', 'gtin'); }
    public function get_results($sql, $format) {
        $this->queries[] = $sql;
        return array(array('sku' => 'PRD00070', 'descricao' => $this->rma ? 'Nome do RMA' : 'Nome local', 'gtin' => $this->rma ? '7891234567890' : ''),
            array('sku' => 'CPU I3 2120', 'descricao' => 'CPU', 'gtin' => ''),
            array('sku' => '00123', 'descricao' => 'Produto numerico', 'gtin' => ''));
    }
};
require __DIR__ . '/../ns-rastreio-plugin/includes/product-catalog.php';
check(nsr_rma_products_available() === $rma, 'Detect RMA availability');
$map = nsr_get_catalog_products(array(' prd00070 ', 'CPU I3 2120', '00123', '123', 'missing'));
check(isset($map['PRD00070'], $map['CPU I3 2120'], $map['00123']), 'SKU matching preserves spaces and zeroes');
check(!isset($map['123'], $map['MISSING']), 'No approximate matches');
check(strpos($wpdb->queries[0], $rma ? 'wp_rma_produtos' : 'wp_nsr_produtos') !== false, 'Correct catalog');
check(strpos($wpdb->queries[0], $rma ? 'nome AS descricao' : 'SELECT sku, descricao') !== false, 'Correct field mapping');
$rows = nsr_enrich_product_records(array(
    array('sku' => 'prd00070', 'descricao' => 'Descricao antiga', 'nota_fiscal' => '3868', 'data_venda' => '2026-06-09'),
    array('sku' => 'MISSING', 'descricao' => 'Descricao da nota'),
    array('sku' => '', 'descricao' => 'Sem SKU')));
check($rows[0]['descricao'] === ($rma ? 'Nome do RMA' : 'Nome local'), 'Current catalog name');
check($rows[0]['gtin'] === ($rma ? '7891234567890' : ''), 'GTIN from catalog');
check($rows[0]['nota_fiscal'] === '3868' && $rows[0]['data_venda'] === '2026-06-09', 'Invoice history preserved');
check($rows[1]['descricao'] === 'Descricao da nota' && $rows[2]['descricao'] === 'Sem SKU', 'Missing SKU preserves existing description');
check(nsr_enrich_product_records(array()) === array(), 'Empty results');
echo 'Product catalog (' . ($rma ? 'RMA' : 'local') . "): all checks passed\n";
