<?php
define('ABSPATH', __DIR__);
define('ARRAY_A', 'ARRAY_A');
function nsr_get_table_name() { return 'wp_ns_rastreio'; }
function remove_accents($text) { return $text; }
class WP_Error {
    public function __construct($code, $message) { $this->message = $message; }
    public function get_error_message() { return $this->message; }
}
function check($condition, $message) {
    if (!$condition) { throw new RuntimeException($message); }
}
$source = file_get_contents(__DIR__ . '/../ns-rastreio-plugin/ns-rastreio-plugin.php');
preg_match('/function nsr_normalize_lookup_value\([^}]+\}/s', $source, $match);
eval($match[0]);
require __DIR__ . '/../ns-rastreio-plugin/includes/batch-search.php';

$serials = nsr_parse_batch_serials("00123\r\nAB-12; ab12,XYZ\t00123|9");
check(count($serials) === 4, 'Deduplicates normalized values, accepts short and alphabetic NS');
check(reset($serials) === '00123', 'Preserves leading zeros and input order');
check(nsr_parse_batch_serials('') === array(), 'Empty list');
foreach (array('06/10/2024' => '06/10/2026', '2024-10-06' => '06/10/2026',
    '29/02/2024' => '28/02/2026', '2024-01-31 12:34:56' => '31/01/2026') as $input => $expected) {
    check(nsr_batch_expiration($input) === $expected, 'Expiration for ' . $input);
}
foreach (array('', '31/02/2024', 'garbage') as $input) {
    check(nsr_batch_expiration($input) === 'Data de venda ausente ou invalida', 'Reject invalid date');
}
$wpdb = new class {
    public $last_error = '';
    public $calls = 0;
    public $args;
    public function prepare($sql, $args) {
        check(strpos($sql, 'IN (') !== false, 'Exact prepared batch query');
        $this->args = $args;
        return $sql;
    }
    public function get_results($sql, $format) {
        $this->calls++;
        if ($this->last_error !== '') { return null; }
        return in_array('00123', $this->args, true) ? array(
            array('ns_normalizado' => '00123', 'nota_fiscal' => 'NF1'),
            array('ns_normalizado' => '00123', 'nota_fiscal' => 'NF2')
        ) : array();
    }
};
$groups = nsr_find_batch_serials($serials);
check(count($groups['00123']) === 2, 'Keeps multiple invoices for the same NS');
check($groups['XYZ'] === array(), 'Reports missing serials');
check(array_keys($groups) === array_keys($serials), 'Preserves pasted order');
check($wpdb->calls === 1, 'Queries list together');
$many = array();
for ($i = 0; $i < 201; $i++) { $many['NS' . $i] = 'NS' . $i; }
$wpdb->calls = 0;
nsr_find_batch_serials($many);
check($wpdb->calls === 3, 'Chunks large lists');
$wpdb->last_error = 'database unavailable';
check(nsr_find_batch_serials($serials) instanceof WP_Error, 'Database error is not a missing NS');
echo "Batch search: all checks passed\n";
