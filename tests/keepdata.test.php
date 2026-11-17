<?php
define('ABSPATH', __DIR__);
$saved_payload = null;
$last_request = null;
class WP_Error {
    private $message;
    public function __construct($code, $message) { $this->message = $message; }
    public function get_error_message() { return $this->message; }
}
function is_wp_error($value) { return $value instanceof WP_Error; }
function get_option($key, $default = '') {
    if ($key === 'nsr_keepdata_api_url') return 'https://api.keepdata.net.br:8765';
    if ($key === 'nsr_keepdata_api_token') return 'test-token';
    return $default;
}
function untrailingslashit($value) { return rtrim($value, '/'); }
function wp_remote_post($url, $args) {
    global $last_request;
    $last_request = array($url, $args);
    return array('response' => array('code' => 200), 'body' => json_encode(array(
        'success' => true, 'found' => true, 'serial_number' => '2091352',
        'document_number' => '12981', 'document_date' => '2025-09-23',
        'product_code' => '3372',
        'product_description' => 'SSD 2.5 SATA3 240GB KEEPDATA KDS240G-L21 500/550',
    )));
}
function wp_remote_retrieve_response_code($response) { return $response['response']['code']; }
function wp_remote_retrieve_body($response) { return $response['body']; }
function wp_json_encode($value) { return json_encode($value); }
function sanitize_text_field($value) { return trim(strip_tags($value)); }
function nsr_normalize_lookup_value($value) { return preg_replace('/\D+/', '', (string) $value); }
function nsr_upsert_ns_record($payload) { global $saved_payload; $saved_payload = $payload; return 1; }

require_once __DIR__ . '/../ns-rastreio-plugin/includes/keepdata.php';

function check($condition, $message) {
    if (!$condition) { fwrite(STDERR, "FAIL: {$message}\n"); exit(1); }
}

$result = nsr_keepdata_fetch_and_cache('2091352');
check(is_array($result), 'Result is mapped');
check($saved_payload['ns'] === '2091352', 'Serial mapped');
check($saved_payload['nota_fiscal'] === '12981', 'Invoice mapped');
check($saved_payload['pedido'] === '', 'Order stays empty');
check($saved_payload['sku'] === '3372', 'Product code maps to SKU');
check(strpos($saved_payload['descricao'], 'SSD 2.5') === 0, 'Description mapped');
check($saved_payload['data_venda'] === '2025-09-23', 'Document date maps to sale date');
check($saved_payload['origem_arquivo'] === 'keepdata-api', 'Origin identifies API cache');
check(nsr_record_origin_label($saved_payload) === 'PY', 'Keepdata cache is labeled PY');
check(nsr_record_origin_label(array('origem_arquivo' => 'importacao.csv')) === 'BR', 'Local import is labeled BR');
check($last_request[0] === 'https://api.keepdata.net.br:8765/v1/garantias/consultar', 'Correct endpoint');
check($last_request[1]['headers']['Authorization'] === 'Bearer test-token', 'Bearer header used');
echo "Keepdata cache: all checks passed\n";
