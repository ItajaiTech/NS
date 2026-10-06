<?php
define('ABSPATH', __DIR__);
function remove_accents($text) { return $text; }
function check($value, $message) { if (!$value) { throw new RuntimeException($message); } }
$source = file_get_contents(__DIR__ . '/../ns-rastreio-plugin/ns-rastreio-plugin.php');
foreach (array('nsr_normalize_header_text', 'nsr_header_contains_phrase', 'nsr_detect_columns') as $name) {
    $start = strpos($source, 'function ' . $name . '(');
    // Extract only the function, stopping at the next documentation block.
    $end = strpos($source, "\n/**", $start);
    eval(substr($source, $start, $end - $start));
}
$new = nsr_detect_columns(array('NS', 'Nota fiscal', 'Pedido', 'SKU', 'Descricao', 'Data de venda', 'GTIN/EAN', 'Expiracao', 'Quantidade', 'Valor'));
check($new === array('ns'=>0,'nf'=>1,'pedido'=>2,'sku'=>3,'descricao'=>4,'quantidade'=>8,'valor'=>9,'data_venda'=>5), 'New layout mapping');
$old = nsr_detect_columns(array('Numero','Numero (Nota Fiscal)','Quantidade de produtos','Valor total da venda','Observacoes internas','Codigo (SKU)','Descricao do produto','Data da venda'));
check($old['ns'] === 4 && $old['nf'] === 1 && $old['pedido'] === 0 && $old['data_venda'] === 7, 'Legacy layout mapping');
$screen = nsr_detect_columns(array('NS consultado','Resultado','NS encontrado','Nota fiscal','Pedido','SKU','Descricao','GTIN/EAN','Expiracao','Data da compra'));
check($screen['ns'] === 2 && $screen['data_venda'] === 9, 'Found NS and purchase date mapping');
$dates = nsr_detect_columns(array('Expiracao','Data da nota fiscal'));
check($dates['data_venda'] === 1, 'Expiration is not purchase date');
require __DIR__ . '/../ns-rastreio-plugin/includes/batch-search.php';
check(nsr_format_sale_date('2026-06-09') === '09/06/2026', 'ISO date formatting');
check(nsr_format_sale_date('09/06/2026') === '09/06/2026', 'Brazilian date formatting');
check(nsr_format_sale_date('31/02/2026') === '31/02/2026', 'Invalid source retained without invented date');
echo "Import/export layout: all checks passed\n";
