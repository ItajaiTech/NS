<?php
$file = 'C:\NS\TECH.xlsx';

$zip = new ZipArchive();
$zip->open($file);
$sheet_content = $zip->getFromName('xl/worksheets/sheet1.xml');
$zip->close();

$xml = simplexml_load_string($sheet_content, 'SimpleXMLElement', LIBXML_NOCDATA);

// Testar acesso direto aos elementos
echo "Teste 1 - Acesso direto sem namespace:\n";
if (isset($xml->sheetData->row[0]->c[0])) {
    $cell = $xml->sheetData->row[0]->c[0];
    echo "Tipo: " . $cell['t'] . "\n";
    echo "Ref: " . $cell['r'] . "\n";
    
    if (isset($cell->is->t)) {
        echo "Valor (is->t): " . $cell->is->t . "\n";
    } else {
        echo "is->t não encontrado\n";
        echo "Estrutura da célula:\n";
        print_r($cell);
    }
}

echo "\n\nTeste 2 - Com namespace:\n";
$namespaces = $xml->getNamespaces(true);
print_r($namespaces);

$main_ns = isset($namespaces['']) ? $namespaces[''] : 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
echo "\nUsando namespace: $main_ns\n";

$root = $xml->children($main_ns);
if (isset($root->sheetData->row[0]->c[0])) {
    $cell = $root->sheetData->row[0]->c[0];
    echo "Tipo: " . $cell['t'] . "\n";
    
    // Tentar sem namespace
    if (isset($cell->is)) {
        echo "is encontrado sem namespace\n";
        echo "Valor: " . $cell->is->t . "\n";
    }
    
    // Tentar com namespace
    $is_with_ns = $cell->children($main_ns)->is;
    if ($is_with_ns) {
        echo "is encontrado COM namespace\n";
        $t_with_ns = $is_with_ns->children($main_ns)->t;
        echo "Valor: " . $t_with_ns . "\n";
    }
}
