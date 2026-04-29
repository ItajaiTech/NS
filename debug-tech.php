<?php
// Debug script para TECH.xlsx

$file = 'C:\NS\TECH.xlsx';

if (!file_exists($file)) {
    die("Arquivo não encontrado: $file\n");
}

echo "Arquivo encontrado: $file\n";
echo "Tamanho: " . filesize($file) . " bytes\n\n";

// Tentar abrir o ZIP
$zip = new ZipArchive();
$result = $zip->open($file);

if ($result !== true) {
    die("Erro ao abrir ZIP. Código: $result\n");
}

echo "ZIP aberto com sucesso!\n";
echo "Número de arquivos no ZIP: " . $zip->numFiles . "\n\n";

// Listar arquivos dentro do ZIP
echo "Arquivos dentro do XLSX:\n";
for ($i = 0; $i < $zip->numFiles; $i++) {
    $stat = $zip->statIndex($i);
    echo "  " . $stat['name'] . " (" . $stat['size'] . " bytes)\n";
}

// Procurar sheets
echo "\n\nProcurando sheets (xl/worksheets/):\n";
$sheet_files = array();
for ($i = 0; $i < $zip->numFiles; $i++) {
    $name = $zip->getNameIndex($i);
    if (stripos($name, 'xl/worksheets/sheet') !== false && substr($name, -4) === '.xml') {
        $sheet_files[] = $name;
        echo "  Encontrado: $name\n";
    }
}

if (empty($sheet_files)) {
    die("\nNENHUM SHEET ENCONTRADO!\n");
}

// Ler primeira sheet
$first_sheet = $sheet_files[0];
echo "\n\nLendo primeira sheet: $first_sheet\n";

$xml_content = $zip->getFromName($first_sheet);
if ($xml_content === false) {
    die("Erro ao ler conteúdo do sheet\n");
}

echo "Conteúdo XML (primeiros 1000 chars):\n";
echo substr($xml_content, 0, 1000) . "\n\n";

// Parse XML
libxml_use_internal_errors(true);
$xml = simplexml_load_string($xml_content);

if ($xml === false) {
    echo "Erro ao fazer parse do XML:\n";
    foreach (libxml_get_errors() as $error) {
        echo "  " . $error->message;
    }
    libxml_clear_errors();
    die();
}

echo "XML parseado com sucesso!\n";

// Contar linhas
$row_count = 0;
if (isset($xml->sheetData) && isset($xml->sheetData->row)) {
    $row_count = count($xml->sheetData->row);
}

echo "Número de linhas encontradas: $row_count\n\n";

// Mostrar primeiras 5 linhas
if ($row_count > 0) {
    echo "Primeiras linhas (estrutura):\n";
    $counter = 0;
    foreach ($xml->sheetData->row as $row) {
        echo "Linha " . $row['r'] . ":\n";
        if (isset($row->c)) {
            foreach ($row->c as $cell) {
                $ref = (string)$cell['r'];
                $type = isset($cell['t']) ? (string)$cell['t'] : '';
                $value = isset($cell->v) ? (string)$cell->v : '';
                echo "  Cell $ref (type: $type): $value\n";
            }
        }
        echo "\n";
        
        $counter++;
        if ($counter >= 5) break;
    }
}

$zip->close();
echo "\nDiagnóstico concluído!\n";
