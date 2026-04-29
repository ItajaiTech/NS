<?php
// Teste rápido da função corrigida

function nsr_excel_column_to_index($column) {
    $column = strtoupper($column);
    $length = strlen($column);
    $index = 0;

    for ($i = 0; $i < $length; $i++) {
        $char_value = ord($column[$i]) - ord('A') + 1;
        $index = ($index * 26) + $char_value;
    }

    return $index - 1;
}

function nsr_row_has_any_value($row) {
    foreach ($row as $cell) {
        if (trim((string) $cell) !== '') {
            return true;
        }
    }
    return false;
}

$file = 'C:\NS\TECH.xlsx';
$zip = new ZipArchive();
$zip->open($file);
$sheet_content = $zip->getFromName('xl/worksheets/sheet1.xml');
$zip->close();

$xml = simplexml_load_string($sheet_content, 'SimpleXMLElement', LIBXML_NOCDATA);

if (!isset($xml->sheetData)) {
    die("Erro: sheetData não encontrado\n");
}

$rows = array();

foreach ($xml->sheetData->row as $row_node) {
    $row_values = array();

    foreach ($row_node->c as $cell) {
        $ref = (string) $cell['r'];
        $column = preg_replace('/\d+/', '', $ref);
        $index = nsr_excel_column_to_index($column);
        $type = (string) $cell['t'];
        $value = '';

        if ($type === 'inlineStr') {
            if (isset($cell->is->t)) {
                $value = (string) $cell->is->t;
            } elseif (isset($cell->is->r)) {
                foreach ($cell->is->r as $run) {
                    $value .= (string) $run->t;
                }
            }
        } else {
            if (isset($cell->v)) {
                $value = (string) $cell->v;
            }
        }

        $row_values[$index] = trim($value);
    }

    if (!empty($row_values) && nsr_row_has_any_value($row_values)) {
        ksort($row_values, SORT_NUMERIC);
        $rows[] = $row_values;
    }
    
    // Mostrar apenas primeiras 5 linhas
    if (count($rows) >= 5) {
        break;
    }
}

echo "Total de linhas lidas: " . count($rows) . "\n\n";

foreach ($rows as $i => $row) {
    echo "Linha $i:\n";
    foreach ($row as $col => $val) {
        if ($val !== '') {
            echo "  [$col] = $val\n";
        }
    }
    echo "\n";
}
