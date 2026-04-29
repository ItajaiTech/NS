<?php
$file = 'C:\NS\TECH.xlsx';
$zip = new ZipArchive();
$zip->open($file);
$xml = simplexml_load_string($zip->getFromName('xl/worksheets/sheet1.xml'));
$zip->close();

$count = 0;
$linha_num = 0;

foreach ($xml->sheetData->row as $r) {
    $linha_num++;
    foreach ($r->c as $c) {
        $ref = (string)$c['r'];
        
        // Verificar se é coluna E
        if (strpos($ref, 'E') === 0) {
            $val = '';
            
            if (isset($c->is->t)) {
                $val = (string)$c->is->t;
            } elseif (isset($c->v)) {
                $val = (string)$c->v;
            }
            
            if ($val !== '' && $val !== 'Observações internas') {
                echo "Linha $linha_num (ref: $ref): $val\n";
                $count++;
                
                if ($count >= 20) {
                    break 2;
                }
            }
        }
    }
}

echo "\nTotal de valores não vazios encontrados na coluna E: $count\n";
