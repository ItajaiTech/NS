<?php
$file = 'C:\NS\TECH.xlsx';
$zip = new ZipArchive();
$zip->open($file);
$xml = simplexml_load_string($zip->getFromName('xl/worksheets/sheet1.xml'));
$zip->close();

// Procurar célula E com muitos runs
foreach ($xml->sheetData->row as $r) {
    foreach ($r->c as $c) {
        $ref = (string)$c['r'];
        
        if (strpos($ref, 'E') === 0 && isset($c->is) && isset($c->is->r)) {
            $run_count = count($c->is->r);
            
            // Se tem muitos runs, provavelmente são múltiplos NSs
            if ($run_count > 5) {
                echo "Célula $ref contém $run_count runs!\n\n";
                
                $valor_completo = '';
                foreach ($c->is->r as $idx => $run) {
                    $text = (string)$run->t;
                    $valor_completo .= $text;
                }
                
                // Mostrar estrutura
                echo "Valor completo (JSON com quebras visíveis):\n";
                echo json_encode($valor_completo) . "\n\n";
                
                // Quebrar por linhas
                $lines = explode("\n", $valor_completo);
                echo "Total de linhas: " . count($lines) . "\n";
                echo "Primeiras 20 linhas:\n";
                
                $count = 0;
                foreach ($lines as $idx => $line) {
                    $line_trimmed = trim($line);
                    if ($line_trimmed !== '') {
                        echo "  [$idx] $line_trimmed\n";
                        $count++;
                        if ($count >= 20) break;
                    }
                }
                
                return;
            }
        }
    }
}

echo "Nenhuma célula E com múltiplos runs encontrada\n";
