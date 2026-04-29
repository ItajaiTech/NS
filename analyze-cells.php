<?php
$file = 'C:\NS\TECH.xlsx';
$zip = new ZipArchive();
$zip->open($file);
$xml = simplexml_load_string($zip->getFromName('xl/worksheets/sheet1.xml'));
$zip->close();

// Encontrar primeira célula E com conteúdo
foreach ($xml->sheetData->row as $r) {
    foreach ($r->c as $c) {
        $ref = (string)$c['r'];
        
        if (strpos($ref, 'E') === 0) {
            // Verificar se tem inlineStr com múltiplas "runs"
            if (isset($c->is)) {
                $val = '';
                
                // Cada <r> (run) é um segmento de texto
                if (isset($c->is->r)) {
                    echo "Célula $ref tem " . count($c->is->r) . " runs (segmentos de texto)\n\n";
                    
                    $run_num = 0;
                    foreach ($c->is->r as $run) {
                        $run_num++;
                        $run_text = (string)$run->t;
                        echo "Run $run_num: ";
                        echo json_encode($run_text) . "\n";
                        $val .= $run_text;
                    }
                } elseif (isset($c->is->t)) {
                    echo "Célula $ref contém texto simples:\n";
                    echo json_encode((string)$c->is->t) . "\n";
                    $val = (string)$c->is->t;
                }
                
                // Mostrar valor completo com quebras visíveis
                echo "\nValor completo (quebras como \\n):\n";
                echo json_encode($val) . "\n";
                
                // Contar quantas linhas tem
                $lines = explode("\n", $val);
                echo "\nTotal de linhas: " . count($lines) . "\n";
                echo "Primeiras 10 linhas:\n";
                for ($i = 0; $i < min(10, count($lines)); $i++) {
                    $line = trim($lines[$i]);
                    if ($line !== '') {
                        echo "  [$i] $line\n";
                    }
                }
                
                return;
            }
        }
    }
}

echo "Nenhuma célula E com inlineStr encontrada com múltiplos valores\n";
