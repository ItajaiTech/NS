<?php
/*
 * Plugin Name: NS Rastreio
 * Description: Importa planilhas Excel/CSV para consultar NS e encontrar numero da NF ou numero do pedido.
 * Version: 1.3.2
 * Author: Itajaitech
 */

if (!defined('ABSPATH')) {
    exit;
}

define('NSR_PLUGIN_VERSION', '1.4.0');
define('NSR_PLUGIN_SLUG', 'ns-rastreio');

/**
 * Extrai NSs individuais de uma celula que pode conter texto misto.
 *
 * Aceita NSs separados por virgula, ponto e virgula, quebra de linha
 * ou espacos, e ignora palavras sem numeros.
 *
 * @param string $value
 * @return array
 */
function nsr_extract_ns_values($value) {
    $value = strtoupper(remove_accents((string) $value));
    $value = trim($value);

    if ($value === '') {
        return array();
    }

    // Normaliza separadores comuns e remove simbolos mantendo apenas texto util.
    $value = preg_replace('/[\r\n\t,;|]+/', ' ', $value);
    $value = preg_replace('/\s+/', ' ', $value);

    $parts = preg_split('/\s+/', $value);
    $tokens = array();

    foreach ($parts as $part) {
        $part = preg_replace('/[^A-Z0-9]/', '', (string) $part);
        if ($part === '') {
            continue;
        }

        // NS precisa conter ao menos um digito e ter tamanho minimo.
        if (!preg_match('/\d/', $part)) {
            continue;
        }

        if (strlen($part) < 6) {
            continue;
        }

        $tokens[] = $part;
    }

    $tokens = array_values(array_unique($tokens));
    if (!empty($tokens)) {
        return $tokens;
    }

    $fallback = nsr_normalize_lookup_value($value);
    if ($fallback === '') {
        return array();
    }

    if (!preg_match('/\d/', $fallback)) {
        return array();
    }

    if (strlen($fallback) > 80) {
        return array();
    }

    return array($fallback);
}

/**
 * Retorna o nome da tabela usada pelo plugin.
 *
 * @return string
 */
function nsr_get_table_name() {
    global $wpdb;
    return $wpdb->prefix . 'ns_rastreio';
}

/**
 * Cria/atualiza tabela do plugin.
 */
function nsr_activate_plugin() {
    global $wpdb;

    $table_name = nsr_get_table_name();
    $charset_collate = $wpdb->get_charset_collate();

    $sql = "CREATE TABLE {$table_name} (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
        ns VARCHAR(255) NOT NULL,
        ns_normalizado VARCHAR(255) NOT NULL,
        nota_fiscal VARCHAR(100) DEFAULT '',
        pedido VARCHAR(100) DEFAULT '',
        sku VARCHAR(255) DEFAULT '',
        descricao TEXT DEFAULT '',
        quantidade VARCHAR(50) DEFAULT '',
        valor VARCHAR(50) DEFAULT '',
        data_venda VARCHAR(50) DEFAULT '',
        origem_arquivo VARCHAR(255) DEFAULT '',
        linha_origem INT UNSIGNED DEFAULT NULL,
        created_at DATETIME NOT NULL,
        updated_at DATETIME NOT NULL,
        PRIMARY KEY (id),
        UNIQUE KEY uniq_ns_nf_pedido (ns_normalizado, nota_fiscal, pedido),
        KEY idx_ns_normalizado (ns_normalizado),
        KEY idx_nota_fiscal (nota_fiscal),
        KEY idx_pedido (pedido)
    ) {$charset_collate};";

    require_once ABSPATH . 'wp-admin/includes/upgrade.php';
    dbDelta($sql);

    // Migracao de estrutura legada (versao 1.0.0 tinha unicidade por NS).
    $legacy_unique = $wpdb->get_var("SHOW INDEX FROM {$table_name} WHERE Key_name = 'uniq_ns_normalizado'");
    if (!empty($legacy_unique)) {
        $wpdb->query("ALTER TABLE {$table_name} DROP INDEX uniq_ns_normalizado");
    }

    $index_ns = $wpdb->get_var("SHOW INDEX FROM {$table_name} WHERE Key_name = 'idx_ns_normalizado'");
    if (empty($index_ns)) {
        $wpdb->query("ALTER TABLE {$table_name} ADD KEY idx_ns_normalizado (ns_normalizado)");
    }

    $index_combo = $wpdb->get_var("SHOW INDEX FROM {$table_name} WHERE Key_name = 'uniq_ns_nf_pedido'");
    if (empty($index_combo)) {
        $wpdb->query("ALTER TABLE {$table_name} ADD UNIQUE KEY uniq_ns_nf_pedido (ns_normalizado, nota_fiscal, pedido)");
    }

    update_option('nsr_db_version', NSR_PLUGIN_VERSION);
}
register_activation_hook(__FILE__, 'nsr_activate_plugin');

/**
 * Garante migracao de schema para installs ja ativos.
 */
function nsr_maybe_migrate_schema() {
    $installed = get_option('nsr_db_version', '0');
    if (version_compare($installed, NSR_PLUGIN_VERSION, '>=')) {
        return;
    }

    nsr_activate_plugin();
}
add_action('plugins_loaded', 'nsr_maybe_migrate_schema');

/**
 * Migra registros legados onde varios NSs foram salvos em uma unica linha.
 */
function nsr_maybe_split_legacy_compound_ns_rows() {
    global $wpdb;

    $migration_flag = get_option('nsr_compound_ns_migration_version', '0');
    if (version_compare($migration_flag, NSR_PLUGIN_VERSION, '>=')) {
        return;
    }

    $table_name = nsr_get_table_name();
    $table_exists = $wpdb->get_var($wpdb->prepare('SHOW TABLES LIKE %s', $table_name));
    if ($table_exists !== $table_name) {
        update_option('nsr_compound_ns_migration_version', NSR_PLUGIN_VERSION);
        return;
    }

    $rows = $wpdb->get_results(
        "SELECT id, ns, nota_fiscal, pedido, sku, descricao, quantidade, valor, data_venda, origem_arquivo, linha_origem
         FROM {$table_name}
         WHERE ns LIKE '% %' OR ns LIKE '%,%' OR ns LIKE '%;%'",
        ARRAY_A
    );

    if (empty($rows)) {
        update_option('nsr_compound_ns_migration_version', NSR_PLUGIN_VERSION);
        return;
    }

    foreach ($rows as $row) {
        $ns_values = nsr_extract_ns_values($row['ns']);
        if (count($ns_values) <= 1) {
            continue;
        }

        $inserted_any = false;
        foreach ($ns_values as $ns_value) {
            $ns_normalizado = nsr_normalize_lookup_value($ns_value);
            if ($ns_normalizado === '') {
                continue;
            }

            $sql = $wpdb->prepare(
                "INSERT INTO {$table_name}
                (ns, ns_normalizado, nota_fiscal, pedido, sku, descricao, quantidade, valor, data_venda, origem_arquivo, linha_origem, created_at, updated_at)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %d, NOW(), NOW())
                ON DUPLICATE KEY UPDATE
                    ns = VALUES(ns),
                    sku = VALUES(sku),
                    descricao = VALUES(descricao),
                    quantidade = VALUES(quantidade),
                    valor = VALUES(valor),
                    data_venda = VALUES(data_venda),
                    origem_arquivo = VALUES(origem_arquivo),
                    linha_origem = VALUES(linha_origem),
                    updated_at = NOW()",
                $ns_value,
                $ns_normalizado,
                (string) $row['nota_fiscal'],
                (string) $row['pedido'],
                (string) $row['sku'],
                (string) $row['descricao'],
                (string) $row['quantidade'],
                (string) $row['valor'],
                (string) $row['data_venda'],
                (string) $row['origem_arquivo'],
                (int) $row['linha_origem']
            );

            $affected = $wpdb->query($sql);
            if ($affected !== false) {
                $inserted_any = true;
            }
        }

        if ($inserted_any) {
            $wpdb->delete($table_name, array('id' => (int) $row['id']), array('%d'));
        }
    }

    update_option('nsr_compound_ns_migration_version', NSR_PLUGIN_VERSION);
}
add_action('plugins_loaded', 'nsr_maybe_split_legacy_compound_ns_rows', 30);
/**
 * Normaliza texto de cabecalho para facilitar matching.
 *
 * @param string $text
 * @return string
 */
function nsr_normalize_header_text($text) {
    $text = strtoupper(remove_accents((string) $text));
    $text = preg_replace('/[^A-Z0-9]+/', ' ', $text);
    $text = trim((string) preg_replace('/\s+/', ' ', $text));
    return $text;
}

/**
 * Normaliza valor para consulta e deduplicacao de NS.
 *
 * @param string $value
 * @return string
 */
function nsr_normalize_lookup_value($value) {
    $value = strtoupper(remove_accents((string) $value));
    $value = trim($value);
    $value = preg_replace('/[\s\-\._\/]+/', '', $value);
    return (string) $value;
}

/**
 * Limpa valor de celula para persistencia.
 *
 * @param mixed $value
 * @return string
 */
function nsr_clean_cell_value($value) {
    $value = trim((string) $value);

    if ($value === '') {
        return '';
    }

    // Corrige numeros vindos do Excel no formato 1234.0.
    if (preg_match('/^\d+\.0+$/', $value)) {
        $value = preg_replace('/\.0+$/', '', $value);
    }

    // Tenta converter notacao cientifica quando for inteiramente numerica.
    if (preg_match('/^\d+(\.\d+)?E\+\d+$/i', $value)) {
        $converted = sprintf('%.0f', (float) $value);
        if ($converted !== '0') {
            $value = $converted;
        }
    }

    return sanitize_text_field($value);
}

/**
 * Verifica se um row possui ao menos um valor preenchido.
 *
 * @param array $row
 * @return bool
 */
function nsr_row_has_any_value($row) {
    foreach ($row as $cell) {
        if (trim((string) $cell) !== '') {
            return true;
        }
    }
    return false;
}

/**
 * Converte letra de coluna Excel para indice base 0.
 *
 * @param string $column
 * @return int
 */
function nsr_excel_column_to_index($column) {
    $column = strtoupper($column);
    $length = strlen($column);
    $index = 0;

    for ($i = 0; $i < $length; $i++) {
        $char_value = ord($column[$i]) - 64;
        if ($char_value < 1 || $char_value > 26) {
            return 0;
        }
        $index = ($index * 26) + $char_value;
    }

    return $index - 1;
}

/**
 * Identifica delimitador mais provavel em CSV.
 *
 * @param string $line
 * @return string
 */
function nsr_detect_csv_delimiter($line) {
    $delimiters = array(';', ',', "\t");
    $best_delimiter = ';';
    $best_score = -1;

    foreach ($delimiters as $delimiter) {
        $score = substr_count($line, $delimiter);
        if ($score > $best_score) {
            $best_score = $score;
            $best_delimiter = $delimiter;
        }
    }

    return $best_delimiter;
}

/**
 * Le linhas de arquivo CSV.
 *
 * @param string $file_path
 * @return array|WP_Error
 */
function nsr_read_csv_rows($file_path) {
    $handle = fopen($file_path, 'rb');
    if (!$handle) {
        return new WP_Error('nsr_csv_open_error', 'Nao foi possivel abrir o arquivo CSV.');
    }

    $first_line = fgets($handle);
    if ($first_line === false) {
        fclose($handle);
        return array();
    }

    $delimiter = nsr_detect_csv_delimiter($first_line);
    rewind($handle);

    $rows = array();
    while (($data = fgetcsv($handle, 0, $delimiter)) !== false) {
        $row = array();
        foreach ($data as $cell) {
            $row[] = trim((string) $cell);
        }

        if (nsr_row_has_any_value($row)) {
            $rows[] = $row;
        }
    }

    fclose($handle);
    return $rows;
}

/**
 * Lista arquivos de planilha dentro de XLSX.
 *
 * @param ZipArchive $zip
 * @return array
 */
function nsr_xlsx_list_sheet_files($zip) {
    $sheet_files = array();

    for ($i = 0; $i < $zip->numFiles; $i++) {
        $stat = $zip->statIndex($i);
        if (!$stat || empty($stat['name'])) {
            continue;
        }

        $name = $stat['name'];
        if (preg_match('#^xl/worksheets/sheet\d+\.xml$#i', $name)) {
            $sheet_files[] = $name;
        }
    }

    natsort($sheet_files);
    return array_values($sheet_files);
}

/**
 * Le shared strings de arquivo XLSX.
 *
 * @param ZipArchive $zip
 * @return array
 */
function nsr_xlsx_read_shared_strings($zip) {
    $shared_strings = array();
    $xml_content = $zip->getFromName('xl/sharedStrings.xml');

    if ($xml_content === false) {
        return $shared_strings;
    }

    $xml = simplexml_load_string($xml_content, 'SimpleXMLElement', LIBXML_NOCDATA);
    if ($xml === false) {
        return $shared_strings;
    }

    $namespaces = $xml->getNamespaces(true);
    $main_ns = isset($namespaces['']) ? $namespaces[''] : 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    $root = $xml->children($main_ns);

    foreach ($root->si as $si) {
        if (isset($si->t)) {
            $shared_strings[] = (string) $si->t;
            continue;
        }

        $value = '';
        if (isset($si->r)) {
            foreach ($si->r as $run) {
                $value .= (string) $run->t;
            }
        }
        $shared_strings[] = $value;
    }

    return $shared_strings;
}

/**
 * Le linhas de arquivo XLSX (primeira planilha).
 *
 * @param string $file_path
 * @return array|WP_Error
 */
function nsr_read_xlsx_rows($file_path) {
    if (!class_exists('ZipArchive')) {
        return new WP_Error('nsr_zip_missing', 'Extensao ZipArchive nao disponivel no servidor PHP.');
    }

    $zip = new ZipArchive();
    if ($zip->open($file_path) !== true) {
        return new WP_Error('nsr_xlsx_open_error', 'Nao foi possivel abrir o arquivo XLSX.');
    }

    $shared_strings = nsr_xlsx_read_shared_strings($zip);
    $sheet_files = nsr_xlsx_list_sheet_files($zip);

    if (empty($sheet_files)) {
        $zip->close();
        return new WP_Error('nsr_xlsx_sheet_missing', 'Nenhuma planilha foi encontrada dentro do XLSX.');
    }

    $sheet_content = $zip->getFromName($sheet_files[0]);
    $zip->close();

    if ($sheet_content === false) {
        return new WP_Error('nsr_xlsx_sheet_read_error', 'Nao foi possivel ler a primeira planilha do XLSX.');
    }

    $xml = simplexml_load_string($sheet_content, 'SimpleXMLElement', LIBXML_NOCDATA);
    if ($xml === false) {
        return new WP_Error('nsr_xlsx_parse_error', 'Falha ao interpretar XML da planilha XLSX.');
    }

    if (!isset($xml->sheetData)) {
        return new WP_Error('nsr_xlsx_no_data', 'A planilha XLSX nao contem dados validos.');
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

            if ($type === 's') {
                $shared_index = isset($cell->v) ? (int) $cell->v : -1;
                $value = isset($shared_strings[$shared_index]) ? $shared_strings[$shared_index] : '';
            } elseif ($type === 'inlineStr') {
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
    }

    return $rows;
}

/**
 * Le linhas de XLSX/CSV.
 *
 * @param string $file_path
 * @param string $file_name
 * @return array|WP_Error
 */
function nsr_read_spreadsheet_rows($file_path, $file_name) {
    $extension = strtolower(pathinfo($file_name, PATHINFO_EXTENSION));

    if ($extension === 'csv') {
        return nsr_read_csv_rows($file_path);
    }

    if ($extension === 'xlsx') {
        return nsr_read_xlsx_rows($file_path);
    }

    return new WP_Error('nsr_invalid_extension', 'Formato de arquivo nao suportado. Use .xlsx ou .csv.');
}

/**
 * Verifica se cabecalho contem determinada frase.
 *
 * @param string $label
 * @param string $phrase
 * @return bool
 */
function nsr_header_contains_phrase($label, $phrase) {
    $label = ' ' . $label . ' ';
    $phrase = ' ' . $phrase . ' ';
    return strpos($label, $phrase) !== false;
}

/**
 * Detecta colunas necessarias no cabecalho.
 *
 * @param array $header_row
 * @return array
 */
function nsr_detect_columns($header_row) {
    $mapping = array(
        'ns' => null,
        'nf' => null,
        'pedido' => null,
        'sku' => null,
        'descricao' => null,
        'quantidade' => null,
        'valor' => null,
        'data_venda' => null,
    );

    foreach ($header_row as $index => $label_raw) {
        $label = nsr_normalize_header_text($label_raw);
        if ($label === '') {
            continue;
        }

        // NS vem de "Observacoes internas"
        if ($mapping['ns'] === null) {
            $ns_candidates = array(
                'OBSERVACOES INTERNAS',
                'OBSERVACAO INTERNA',
                'OBS INTERNAS',
                'OBS INTERNA',
                'OBSERVACOES',
            );

            foreach ($ns_candidates as $candidate) {
                if (nsr_header_contains_phrase($label, $candidate)) {
                    $mapping['ns'] = (int) $index;
                    break;
                }
            }
        }

        // NF: "Numero (Nota Fiscal)"
        if ($mapping['nf'] === null) {
            $nf_candidates = array(
                'NUMERO NOTA FISCAL',
                'NOTA FISCAL',
                'NUMERO NF',
                'NF',
                'NFE',
            );

            foreach ($nf_candidates as $candidate) {
                if (nsr_header_contains_phrase($label, $candidate)) {
                    $mapping['nf'] = (int) $index;
                    break;
                }
            }
        }

        // Pedido: "Numero" (generico, por isso verificado por ultimo)
        if ($mapping['pedido'] === null && $label === 'NUMERO') {
            $mapping['pedido'] = (int) $index;
        }

        // SKU: "Codigo (SKU)"
        if ($mapping['sku'] === null) {
            $sku_candidates = array(
                'CODIGO SKU',
                'SKU',
                'CODIGO',
            );

            foreach ($sku_candidates as $candidate) {
                if (nsr_header_contains_phrase($label, $candidate)) {
                    $mapping['sku'] = (int) $index;
                    break;
                }
            }
        }

        // Descricao: "Descricao do produto"
        if ($mapping['descricao'] === null) {
            $desc_candidates = array(
                'DESCRICAO DO PRODUTO',
                'DESCRICAO PRODUTO',
                'DESCRICAO',
                'PRODUTO',
            );

            foreach ($desc_candidates as $candidate) {
                if (nsr_header_contains_phrase($label, $candidate)) {
                    $mapping['descricao'] = (int) $index;
                    break;
                }
            }
        }

        // Quantidade: "Quantidade de produtos"
        if ($mapping['quantidade'] === null) {
            $qtd_candidates = array(
                'QUANTIDADE DE PRODUTOS',
                'QUANTIDADE PRODUTOS',
                'QUANTIDADE',
                'QTD',
            );

            foreach ($qtd_candidates as $candidate) {
                if (nsr_header_contains_phrase($label, $candidate)) {
                    $mapping['quantidade'] = (int) $index;
                    break;
                }
            }
        }

        // Valor: "Valor total da venda"
        if ($mapping['valor'] === null) {
            $val_candidates = array(
                'VALOR TOTAL DA VENDA',
                'VALOR TOTAL VENDA',
                'VALOR VENDA',
                'VALOR TOTAL',
                'VALOR',
            );

            foreach ($val_candidates as $candidate) {
                if (nsr_header_contains_phrase($label, $candidate)) {
                    $mapping['valor'] = (int) $index;
                    break;
                }
            }
        }

        // Data: "Data da venda"
        if ($mapping['data_venda'] === null) {
            $data_candidates = array(
                'DATA DA VENDA',
                'DATA VENDA',
                'DATA',
            );

            foreach ($data_candidates as $candidate) {
                if (nsr_header_contains_phrase($label, $candidate)) {
                    $mapping['data_venda'] = (int) $index;
                    break;
                }
            }
        }
    }

    return $mapping;
}

/**
 * Le valor de celula por indice de coluna.
 *
 * @param array $row
 * @param int|null $index
 * @return string
 */
function nsr_get_cell_by_index($row, $index) {
    if ($index === null) {
        return '';
    }

    if (!isset($row[$index])) {
        return '';
    }

    return nsr_clean_cell_value($row[$index]);
}

/**
 * Prepara ambiente para importacao longa sem estourar timeout do PHP.
 */
function nsr_prepare_long_import_runtime() {
    if (function_exists('ignore_user_abort')) {
        @ignore_user_abort(true);
    }

    if (function_exists('set_time_limit')) {
        @set_time_limit(0);
    }

    @ini_set('max_execution_time', '0');
}

/**
 * Importa uma planilha no banco.
 *
 * @param string $file_path
 * @param string $file_name
 * @return array|WP_Error
 */
function nsr_import_file($file_path, $file_name) {
    global $wpdb;

    nsr_prepare_long_import_runtime();

    $rows = nsr_read_spreadsheet_rows($file_path, $file_name);
    if (is_wp_error($rows)) {
        return $rows;
    }

    if (empty($rows)) {
        return new WP_Error('nsr_empty_file', 'Arquivo sem dados para importar.');
    }

    $header = $rows[0];
    $mapping = nsr_detect_columns($header);

    if ($mapping['ns'] === null) {
        return new WP_Error(
            'nsr_ns_column_missing',
            'Nao foi encontrada coluna de NS. Certifique-se de ter a coluna "Observacoes internas" com o numero de serie.'
        );
    }

    if ($mapping['nf'] === null && $mapping['pedido'] === null) {
        return new WP_Error(
            'nsr_nf_pedido_missing',
            'Nao foi encontrada coluna de NF ou Pedido. Inclua pelo menos uma delas no cabecalho.'
        );
    }

    $table_name = nsr_get_table_name();

    $inserted = 0;
    $updated = 0;
    $ignored = 0;
    $errors = 0;
    $batch_commit_every = 500;
    $rows_since_commit = 0;

    // Em bases grandes, commit por lote reduz custo de I/O e evita timeout.
    $wpdb->query('START TRANSACTION');

    for ($i = 1; $i < count($rows); $i++) {
        if (($i % 200) === 0 && function_exists('set_time_limit')) {
            @set_time_limit(30);
        }

        $row = $rows[$i];

        $ns_raw = nsr_get_cell_by_index($row, $mapping['ns']);
        $nf = nsr_get_cell_by_index($row, $mapping['nf']);
        $pedido = nsr_get_cell_by_index($row, $mapping['pedido']);
        $sku = nsr_get_cell_by_index($row, $mapping['sku']);
        $descricao = nsr_get_cell_by_index($row, $mapping['descricao']);
        $quantidade = nsr_get_cell_by_index($row, $mapping['quantidade']);
        $valor = nsr_get_cell_by_index($row, $mapping['valor']);
        $data_venda = nsr_get_cell_by_index($row, $mapping['data_venda']);

        if ($ns_raw === '') {
            $ignored++;
            continue;
        }

        if ($nf === '' && $pedido === '') {
            $ignored++;
            continue;
        }

        $ns_values = nsr_extract_ns_values($ns_raw);
        if (empty($ns_values)) {
            $ignored++;
            continue;
        }

        foreach ($ns_values as $ns) {

            $ns_normalizado = nsr_normalize_lookup_value($ns);
            if ($ns_normalizado === '') {
                $ignored++;
                continue;
            }

            $sql = $wpdb->prepare(
                "INSERT INTO {$table_name}
                (ns, ns_normalizado, nota_fiscal, pedido, sku, descricao, quantidade, valor, data_venda, origem_arquivo, linha_origem, created_at, updated_at)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %d, NOW(), NOW())
                ON DUPLICATE KEY UPDATE
                    ns = VALUES(ns),
                    sku = VALUES(sku),
                    descricao = VALUES(descricao),
                    quantidade = VALUES(quantidade),
                    valor = VALUES(valor),
                    data_venda = VALUES(data_venda),
                    origem_arquivo = VALUES(origem_arquivo),
                    linha_origem = VALUES(linha_origem),
                    updated_at = NOW()",
                $ns,
                $ns_normalizado,
                $nf,
                $pedido,
                $sku,
                $descricao,
                $quantidade,
                $valor,
                $data_venda,
                $file_name,
                ($i + 1)
            );

            $affected = $wpdb->query($sql);

            if ($affected === false) {
                $errors++;
                continue;
            }

            $rows_since_commit++;
            if ($rows_since_commit >= $batch_commit_every) {
                $wpdb->query('COMMIT');
                $wpdb->query('START TRANSACTION');
                $rows_since_commit = 0;
            }

            if ((int) $affected === 1) {
                $inserted++;
            } else {
                $updated++;
            }
        }
    }

    $wpdb->query('COMMIT');

    return array(
        'file' => $file_name,
        'inserted' => $inserted,
        'updated' => $updated,
        'ignored' => $ignored,
        'errors' => $errors,
        'processed_rows' => max(0, count($rows) - 1),
    );
}

/**
 * Busca registros por NS.
 *
 * @param string $ns
 * @param bool $partial
 * @param int $limit
 * @return array
 */
function nsr_find_by_ns($ns, $partial = false, $limit = 100) {
    global $wpdb;

    $ns_raw = trim((string) $ns);
    $ns_normalizado = nsr_normalize_lookup_value($ns);
    if ($ns_raw === '' && $ns_normalizado === '') {
        return array();
    }

    $table_name = nsr_get_table_name();
    $limit = max(1, min(absint($limit), 500));

    if ($partial) {
        if ($ns_normalizado === '') {
            $sql = $wpdb->prepare(
                "SELECT ns, nota_fiscal, pedido, sku, descricao, quantidade, valor, data_venda, origem_arquivo, updated_at
                 FROM {$table_name}
                 WHERE ns LIKE %s
                 ORDER BY updated_at DESC, id DESC
                 LIMIT %d",
                '%' . $wpdb->esc_like($ns_raw) . '%',
                $limit
            );

            return $wpdb->get_results($sql, ARRAY_A);
        }

        $sql = $wpdb->prepare(
            "SELECT ns, nota_fiscal, pedido, sku, descricao, quantidade, valor, data_venda, origem_arquivo, updated_at
             FROM {$table_name}
             WHERE ns_normalizado LIKE %s OR ns LIKE %s
             ORDER BY updated_at DESC, id DESC
             LIMIT %d",
            '%' . $wpdb->esc_like($ns_normalizado) . '%',
            '%' . $wpdb->esc_like($ns_raw) . '%',
            $limit
        );

        return $wpdb->get_results($sql, ARRAY_A);
    }

    $sql = $wpdb->prepare(
        "SELECT ns, nota_fiscal, pedido, sku, descricao, quantidade, valor, data_venda, origem_arquivo, updated_at
         FROM {$table_name}
         WHERE ns_normalizado = %s
         ORDER BY updated_at DESC, id DESC
         LIMIT %d",
        $ns_normalizado,
        $limit
    );

    return $wpdb->get_results($sql, ARRAY_A);
}

/**
 * Registra menu administrativo.
 */
function nsr_register_admin_menu() {
    add_menu_page(
        'NS Rastreio',
        'NS Rastreio',
        'manage_options',
        NSR_PLUGIN_SLUG,
        'nsr_render_admin_page',
        'dashicons-database-view',
        56
    );
}
add_action('admin_menu', 'nsr_register_admin_menu');

/**
 * Processa envio de importacao no admin.
 *
 * @return array
 */
function nsr_handle_import_submission() {
    $messages = array(
        'success' => array(),
        'error' => array(),
    );

    if (!isset($_POST['nsr_import_submit'])) {
        return $messages;
    }

    if (!current_user_can('manage_options')) {
        $messages['error'][] = 'Permissao insuficiente para importar planilhas.';
        return $messages;
    }

    check_admin_referer('nsr_import_files', 'nsr_import_nonce');

    if (empty($_FILES['nsr_files']) || !is_array($_FILES['nsr_files']['name'])) {
        $messages['error'][] = 'Nenhum arquivo foi enviado.';
        return $messages;
    }

    $allowed_extensions = array('xlsx', 'csv');
    $file_count = count($_FILES['nsr_files']['name']);
    $processed_files = 0;

    for ($i = 0; $i < $file_count; $i++) {
        $original_name = isset($_FILES['nsr_files']['name'][$i]) ? wp_unslash($_FILES['nsr_files']['name'][$i]) : '';
        $safe_name = sanitize_file_name((string) $original_name);
        $tmp_name = isset($_FILES['nsr_files']['tmp_name'][$i]) ? $_FILES['nsr_files']['tmp_name'][$i] : '';
        $error_code = isset($_FILES['nsr_files']['error'][$i]) ? (int) $_FILES['nsr_files']['error'][$i] : UPLOAD_ERR_NO_FILE;

        if ($safe_name === '') {
            continue;
        }

        $processed_files++;

        if ($error_code !== UPLOAD_ERR_OK) {
            $messages['error'][] = sprintf('Falha no upload do arquivo %s (codigo %d).', $safe_name, $error_code);
            continue;
        }

        $extension = strtolower(pathinfo($safe_name, PATHINFO_EXTENSION));
        if (!in_array($extension, $allowed_extensions, true)) {
            $messages['error'][] = sprintf('Arquivo %s ignorado: formato nao suportado (use .xlsx ou .csv).', $safe_name);
            continue;
        }

        $result = nsr_import_file($tmp_name, $safe_name);

        if (is_wp_error($result)) {
            $messages['error'][] = sprintf('Erro em %s: %s', $safe_name, $result->get_error_message());
            continue;
        }

        $messages['success'][] = sprintf(
            '%s: %d inserido(s), %d atualizado(s), %d ignorado(s), %d erro(s).',
            $safe_name,
            $result['inserted'],
            $result['updated'],
            $result['ignored'],
            $result['errors']
        );
    }

    if ($processed_files === 0) {
        $messages['error'][] = 'Nenhum arquivo valido foi selecionado.';
    }

    return $messages;
}

/**
 * Exporta toda a base para CSV no layout de importacao.
 */
function nsr_handle_export_csv() {
    if (!current_user_can('manage_options')) {
        wp_die('Permissao insuficiente para exportar dados.');
    }

    check_admin_referer('nsr_export_csv', 'nsr_export_nonce');

    @set_time_limit(0);

    $filename = 'ns-rastreio-export-' . gmdate('Ymd-His') . '.csv';

    nocache_headers();
    header('Content-Type: text/csv; charset=utf-8');
    header('Content-Disposition: attachment; filename=' . $filename);

    $output = fopen('php://output', 'w');
    if ($output === false) {
        wp_die('Nao foi possivel gerar o arquivo de exportacao.');
    }

    // BOM UTF-8 ajuda o Excel a abrir acentuacao corretamente.
    fwrite($output, "\xEF\xBB\xBF");

    fputcsv(
        $output,
        array(
            'Numero',
            'Numero (Nota Fiscal)',
            'Quantidade de produtos',
            'Valor total da venda',
            'Observacoes internas',
            'Codigo (SKU)',
            'Descricao do produto',
            'Data da venda',
        ),
        ';'
    );

    global $wpdb;
    $table_name = nsr_get_table_name();
    $batch_size = 2000;
    $offset = 0;

    while (true) {
        $rows = $wpdb->get_results(
            $wpdb->prepare(
                "SELECT pedido, nota_fiscal, quantidade, valor, ns, sku, descricao, data_venda
                 FROM {$table_name}
                 ORDER BY id ASC
                 LIMIT %d OFFSET %d",
                $batch_size,
                $offset
            ),
            ARRAY_A
        );

        if (empty($rows)) {
            break;
        }

        foreach ($rows as $row) {
            fputcsv(
                $output,
                array(
                    (string) $row['pedido'],
                    (string) $row['nota_fiscal'],
                    (string) $row['quantidade'],
                    (string) $row['valor'],
                    (string) $row['ns'],
                    (string) $row['sku'],
                    (string) $row['descricao'],
                    (string) $row['data_venda'],
                ),
                ';'
            );
        }

        $offset += $batch_size;
    }

    fclose($output);
    exit;
}
add_action('admin_post_nsr_export_csv', 'nsr_handle_export_csv');

/**
 * Renderiza pagina admin do plugin.
 */
function nsr_render_admin_page() {
    global $wpdb;

    $messages = nsr_handle_import_submission();

    foreach ($messages['success'] as $message) {
        add_settings_error('nsr_messages', 'nsr_success_' . wp_rand(), $message, 'updated');
    }
    foreach ($messages['error'] as $message) {
        add_settings_error('nsr_messages', 'nsr_error_' . wp_rand(), $message, 'error');
    }

    $table_name = nsr_get_table_name();
    $total_records = (int) $wpdb->get_var("SELECT COUNT(1) FROM {$table_name}");
    $total_ns_unicos = (int) $wpdb->get_var("SELECT COUNT(DISTINCT ns_normalizado) FROM {$table_name}");

    $admin_search_value = isset($_GET['nsr_admin_ns']) ? sanitize_text_field(wp_unslash($_GET['nsr_admin_ns'])) : '';
    $admin_is_partial = (isset($_GET['nsr_admin_partial']) && $_GET['nsr_admin_partial'] === '1');
    $admin_results = array();
    if ($admin_search_value !== '') {
        $admin_results = nsr_find_by_ns($admin_search_value, $admin_is_partial, 200);
    }
    ?>
    <div class="wrap">
        <h1>NS Rastreio</h1>
        <?php settings_errors('nsr_messages'); ?>

        <p>
            <strong>Total de registros na base:</strong> <?php echo esc_html((string) $total_records); ?>
            | <strong>NS unicos:</strong> <?php echo esc_html((string) $total_ns_unicos); ?>
        </p>

        <h2>1) Importar planilhas</h2>
        <p>Envie arquivos <code>.xlsx</code> ou <code>.csv</code> com cabecalho. Colunas obrigatorias: <code>Observacoes internas</code> (NS), <code>Numero (Nota Fiscal)</code> e <code>Numero</code> (Pedido).</p>
        <p>Colunas opcionais: <code>Codigo (SKU)</code>, <code>Descricao do produto</code>, <code>Quantidade de produtos</code>, <code>Valor total da venda</code>, <code>Data da venda</code>.</p>

        <form method="post" enctype="multipart/form-data" style="margin-bottom:24px;">
            <?php wp_nonce_field('nsr_import_files', 'nsr_import_nonce'); ?>
            <input type="file" name="nsr_files[]" multiple accept=".xlsx,.csv" required />
            <p>
                <button type="submit" name="nsr_import_submit" class="button button-primary">Importar Arquivos</button>
            </p>
        </form>

        <h2>2) Exportar planilha (migracao)</h2>
        <p>Baixe um <code>.csv</code> com todos os registros no mesmo layout de importacao do plugin (ideal para levar para outra hospedagem).</p>
        <form method="post" action="<?php echo esc_url(admin_url('admin-post.php')); ?>" style="margin-bottom:24px;">
            <input type="hidden" name="action" value="nsr_export_csv" />
            <?php wp_nonce_field('nsr_export_csv', 'nsr_export_nonce'); ?>
            <button type="submit" class="button button-secondary">Exportar CSV completo</button>
        </form>

        <h2>3) Teste rapido da consulta (admin)</h2>
        <form method="get" style="display:flex;gap:8px;align-items:center;max-width:760px;flex-wrap:wrap;">
            <input type="hidden" name="page" value="<?php echo esc_attr(NSR_PLUGIN_SLUG); ?>" />
            <input type="text" name="nsr_admin_ns" value="<?php echo esc_attr($admin_search_value); ?>" placeholder="Digite o numero de serie (NS)" style="flex:1;" />
            <label style="display:flex;align-items:center;gap:6px;">
                <input type="checkbox" name="nsr_admin_partial" value="1" <?php checked($admin_is_partial); ?> />
                Busca parcial
            </label>
            <button type="submit" class="button">Buscar NS</button>
        </form>

        <?php if ($admin_search_value !== '') : ?>
            <div style="margin-top:12px;">
                <?php if (!empty($admin_results)) : ?>
                    <p><strong><?php echo esc_html((string) count($admin_results)); ?></strong> resultado(s) encontrado(s).</p>
                    <table class="widefat striped" style="max-width:100%;font-size:13px;">
                        <thead>
                            <tr>
                                <th>NS</th>
                                <th>NF</th>
                                <th>Pedido</th>
                                <th>SKU</th>
                                <th>Descrição</th>
                                <th>Qtd</th>
                                <th>Valor</th>
                                <th>Data Venda</th>
                                <th>Arquivo</th>
                                <th>Atualizado</th>
                            </tr>
                        </thead>
                        <tbody>
                            <?php foreach ($admin_results as $admin_result) : ?>
                                <tr>
                                    <td><?php echo esc_html($admin_result['ns']); ?></td>
                                    <td><?php echo esc_html($admin_result['nota_fiscal']); ?></td>
                                    <td><?php echo esc_html($admin_result['pedido']); ?></td>
                                    <td><?php echo esc_html($admin_result['sku']); ?></td>
                                    <td><?php echo esc_html($admin_result['descricao']); ?></td>
                                    <td><?php echo esc_html($admin_result['quantidade']); ?></td>
                                    <td><?php echo esc_html($admin_result['valor']); ?></td>
                                    <td><?php echo esc_html($admin_result['data_venda']); ?></td>
                                    <td><?php echo esc_html($admin_result['origem_arquivo']); ?></td>
                                    <td><?php echo esc_html($admin_result['updated_at']); ?></td>
                                </tr>
                            <?php endforeach; ?>
                        </tbody>
                    </table>
                <?php else : ?>
                    <div style="padding:12px;background-color:#fff3cd;border:1px solid #ffc107;border-radius:4px;margin-top:8px;color:#856404;">
                        <strong>⚠ NS não encontrado</strong><br/>
                        Nenhum resultado encontrado para: <strong><?php echo esc_html($admin_search_value); ?></strong>
                        <br/><small style="display:block;margin-top:6px;">Verifique o valor digitado ou marque "Busca parcial" para procurar por trecho.</small>
                    </div>
                <?php endif; ?>
            </div>
        <?php endif; ?>

        <h2>4) Consulta no site (navegador)</h2>
        <p>Crie uma pagina no WordPress e use o shortcode: <code>[ns_rastreio_consulta]</code>.</p>
    </div>
    <?php
}

/**
 * Renderiza formulario de consulta no frontend.
 *
 * @return string
 */
function nsr_render_shortcode() {
    $ns_value = isset($_GET['nsr_ns']) ? sanitize_text_field(wp_unslash($_GET['nsr_ns'])) : '';
    $is_partial = (isset($_GET['nsr_partial']) && $_GET['nsr_partial'] === '1');
    $results = array();

    if ($ns_value !== '') {
        $results = nsr_find_by_ns($ns_value, $is_partial, 100);
    }

    ob_start();
    ?>
    <div class="nsr-widget" style="max-width:760px;padding:16px;border:1px solid #dcdcde;border-radius:8px;">
        <form method="get" style="display:flex;gap:8px;flex-wrap:wrap;">
            <input
                type="text"
                name="nsr_ns"
                value="<?php echo esc_attr($ns_value); ?>"
                placeholder="Digite o numero de serie (NS)"
                style="flex:1;min-width:220px;padding:10px;"
                required
            />
            <label style="display:flex;align-items:center;gap:6px;">
                <input type="checkbox" name="nsr_partial" value="1" <?php checked($is_partial); ?> />
                Busca parcial
            </label>
            <button type="submit" style="padding:10px 16px;cursor:pointer;">Consultar</button>
        </form>

        <?php if ($ns_value !== '') : ?>
            <div style="margin-top:16px;">
                <?php if (!empty($results)) : ?>
                    <p><strong><?php echo esc_html((string) count($results)); ?></strong> resultado(s) encontrado(s).</p>
                    <table style="width:100%;border-collapse:collapse;font-size:14px;">
                        <thead>
                            <tr>
                                <th style="text-align:left;border-bottom:2px solid #ddd;padding:8px;">NS</th>
                                <th style="text-align:left;border-bottom:2px solid #ddd;padding:8px;">NF</th>
                                <th style="text-align:left;border-bottom:2px solid #ddd;padding:8px;">Pedido</th>
                                <th style="text-align:left;border-bottom:2px solid #ddd;padding:8px;">SKU</th>
                                <th style="text-align:left;border-bottom:2px solid #ddd;padding:8px;">Descrição</th>
                                <th style="text-align:left;border-bottom:2px solid #ddd;padding:8px;">Qtd</th>
                                <th style="text-align:left;border-bottom:2px solid #ddd;padding:8px;">Valor</th>
                                <th style="text-align:left;border-bottom:2px solid #ddd;padding:8px;">Data</th>
                            </tr>
                        </thead>
                        <tbody>
                            <?php foreach ($results as $row) : ?>
                                <tr>
                                    <td style="padding:8px;border-bottom:1px solid #eee;"><?php echo esc_html($row['ns']); ?></td>
                                    <td style="padding:8px;border-bottom:1px solid #eee;"><?php echo esc_html($row['nota_fiscal']); ?></td>
                                    <td style="padding:8px;border-bottom:1px solid #eee;"><?php echo esc_html($row['pedido']); ?></td>
                                    <td style="padding:8px;border-bottom:1px solid #eee;"><?php echo esc_html($row['sku']); ?></td>
                                    <td style="padding:8px;border-bottom:1px solid #eee;"><?php echo esc_html($row['descricao']); ?></td>
                                    <td style="padding:8px;border-bottom:1px solid #eee;"><?php echo esc_html($row['quantidade']); ?></td>
                                    <td style="padding:8px;border-bottom:1px solid #eee;"><?php echo esc_html($row['valor']); ?></td>
                                    <td style="padding:8px;border-bottom:1px solid #eee;"><?php echo esc_html($row['data_venda']); ?></td>
                                </tr>
                            <?php endforeach; ?>
                        </tbody>
                    </table>
                <?php else : ?>
                    <div style="padding:12px;background-color:#fff3cd;border:1px solid #ffc107;border-radius:4px;margin-top:8px;color:#856404;">
                        <strong>⚠ NS não encontrado</strong><br/>
                        Nenhum registro encontrado para: <strong><?php echo esc_html($ns_value); ?></strong>
                        <br/><small style="display:block;margin-top:6px;">Verifique se digitou corretamente ou tente marcar "Busca parcial".</small>
                    </div>
                <?php endif; ?>
            </div>
        <?php endif; ?>
    </div>
    <?php

    return (string) ob_get_clean();
}
add_shortcode('ns_rastreio_consulta', 'nsr_render_shortcode');
add_shortcode('ns_rastreio', 'nsr_render_shortcode');
