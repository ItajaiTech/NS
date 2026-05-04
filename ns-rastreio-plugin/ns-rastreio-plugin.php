<?php
/*
 * Plugin Name: NS Rastreio
 * Description: Importa planilhas Excel/CSV para consultar NS e encontrar numero da NF ou numero do pedido.
 * Version: 1.5.0
 * Author: Itajaitech
 */

if (!defined('ABSPATH')) {
    exit;
}

define('NSR_PLUGIN_VERSION', '1.5.0');
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
 * Retorna o nome da tabela de produtos cadastrados.
 *
 * @return string
 */
function nsr_get_products_table_name() {
    global $wpdb;
    return $wpdb->prefix . 'nsr_produtos';
}

/**
 * Retorna o nome da tabela de sessoes de bipagem.
 *
 * @return string
 */
function nsr_get_scan_sessions_table_name() {
    global $wpdb;
    return $wpdb->prefix . 'nsr_scan_sessions';
}

/**
 * Cria/atualiza tabela do plugin.
 */
function nsr_activate_plugin() {
    global $wpdb;

    $table_name = nsr_get_table_name();
    $products_table_name = nsr_get_products_table_name();
    $sessions_table_name = nsr_get_scan_sessions_table_name();
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

    $products_sql = "CREATE TABLE {$products_table_name} (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
        sku VARCHAR(255) NOT NULL,
        descricao TEXT DEFAULT '',
        created_at DATETIME NOT NULL,
        updated_at DATETIME NOT NULL,
        PRIMARY KEY (id),
        UNIQUE KEY uniq_sku (sku)
    ) {$charset_collate};";

    $sessions_sql = "CREATE TABLE {$sessions_table_name} (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
        session_token VARCHAR(80) NOT NULL,
        pedido VARCHAR(100) DEFAULT '',
        nota_fiscal VARCHAR(100) DEFAULT '',
        origem_arquivo VARCHAR(255) DEFAULT '',
        dados LONGTEXT NULL,
        created_at DATETIME NOT NULL,
        updated_at DATETIME NOT NULL,
        PRIMARY KEY (id),
        UNIQUE KEY uniq_session_token (session_token),
        KEY idx_updated_at (updated_at)
    ) {$charset_collate};";

    require_once ABSPATH . 'wp-admin/includes/upgrade.php';
    dbDelta($sql);
    dbDelta($products_sql);
    dbDelta($sessions_sql);

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
 * Faz insert/update de um registro de NS no padrao oficial do plugin.
 *
 * @param array $payload
 * @return int|false
 */
function nsr_upsert_ns_record($payload) {
    global $wpdb;

    $table_name = nsr_get_table_name();
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
        (string) $payload['ns'],
        (string) $payload['ns_normalizado'],
        (string) $payload['nota_fiscal'],
        (string) $payload['pedido'],
        (string) $payload['sku'],
        (string) $payload['descricao'],
        (string) $payload['quantidade'],
        (string) $payload['valor'],
        (string) $payload['data_venda'],
        (string) $payload['origem_arquivo'],
        (int) $payload['linha_origem']
    );

    return $wpdb->query($sql);
}

/**
 * Detecta colunas de arquivo de produtos.
 *
 * @param array $header_row
 * @return array
 */
function nsr_detect_product_columns($header_row) {
    $mapping = array(
        'sku' => null,
        'descricao' => null,
    );

    foreach ($header_row as $index => $label_raw) {
        $label = nsr_normalize_header_text($label_raw);
        if ($label === '') {
            continue;
        }

        if ($mapping['sku'] === null) {
            $sku_candidates = array('SKU', 'CODIGO SKU', 'CODIGO', 'COD PRODUTO');
            foreach ($sku_candidates as $candidate) {
                if (nsr_header_contains_phrase($label, $candidate)) {
                    $mapping['sku'] = (int) $index;
                    break;
                }
            }
        }

        if ($mapping['descricao'] === null) {
            $desc_candidates = array('DESCRICAO', 'DESCRICAO PRODUTO', 'PRODUTO');
            foreach ($desc_candidates as $candidate) {
                if (nsr_header_contains_phrase($label, $candidate)) {
                    $mapping['descricao'] = (int) $index;
                    break;
                }
            }
        }
    }

    return $mapping;
}

/**
 * Importa cadastro de produtos via CSV/XLSX.
 *
 * @param string $file_path
 * @param string $file_name
 * @return array|WP_Error
 */
function nsr_import_products_file($file_path, $file_name) {
    global $wpdb;

    $rows = nsr_read_spreadsheet_rows($file_path, $file_name);
    if (is_wp_error($rows)) {
        return $rows;
    }

    if (empty($rows)) {
        return new WP_Error('nsr_empty_products_file', 'Arquivo de produtos sem dados.');
    }

    $mapping = nsr_detect_product_columns($rows[0]);
    if ($mapping['sku'] === null) {
        return new WP_Error('nsr_products_sku_missing', 'Nao foi encontrada coluna SKU no arquivo de produtos.');
    }

    $products_table = nsr_get_products_table_name();
    $inserted = 0;
    $updated = 0;
    $ignored = 0;

    for ($i = 1; $i < count($rows); $i++) {
        $row = $rows[$i];
        $sku = strtoupper(nsr_get_cell_by_index($row, $mapping['sku']));
        $descricao = nsr_get_cell_by_index($row, $mapping['descricao']);
        if ($sku === '') {
            $ignored++;
            continue;
        }

        $sql = $wpdb->prepare(
            "INSERT INTO {$products_table} (sku, descricao, created_at, updated_at)
             VALUES (%s, %s, NOW(), NOW())
             ON DUPLICATE KEY UPDATE
                descricao = VALUES(descricao),
                updated_at = NOW()",
            $sku,
            $descricao
        );

        $affected = $wpdb->query($sql);
        if ($affected === false) {
            $ignored++;
            continue;
        }

        if ((int) $affected === 1) {
            $inserted++;
        } else {
            $updated++;
        }
    }

    return array(
        'inserted' => $inserted,
        'updated' => $updated,
        'ignored' => $ignored,
    );
}

/**
 * Carrega produtos existentes para uma lista de SKUs.
 *
 * @param array $skus
 * @return array
 */
function nsr_get_products_by_skus($skus) {
    global $wpdb;

    $skus = array_values(array_unique(array_filter(array_map('strtoupper', $skus))));
    if (empty($skus)) {
        return array();
    }

    $table = nsr_get_products_table_name();
    $placeholders = implode(',', array_fill(0, count($skus), '%s'));
    $sql = $wpdb->prepare(
        "SELECT sku, descricao FROM {$table} WHERE sku IN ({$placeholders})",
        $skus
    );

    $rows = $wpdb->get_results($sql, ARRAY_A);
    $map = array();
    foreach ($rows as $row) {
        $map[strtoupper((string) $row['sku'])] = (string) $row['descricao'];
    }

    return $map;
}

/**
 * Gera token de sessao para bipagem.
 *
 * @return string
 */
function nsr_generate_scan_session_token() {
    return wp_generate_password(24, false, false);
}

/**
 * Remove escapes de texto literal do PDF.
 *
 * @param string $text
 * @return string
 */
function nsr_pdf_unescape_literal_text($text) {
    $text = preg_replace_callback('/\\\\([nrtbf\\\\\(\)])/u', function ($matches) {
        $map = array(
            'n' => "\n",
            'r' => "\r",
            't' => "\t",
            'b' => "\b",
            'f' => "\f",
            '\\' => '\\',
            '(' => '(',
            ')' => ')',
        );

        return isset($map[$matches[1]]) ? $map[$matches[1]] : $matches[1];
    }, $text);

    $text = preg_replace_callback('/\\\\([0-7]{1,3})/u', function ($matches) {
        return chr(octdec($matches[1]));
    }, $text);

    return (string) $text;
}

/**
 * Decodifica texto hexadecimal de operacoes Tj/TJ do PDF.
 *
 * @param string $hex
 * @return string
 */
function nsr_pdf_decode_hex_text($hex) {
    $hex = preg_replace('/[^0-9A-Fa-f]/', '', (string) $hex);
    if ($hex === '') {
        return '';
    }

    if ((strlen($hex) % 2) !== 0) {
        $hex = substr($hex, 0, -1);
    }

    if ($hex === '') {
        return '';
    }

    $bin = @pack('H*', $hex);
    if ($bin === false || $bin === '') {
        return '';
    }

    // Muitos PDFs trazem texto em UTF-16BE quando usam string hexadecimal.
    if (strpos($bin, "\x00") !== false && function_exists('mb_convert_encoding')) {
        $converted = @mb_convert_encoding($bin, 'UTF-8', 'UTF-16BE');
        if (is_string($converted) && $converted !== '') {
            return $converted;
        }
    }

    return $bin;
}

/**
 * Decodifica stream em ASCIIHexDecode.
 *
 * @param string $data
 * @return string|false
 */
function nsr_pdf_decode_asciihex_stream($data) {
    $data = trim((string) $data);
    if ($data === '') {
        return false;
    }

    $data = preg_replace('/\s+/', '', $data);
    $data = preg_replace('/>.*$/', '', $data);
    $data = preg_replace('/[^0-9A-Fa-f]/', '', (string) $data);

    if ($data === '') {
        return false;
    }

    if ((strlen($data) % 2) !== 0) {
        $data .= '0';
    }

    $decoded = @pack('H*', $data);
    if ($decoded === false || $decoded === '') {
        return false;
    }

    return $decoded;
}

/**
 * Decodifica stream em ASCII85Decode.
 *
 * @param string $data
 * @return string|false
 */
function nsr_pdf_decode_ascii85_stream($data) {
    $data = (string) $data;
    if ($data === '') {
        return false;
    }

    $data = preg_replace('/\s+/', '', $data);
    $data = preg_replace('/^<~/', '', (string) $data);
    $data = preg_replace('/~>$/', '', (string) $data);

    $out = '';
    $group = '';
    $len = strlen($data);

    for ($i = 0; $i < $len; $i++) {
        $ch = $data[$i];

        if ($ch === 'z') {
            if ($group !== '') {
                return false;
            }
            $out .= "\x00\x00\x00\x00";
            continue;
        }

        $ord = ord($ch);
        if ($ord < 33 || $ord > 117) {
            continue;
        }

        $group .= $ch;

        if (strlen($group) === 5) {
            $value = 0;
            for ($g = 0; $g < 5; $g++) {
                $value = ($value * 85) + (ord($group[$g]) - 33);
            }

            $out .= chr(($value >> 24) & 0xFF);
            $out .= chr(($value >> 16) & 0xFF);
            $out .= chr(($value >> 8) & 0xFF);
            $out .= chr($value & 0xFF);
            $group = '';
        }
    }

    if ($group !== '') {
        $padding = 5 - strlen($group);
        $group_padded = $group . str_repeat('u', $padding);

        $value = 0;
        for ($g = 0; $g < 5; $g++) {
            $value = ($value * 85) + (ord($group_padded[$g]) - 33);
        }

        $chunk = '';
        $chunk .= chr(($value >> 24) & 0xFF);
        $chunk .= chr(($value >> 16) & 0xFF);
        $chunk .= chr(($value >> 8) & 0xFF);
        $chunk .= chr($value & 0xFF);

        $out .= substr($chunk, 0, 4 - $padding);
    }

    return $out;
}

/**
 * Extrai filtros declarados no dicionario do stream.
 *
 * @param string $dictionary
 * @return array
 */
function nsr_pdf_extract_stream_filters($dictionary) {
    $filters = array();
    $dictionary = (string) $dictionary;

    if (preg_match('/\/Filter\s*\[(.*?)\]/s', $dictionary, $m)) {
        if (preg_match_all('/\/([A-Za-z0-9]+)/', $m[1], $f_matches)) {
            foreach ($f_matches[1] as $name) {
                $filters[] = (string) $name;
            }
        }
        return $filters;
    }

    if (preg_match('/\/Filter\s*\/([A-Za-z0-9]+)/', $dictionary, $m)) {
        $filters[] = (string) $m[1];
    }

    return $filters;
}

/**
 * Aplica cadeia de filtros PDF ao stream.
 *
 * @param string $stream
 * @param array $filters
 * @return string
 */
function nsr_pdf_decode_stream_with_filters($stream, $filters) {
    $decoded = (string) $stream;

    if (empty($filters)) {
        return $decoded;
    }

    foreach ($filters as $filter) {
        $f = strtoupper((string) $filter);

        if ($f === 'FLATEDECODE' || $f === 'FL') {
            $try = @gzuncompress($decoded);
            if ($try === false) {
                $try = @gzinflate($decoded);
            }
            if ($try === false) {
                $try = @gzdecode($decoded);
            }
            if ($try !== false && $try !== '') {
                $decoded = $try;
            }
            continue;
        }

        if ($f === 'ASCII85DECODE' || $f === 'A85') {
            $try = nsr_pdf_decode_ascii85_stream($decoded);
            if ($try !== false && $try !== '') {
                $decoded = $try;
            }
            continue;
        }

        if ($f === 'ASCIIHEXDECODE' || $f === 'AHX') {
            $try = nsr_pdf_decode_asciihex_stream($decoded);
            if ($try !== false && $try !== '') {
                $decoded = $try;
            }
            continue;
        }
    }

    return $decoded;
}

/**
 * Extrai texto de um stream de PDF.
 *
 * @param string $stream
 * @return string
 */
function nsr_pdf_extract_text_from_stream($stream) {
    $parts = array();

    if (preg_match_all('/\((?:\\\\.|[^\\\\\)])*\)\s*Tj/s', $stream, $matches)) {
        foreach ($matches[0] as $entry) {
            if (preg_match('/\(((?:\\\\.|[^\\\\\)])*)\)\s*Tj/s', $entry, $txt_match)) {
                $parts[] = nsr_pdf_unescape_literal_text($txt_match[1]);
            }
        }
    }

    if (preg_match_all('/\[(.*?)\]\s*TJ/s', $stream, $arr_matches)) {
        foreach ($arr_matches[1] as $group) {
            if (preg_match_all('/\(((?:\\\\.|[^\\\\\)])*)\)/s', $group, $txt_matches)) {
                foreach ($txt_matches[1] as $txt) {
                    $parts[] = nsr_pdf_unescape_literal_text($txt);
                }
            }

            if (preg_match_all('/<([0-9A-Fa-f\s]+)>/s', $group, $hex_matches)) {
                foreach ($hex_matches[1] as $hex_text) {
                    $decoded_hex = nsr_pdf_decode_hex_text($hex_text);
                    if ($decoded_hex !== '') {
                        $parts[] = $decoded_hex;
                    }
                }
            }

            $parts[] = "\n";
        }
    }

    if (preg_match_all('/<([0-9A-Fa-f\s]+)>\s*Tj/s', $stream, $hex_tj_matches)) {
        foreach ($hex_tj_matches[1] as $hex_text) {
            $decoded_hex = nsr_pdf_decode_hex_text($hex_text);
            if ($decoded_hex !== '') {
                $parts[] = $decoded_hex;
            }
        }
    }

    $text = implode(' ', $parts);
    $text = preg_replace('/[^\PC\n\r\t]/u', ' ', (string) $text);
    $text = preg_replace('/\s+/', ' ', (string) $text);
    return trim((string) $text);
}

/**
 * Le texto bruto de PDF sem dependencias externas.
 *
 * @param string $file_path
 * @return string|WP_Error
 */
function nsr_read_pdf_text($file_path) {
    $binary = @file_get_contents($file_path);
    if ($binary === false || $binary === '') {
        return new WP_Error('nsr_pdf_read_error', 'Nao foi possivel ler o PDF enviado.');
    }

    if (strpos($binary, '%PDF') !== 0) {
        return new WP_Error('nsr_pdf_invalid_file', 'Arquivo enviado nao parece ser um PDF valido.');
    }

    // Tenta usar pdftotext (Poppler) como metodo primario — lida com fontes customizadas/encodings.
    $pdftotext_bin = plugin_dir_path(__FILE__) . 'bin/pdftotext.exe';
    if (file_exists($pdftotext_bin) && function_exists('shell_exec')) {
        $safe_input  = escapeshellarg((string) $file_path);
        $safe_output = escapeshellarg('-');
        // -layout preserva espacamento, -enc UTF-8 garante charset correto.
        $cmd         = '"' . $pdftotext_bin . '" -layout -enc UTF-8 ' . $safe_input . ' ' . $safe_output . ' 2>NUL';
        $result      = @shell_exec($cmd);
        if ($result !== null && trim((string) $result) !== '') {
            return trim((string) $result);
        }
    }

    $text_chunks = array();
    $offset      = 0;
    $binary_len  = strlen($binary);

    // Abordagem baseada em posicao: mais robusta que regex para PDFs com dicionarios aninhados.
    while ($offset < $binary_len) {
        $kw_pos = strpos($binary, 'stream', $offset);
        if ($kw_pos === false) {
            break;
        }

        // 'stream' deve ser seguido imediatamente por \r, \n ou \r\n (PDF spec).
        $after_kw = $kw_pos + 6;
        if ($after_kw >= $binary_len) {
            break;
        }

        $ch = $binary[$after_kw];
        if ($ch !== "\r" && $ch !== "\n") {
            $offset = $after_kw;
            continue;
        }

        $data_start = $after_kw + 1;
        if ($ch === "\r" && $data_start < $binary_len && $binary[$data_start] === "\n") {
            $data_start++;
        }

        // Procura 'endstream' a partir do inicio dos dados.
        $end_kw = strpos($binary, 'endstream', $data_start);
        if ($end_kw === false) {
            break;
        }

        // Remove \r\n antes de 'endstream' (exigido pela spec PDF).
        $data_end = $end_kw;
        if ($data_end > 0 && $binary[$data_end - 1] === "\n") {
            $data_end--;
        }
        if ($data_end > 0 && $binary[$data_end - 1] === "\r") {
            $data_end--;
        }

        // Verifica se Length esta no dicionario para extracao precisa.
        $look_start = max(0, $kw_pos - 1500);
        $pre_region = substr($binary, $look_start, $kw_pos - $look_start);
        $stream_len = -1;
        if (preg_match('/\/Length\s+(\d+)/', $pre_region, $len_m)) {
            $stream_len = (int) $len_m[1];
        }

        if ($stream_len > 0 && ($data_start + $stream_len) <= $binary_len) {
            $stream_data = substr($binary, $data_start, $stream_len);
        } else {
            $stream_data = substr($binary, $data_start, $data_end - $data_start);
        }

        $filters = nsr_pdf_extract_stream_filters($pre_region);
        $decoded = nsr_pdf_decode_stream_with_filters($stream_data, $filters);

        // Tenta todos os metodos de descompressao como fallback.
        $candidates = array($decoded);
        $c1 = @gzuncompress($stream_data);
        if ($c1 !== false && $c1 !== '') {
            $candidates[] = $c1;
        }
        $c2 = @gzinflate($stream_data);
        if ($c2 !== false && $c2 !== '') {
            $candidates[] = $c2;
        }
        $c3 = @gzdecode($stream_data);
        if ($c3 !== false && $c3 !== '') {
            $candidates[] = $c3;
        }
        if (function_exists('zlib_decode')) {
            $c4 = @zlib_decode($stream_data);
            if ($c4 !== false && $c4 !== '') {
                $candidates[] = $c4;
            }
        }

        foreach ($candidates as $candidate) {
            $chunk = nsr_pdf_extract_text_from_stream((string) $candidate);
            if ($chunk !== '') {
                $text_chunks[] = $chunk;
                break;
            }
        }

        $offset = $end_kw + 9;
    }

    $text = trim(implode("\n", $text_chunks));

    if ($text === '') {
        // Fallback: extrai caracteres imprimiveis (util apenas para PDFs sem compressao).
        $printable = preg_replace('/[^\x20-\x7E\r\n\t]/', ' ', $binary);
        $printable = preg_replace('/\s+/', ' ', (string) $printable);
        $text = trim((string) $printable);
    }

    if ($text === '') {
        $file_size_kb = round(strlen($binary) / 1024, 1);
        $stream_count = substr_count($binary, "\nstream") + substr_count($binary, "\rstream");
        return new WP_Error(
            'nsr_pdf_text_empty',
            sprintf(
                'Nao foi possivel extrair texto do PDF (arquivo de imagem/scan). Tamanho: %s KB | Streams: %d. Use a entrada manual de itens abaixo.',
                $file_size_kb,
                $stream_count
            )
        );
    }

    return $text;
}

/**
 * Identifica se token parece SKU valido.
 *
 * @param string $token
 * @return bool
 */
function nsr_is_probable_sku($token) {
    $token = strtoupper(trim((string) $token));
    if ($token === '' || strlen($token) < 3 || strlen($token) > 60) {
        return false;
    }

    if (!preg_match('/^[A-Z0-9._\/-]+$/', $token)) {
        return false;
    }

    // Evita confundir com codigos puramente numericos de pedido/NF.
    if (preg_match('/^\d+$/', $token)) {
        return false;
    }

    return true;
}

/**
 * Converte valor textual de quantidade para inteiro.
 *
 * Exemplos aceitos: 988,00 | 1.375,00 | 150
 *
 * @param string $value
 * @return int
 */
function nsr_parse_quantity_value($value) {
    $value = trim((string) $value);
    if ($value === '') {
        return 0;
    }

    $value = preg_replace('/[^0-9,\.]/', '', $value);
    if ($value === '') {
        return 0;
    }

    // Formato BR: 1.375,00
    if (strpos($value, ',') !== false) {
        $value = str_replace('.', '', $value);
        $value = str_replace(',', '.', $value);
    }

    $qty = (int) round((float) $value);
    return max(0, $qty);
}

/**
 * Extrai itens no formato do pedido KDT (SKU/GTIN + Qtd + Un).
 *
 * @param string $text
 * @return array
 */
function nsr_extract_kdt_items_from_pdf_text($text) {
    $items = array();

    $normalized = strtoupper(remove_accents((string) $text));
    $normalized = preg_replace('/\s+/', ' ', $normalized);

    // Captura sequencias como: PRD00040 7898722600065 988,00 UN
    if (preg_match_all('/\b([A-Z]{2,}[A-Z0-9._\/-]{2,})\b\s+(?:\d{8,14}\s+)?(\d{1,3}(?:\.\d{3})*(?:,\d{2})?|\d{1,6})\s+UN\b/u', $normalized, $matches, PREG_SET_ORDER)) {
        foreach ($matches as $m) {
            $sku = strtoupper((string) $m[1]);
            $qty = nsr_parse_quantity_value($m[2]);

            if (!nsr_is_probable_sku($sku) || $qty <= 0) {
                continue;
            }

            if (!isset($items[$sku])) {
                $items[$sku] = array(
                    'sku' => $sku,
                    'descricao' => '',
                    'quantidade' => 0,
                    'scanned' => array(),
                );
            }

            $items[$sku]['quantidade'] += $qty;
        }
    }

    return $items;
}

/**
 * Extrai pedido, nota fiscal e itens (SKU + quantidade) do texto de PDF.
 *
 * @param string $text
 * @return array
 */
function nsr_extract_order_from_pdf_text($text) {
    $raw_lines = preg_split('/\r\n|\r|\n/', (string) $text);
    $lines = array();

    foreach ($raw_lines as $line) {
        $clean = trim(preg_replace('/\s+/', ' ', (string) $line));
        if ($clean !== '') {
            $lines[] = $clean;
        }
    }

    $pedido = '';
    $nota_fiscal = '';

    foreach ($lines as $line) {
        $norm = strtoupper(remove_accents($line));

        if ($pedido === '' && preg_match('/\bPEDIDO\s+DE\s+VENDA\s+N\s*[\x{00BA}\x{00B0}O]?\s*(\d{1,10})\b/u', $norm, $m)) {
            $pedido = $m[1];
        }

        if ($pedido === '' && preg_match('/\bPEDIDO\b[^0-9]{0,12}(\d{3,})/', $norm, $m)) {
            $pedido = $m[1];
        }

        if ($nota_fiscal === '' && preg_match('/\b(?:NOTA\s+FISCAL|NF(?:E)?)\b[^0-9]{0,12}(\d{3,})/', $norm, $m)) {
            $nota_fiscal = $m[1];
        }
    }

    $items = nsr_extract_kdt_items_from_pdf_text($text);

    // Fallback generico para outros formatos, caso o padrao KDT nao encontre itens.
    if (!empty($items)) {
        return array(
            'pedido' => $pedido,
            'nota_fiscal' => $nota_fiscal,
            'itens' => $items,
        );
    }

    $items = array();

    foreach ($lines as $line) {
        $norm = strtoupper(remove_accents($line));

        $sku = '';
        $qty = 0;

        if (preg_match('/\b(?:SKU|CODIGO(?:\s+DO\s+PRODUTO)?|COD\.)\b\s*[:\-]?\s*([A-Z0-9._\/-]{3,})/', $norm, $m)) {
            $sku = strtoupper($m[1]);
        }

        if (preg_match('/\b(?:QTD|QUANTIDADE)\b\s*[:x\-]?\s*([0-9\.,]{1,20})\b/', $norm, $m)) {
            $qty = nsr_parse_quantity_value($m[1]);
        }

        if ($sku === '') {
            $tokens = preg_split('/\s+/', $norm);
            foreach ($tokens as $token) {
                $token = trim((string) $token, " \t\n\r\0\x0B,;:()[]{}");
                if (nsr_is_probable_sku($token)) {
                    $sku = $token;
                    break;
                }
            }
        }

        if ($qty <= 0 && preg_match('/\b([0-9\.,]{1,20})\b\s*$/', $norm, $m)) {
            $qty = nsr_parse_quantity_value($m[1]);
        }

        if ($sku === '' || $qty <= 0) {
            continue;
        }

        if (!isset($items[$sku])) {
            $items[$sku] = array(
                'sku' => $sku,
                'descricao' => '',
                'quantidade' => 0,
                'scanned' => array(),
            );
        }

        $items[$sku]['quantidade'] += $qty;
    }

    return array(
        'pedido' => $pedido,
        'nota_fiscal' => $nota_fiscal,
        'itens' => $items,
    );
}

/**
 * Salva sessao de bipagem em banco.
 *
 * @param string $token
 * @param array $data
 * @return bool
 */
function nsr_save_scan_session($token, $data) {
    global $wpdb;

    $table = nsr_get_scan_sessions_table_name();
    $json = wp_json_encode($data);
    if ($json === false) {
        return false;
    }

    $sql = $wpdb->prepare(
        "INSERT INTO {$table} (session_token, pedido, nota_fiscal, origem_arquivo, dados, created_at, updated_at)
         VALUES (%s, %s, %s, %s, %s, NOW(), NOW())
         ON DUPLICATE KEY UPDATE
            pedido = VALUES(pedido),
            nota_fiscal = VALUES(nota_fiscal),
            origem_arquivo = VALUES(origem_arquivo),
            dados = VALUES(dados),
            updated_at = NOW()",
        $token,
        isset($data['pedido']) ? (string) $data['pedido'] : '',
        isset($data['nota_fiscal']) ? (string) $data['nota_fiscal'] : '',
        isset($data['origem_arquivo']) ? (string) $data['origem_arquivo'] : '',
        $json
    );

    return $wpdb->query($sql) !== false;
}

/**
 * Carrega sessao de bipagem.
 *
 * @param string $token
 * @return array|null
 */
function nsr_get_scan_session($token) {
    global $wpdb;

    if ($token === '') {
        return null;
    }

    $table = nsr_get_scan_sessions_table_name();
    $row = $wpdb->get_row(
        $wpdb->prepare(
            "SELECT session_token, pedido, nota_fiscal, origem_arquivo, dados
             FROM {$table}
             WHERE session_token = %s
             LIMIT 1",
            $token
        ),
        ARRAY_A
    );

    if (empty($row)) {
        return null;
    }

    $data = json_decode((string) $row['dados'], true);
    if (!is_array($data)) {
        return null;
    }

    $data['session_token'] = (string) $row['session_token'];
    $data['pedido'] = isset($data['pedido']) ? (string) $data['pedido'] : (string) $row['pedido'];
    $data['nota_fiscal'] = isset($data['nota_fiscal']) ? (string) $data['nota_fiscal'] : (string) $row['nota_fiscal'];
    $data['origem_arquivo'] = isset($data['origem_arquivo']) ? (string) $data['origem_arquivo'] : (string) $row['origem_arquivo'];

    return $data;
}

/**
 * Remove sessao de bipagem.
 *
 * @param string $token
 * @return void
 */
function nsr_delete_scan_session($token) {
    global $wpdb;

    if ($token === '') {
        return;
    }

    $wpdb->delete(
        nsr_get_scan_sessions_table_name(),
        array('session_token' => $token),
        array('%s')
    );
}

/**
 * Limpa sessoes antigas de bipagem.
 *
 * @return void
 */
function nsr_cleanup_old_scan_sessions() {
    global $wpdb;

    $table = nsr_get_scan_sessions_table_name();
    $wpdb->query(
        "DELETE FROM {$table}
         WHERE updated_at < (NOW() - INTERVAL 2 DAY)"
    );
}

/**
 * Recalcula flags de validacao da sessao.
 *
 * @param array $session
 * @return array
 */
function nsr_recompute_scan_session_flags($session) {
    $skus = array_keys(isset($session['itens']) ? $session['itens'] : array());
    $products_map = nsr_get_products_by_skus($skus);

    $missing_skus = array();
    foreach ($skus as $sku) {
        if (!isset($products_map[$sku])) {
            $missing_skus[] = $sku;
        } elseif ((string) $session['itens'][$sku]['descricao'] === '') {
            $session['itens'][$sku]['descricao'] = (string) $products_map[$sku];
        }
    }

    $session['missing_skus'] = $missing_skus;
    return $session;
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

            $affected = nsr_upsert_ns_record(array(
                'ns' => $ns,
                'ns_normalizado' => $ns_normalizado,
                'nota_fiscal' => $nf,
                'pedido' => $pedido,
                'sku' => $sku,
                'descricao' => $descricao,
                'quantidade' => $quantidade,
                'valor' => $valor,
                'data_venda' => $data_venda,
                'origem_arquivo' => $file_name,
                'linha_origem' => ($i + 1),
            ));

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
 * Junta mensagens de feedback em um unico array.
 *
 * @param array $base
 * @param array $extra
 * @return array
 */
function nsr_merge_messages($base, $extra) {
    if (isset($extra['success']) && is_array($extra['success'])) {
        $base['success'] = array_merge($base['success'], $extra['success']);
    }
    if (isset($extra['error']) && is_array($extra['error'])) {
        $base['error'] = array_merge($base['error'], $extra['error']);
    }
    return $base;
}

/**
 * Processa importacao de base de produtos.
 *
 * @return array
 */
function nsr_handle_products_import_submission() {
    $messages = array(
        'success' => array(),
        'error' => array(),
    );

    if (!isset($_POST['nsr_products_import_submit'])) {
        return $messages;
    }

    if (!current_user_can('manage_options')) {
        $messages['error'][] = 'Permissao insuficiente para importar produtos.';
        return $messages;
    }

    check_admin_referer('nsr_products_import', 'nsr_products_nonce');

    if (empty($_FILES['nsr_products_file']) || !is_array($_FILES['nsr_products_file'])) {
        $messages['error'][] = 'Nenhum arquivo de produtos foi enviado.';
        return $messages;
    }

    $file = $_FILES['nsr_products_file'];
    $safe_name = sanitize_file_name((string) wp_unslash(isset($file['name']) ? $file['name'] : ''));
    $tmp_name = isset($file['tmp_name']) ? (string) $file['tmp_name'] : '';
    $error_code = isset($file['error']) ? (int) $file['error'] : UPLOAD_ERR_NO_FILE;

    if ($safe_name === '' || $tmp_name === '') {
        $messages['error'][] = 'Arquivo de produtos invalido.';
        return $messages;
    }

    if ($error_code !== UPLOAD_ERR_OK) {
        $messages['error'][] = sprintf('Falha no upload do arquivo de produtos (%d).', $error_code);
        return $messages;
    }

    $extension = strtolower(pathinfo($safe_name, PATHINFO_EXTENSION));
    if (!in_array($extension, array('xlsx', 'csv'), true)) {
        $messages['error'][] = 'Formato nao suportado para produtos. Use .xlsx ou .csv.';
        return $messages;
    }

    $result = nsr_import_products_file($tmp_name, $safe_name);
    if (is_wp_error($result)) {
        $messages['error'][] = $result->get_error_message();
        return $messages;
    }

    $messages['success'][] = sprintf(
        'Produtos: %d inserido(s), %d atualizado(s), %d ignorado(s).',
        (int) $result['inserted'],
        (int) $result['updated'],
        (int) $result['ignored']
    );

    return $messages;
}

/**
 * Processa upload de PDF e fluxo de bipagem de NS.
 *
 * @return array
 */
function nsr_handle_pdf_scan_workflow_submission() {
    $messages = array(
        'success' => array(),
        'error' => array(),
    );

    nsr_cleanup_old_scan_sessions();

    $active_token = '';
    if (isset($_POST['nsr_scan_session_token'])) {
        $active_token = sanitize_text_field(wp_unslash($_POST['nsr_scan_session_token']));
    } elseif (isset($_GET['nsr_scan_session'])) {
        $active_token = sanitize_text_field(wp_unslash($_GET['nsr_scan_session']));
    }

    if (isset($_POST['nsr_pdf_upload_submit'])) {
        if (!current_user_can('manage_options')) {
            $messages['error'][] = 'Permissao insuficiente para carregar PDF.';
        } else {
            check_admin_referer('nsr_pdf_upload', 'nsr_pdf_nonce');

            if (empty($_FILES['nsr_pdf_file']) || !is_array($_FILES['nsr_pdf_file'])) {
                $messages['error'][] = 'Nenhum PDF foi enviado.';
            } else {
                $file = $_FILES['nsr_pdf_file'];
                $safe_name = sanitize_file_name((string) wp_unslash(isset($file['name']) ? $file['name'] : ''));
                $tmp_name = isset($file['tmp_name']) ? (string) $file['tmp_name'] : '';
                $error_code = isset($file['error']) ? (int) $file['error'] : UPLOAD_ERR_NO_FILE;

                if ($safe_name === '' || $tmp_name === '') {
                    $messages['error'][] = 'Arquivo PDF invalido.';
                } elseif ($error_code !== UPLOAD_ERR_OK) {
                    $messages['error'][] = sprintf('Falha no upload do PDF (%d).', $error_code);
                } elseif (strtolower(pathinfo($safe_name, PATHINFO_EXTENSION)) !== 'pdf') {
                    $messages['error'][] = 'Formato nao suportado. Envie um arquivo .pdf.';
                } else {
                    $text = nsr_read_pdf_text($tmp_name);
                    if (is_wp_error($text)) {
                        $messages['error'][] = $text->get_error_message();
                    } else {
                        $parsed = nsr_extract_order_from_pdf_text($text);
                        if (empty($parsed['itens'])) {
                            $sample = esc_html(substr($text, 0, 400));
                            $messages['error'][] = 'Nao foi possivel detectar SKU e quantidade no PDF. Texto extraido (amostra): ' . $sample;
                        } else {
                            $token = nsr_generate_scan_session_token();
                            $session = array(
                                'session_token' => $token,
                                'pedido' => (string) $parsed['pedido'],
                                'nota_fiscal' => (string) $parsed['nota_fiscal'],
                                'origem_arquivo' => $safe_name,
                                'itens' => $parsed['itens'],
                                'missing_skus' => array(),
                            );

                            $session = nsr_recompute_scan_session_flags($session);

                            if (!nsr_save_scan_session($token, $session)) {
                                $messages['error'][] = 'Nao foi possivel salvar a sessao de bipagem.';
                            } else {
                                $active_token = $token;
                                $messages['success'][] = sprintf('PDF %s lido com sucesso. Inicie a bipagem dos NS.', $safe_name);
                            }
                        }
                    }
                }
            }
        }
    }

    // ----- Entrada manual de itens (fallback quando PDF e imagem) -----
    if (isset($_POST['nsr_manual_items_submit'])) {
        if (!current_user_can('manage_options')) {
            $messages['error'][] = 'Permissao insuficiente.';
        } else {
            check_admin_referer('nsr_manual_items', 'nsr_manual_nonce');

            $raw_lines = isset($_POST['nsr_manual_items_text']) ? sanitize_textarea_field(wp_unslash($_POST['nsr_manual_items_text'])) : '';
            $pedido_manual = sanitize_text_field(wp_unslash(isset($_POST['nsr_manual_pedido']) ? $_POST['nsr_manual_pedido'] : ''));
            $nf_manual    = sanitize_text_field(wp_unslash(isset($_POST['nsr_manual_nf']) ? $_POST['nsr_manual_nf'] : ''));

            $itens_manual = array();
            foreach (preg_split('/[\r\n]+/', $raw_lines) as $line) {
                $line = trim($line);
                if ($line === '' || $line[0] === '#') {
                    continue;
                }
                // Suporta separadores: ; | , (tab)
                $parts = preg_split('/[;|,\t]/', $line, 2);
                $sku_raw = isset($parts[0]) ? strtoupper(trim($parts[0])) : '';
                $qty_raw = isset($parts[1]) ? trim($parts[1]) : '1';
                if ($sku_raw === '' || !preg_match('/^[A-Z0-9._\/ -]+$/', $sku_raw)) {
                    continue;
                }
                $qty = nsr_parse_quantity_value($qty_raw);
                if ($qty <= 0) {
                    $qty = 1;
                }
                $itens_manual[$sku_raw] = array(
                    'quantidade' => $qty,
                    'ns_lidos'   => array(),
                );
            }

            if (empty($itens_manual)) {
                $messages['error'][] = 'Nenhum item valido encontrado. Use o formato: SKU;QUANTIDADE (uma linha por item).';
            } else {
                $token  = nsr_generate_scan_session_token();
                $session_data = array(
                    'session_token'  => $token,
                    'pedido'         => $pedido_manual,
                    'nota_fiscal'    => $nf_manual,
                    'origem_arquivo' => 'entrada_manual',
                    'itens'          => $itens_manual,
                    'missing_skus'   => array(),
                );
                $session_data = nsr_recompute_scan_session_flags($session_data);
                if (!nsr_save_scan_session($token, $session_data)) {
                    $messages['error'][] = 'Nao foi possivel salvar a sessao de bipagem.';
                } else {
                    $active_token = $token;
                    $messages['success'][] = sprintf(
                        'Sessao criada manualmente com %d SKU(s). Inicie a bipagem dos NS.',
                        count($itens_manual)
                    );
                }
            }
        }
    }
    // ----- fim entrada manual -----

    $session = nsr_get_scan_session($active_token);

    if (!empty($session)) {
        if (!current_user_can('manage_options')) {
            $messages['error'][] = 'Permissao insuficiente para continuar a bipagem.';
            return array(
                'messages' => $messages,
                'session' => $session,
                'active_token' => $active_token,
            );
        }

        if (isset($_POST['nsr_pedido'])) {
            $session['pedido'] = sanitize_text_field(wp_unslash($_POST['nsr_pedido']));
        }
        if (isset($_POST['nsr_nota_fiscal'])) {
            $session['nota_fiscal'] = sanitize_text_field(wp_unslash($_POST['nsr_nota_fiscal']));
        }

        if (isset($_POST['nsr_scan_add_submit'])) {
            check_admin_referer('nsr_scan_action', 'nsr_scan_nonce');

            $sku = strtoupper(sanitize_text_field(wp_unslash(isset($_POST['nsr_scan_sku']) ? $_POST['nsr_scan_sku'] : '')));
            $ns = sanitize_text_field(wp_unslash(isset($_POST['nsr_scan_ns']) ? $_POST['nsr_scan_ns'] : ''));
            $ns_normalizado = nsr_normalize_lookup_value($ns);

            if ($sku === '' || $ns_normalizado === '') {
                $messages['error'][] = 'Informe SKU e NS para bipagem.';
            } elseif (!isset($session['itens'][$sku])) {
                $messages['error'][] = sprintf('SKU %s nao faz parte do pedido lido no PDF.', $sku);
            } else {
                $expected = (int) $session['itens'][$sku]['quantidade'];
                $already = isset($session['itens'][$sku]['scanned']) ? count($session['itens'][$sku]['scanned']) : 0;

                $all_scanned = array();
                foreach ($session['itens'] as $item) {
                    if (!empty($item['scanned']) && is_array($item['scanned'])) {
                        $all_scanned = array_merge($all_scanned, $item['scanned']);
                    }
                }

                $all_normalized = array_map('nsr_normalize_lookup_value', $all_scanned);
                if (in_array($ns_normalizado, $all_normalized, true)) {
                    $messages['error'][] = 'NS ja foi bipado nesta sessao.';
                } elseif ($already >= $expected) {
                    $messages['error'][] = sprintf('SKU %s ja atingiu a quantidade esperada (%d).', $sku, $expected);
                } else {
                    if (!isset($session['itens'][$sku]['scanned']) || !is_array($session['itens'][$sku]['scanned'])) {
                        $session['itens'][$sku]['scanned'] = array();
                    }
                    $session['itens'][$sku]['scanned'][] = $ns;
                    $messages['success'][] = sprintf('NS %s vinculado ao SKU %s.', $ns, $sku);
                }
            }
        }

        if (isset($_POST['nsr_scan_remove_submit'])) {
            check_admin_referer('nsr_scan_action', 'nsr_scan_nonce');

            $sku = strtoupper(sanitize_text_field(wp_unslash(isset($_POST['nsr_scan_remove_sku']) ? $_POST['nsr_scan_remove_sku'] : '')));
            $ns = sanitize_text_field(wp_unslash(isset($_POST['nsr_scan_remove_ns']) ? $_POST['nsr_scan_remove_ns'] : ''));

            if ($sku !== '' && $ns !== '' && isset($session['itens'][$sku]['scanned']) && is_array($session['itens'][$sku]['scanned'])) {
                $idx = array_search($ns, $session['itens'][$sku]['scanned'], true);
                if ($idx !== false) {
                    unset($session['itens'][$sku]['scanned'][$idx]);
                    $session['itens'][$sku]['scanned'] = array_values($session['itens'][$sku]['scanned']);
                    $messages['success'][] = sprintf('NS %s removido do SKU %s.', $ns, $sku);
                }
            }
        }

        if (isset($_POST['nsr_scan_abort_submit'])) {
            check_admin_referer('nsr_scan_action', 'nsr_scan_nonce');
            nsr_delete_scan_session($active_token);
            $session = null;
            $active_token = '';
            $messages['success'][] = 'Sessao de bipagem encerrada.';
        } elseif (isset($_POST['nsr_scan_finalize_submit'])) {
            check_admin_referer('nsr_scan_action', 'nsr_scan_nonce');

            $session = nsr_recompute_scan_session_flags($session);

            if (empty($session['pedido']) && empty($session['nota_fiscal'])) {
                $messages['error'][] = 'Informe ao menos Pedido ou Nota Fiscal para finalizar.';
            }

            if (!empty($session['missing_skus'])) {
                $messages['error'][] = 'Existem SKUs do pedido sem cadastro de produto.';
            }

            $has_qty_error = false;
            foreach ($session['itens'] as $sku => $item) {
                $expected = (int) $item['quantidade'];
                $scanned = isset($item['scanned']) && is_array($item['scanned']) ? count($item['scanned']) : 0;
                if ($expected !== $scanned) {
                    $has_qty_error = true;
                    $messages['error'][] = sprintf('SKU %s com quantidade divergente: esperado %d, bipado %d.', $sku, $expected, $scanned);
                }
            }

            if (empty($messages['error']) && !$has_qty_error) {
                global $wpdb;
                $wpdb->query('START TRANSACTION');
                $saved = 0;
                $failed = false;

                foreach ($session['itens'] as $sku => $item) {
                    $descricao = isset($item['descricao']) ? (string) $item['descricao'] : '';
                    $expected = (int) $item['quantidade'];
                    $scanned_list = isset($item['scanned']) && is_array($item['scanned']) ? $item['scanned'] : array();

                    foreach ($scanned_list as $ns) {
                        $ns_normalizado = nsr_normalize_lookup_value($ns);
                        if ($ns_normalizado === '') {
                            $failed = true;
                            break;
                        }

                        $affected = nsr_upsert_ns_record(array(
                            'ns' => $ns,
                            'ns_normalizado' => $ns_normalizado,
                            'nota_fiscal' => (string) $session['nota_fiscal'],
                            'pedido' => (string) $session['pedido'],
                            'sku' => $sku,
                            'descricao' => $descricao,
                            'quantidade' => (string) $expected,
                            'valor' => '',
                            'data_venda' => '',
                            'origem_arquivo' => (string) $session['origem_arquivo'],
                            'linha_origem' => 0,
                        ));

                        if ($affected === false) {
                            $failed = true;
                            break;
                        }

                        $saved++;
                    }

                    if ($failed) {
                        break;
                    }
                }

                if ($failed) {
                    $wpdb->query('ROLLBACK');
                    $messages['error'][] = 'Falha ao salvar NS no banco. Nenhum dado foi confirmado.';
                } else {
                    $wpdb->query('COMMIT');
                    nsr_delete_scan_session($active_token);
                    $session = null;
                    $active_token = '';
                    $messages['success'][] = sprintf('Bipagem finalizada com sucesso. %d NS salvo(s).', $saved);
                }
            }
        }

        if (!empty($session)) {
            $session = nsr_recompute_scan_session_flags($session);
            nsr_save_scan_session($active_token, $session);
        }
    }

    return array(
        'messages' => $messages,
        'session' => $session,
        'active_token' => $active_token,
    );
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

    $messages = array(
        'success' => array(),
        'error' => array(),
    );
    $messages = nsr_merge_messages($messages, nsr_handle_import_submission());
    $messages = nsr_merge_messages($messages, nsr_handle_products_import_submission());
    $pdf_workflow = nsr_handle_pdf_scan_workflow_submission();
    $messages = nsr_merge_messages($messages, $pdf_workflow['messages']);
    $scan_session = isset($pdf_workflow['session']) && is_array($pdf_workflow['session']) ? $pdf_workflow['session'] : null;

    foreach ($messages['success'] as $message) {
        add_settings_error('nsr_messages', 'nsr_success_' . wp_rand(), $message, 'updated');
    }
    foreach ($messages['error'] as $message) {
        add_settings_error('nsr_messages', 'nsr_error_' . wp_rand(), $message, 'error');
    }

    $table_name = nsr_get_table_name();
    $products_table = nsr_get_products_table_name();
    $total_records = (int) $wpdb->get_var("SELECT COUNT(1) FROM {$table_name}");
    $total_ns_unicos = (int) $wpdb->get_var("SELECT COUNT(DISTINCT ns_normalizado) FROM {$table_name}");
    $total_products = (int) $wpdb->get_var("SELECT COUNT(1) FROM {$products_table}");

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
            | <strong>Produtos cadastrados:</strong> <?php echo esc_html((string) $total_products); ?>
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

        <h2>2) Importar base de produtos (SKU x Descricao)</h2>
        <p>Envie arquivo <code>.xlsx</code> ou <code>.csv</code> com colunas de SKU e descricao para validar o pedido do PDF.</p>
        <form method="post" enctype="multipart/form-data" style="margin-bottom:24px;">
            <?php wp_nonce_field('nsr_products_import', 'nsr_products_nonce'); ?>
            <input type="file" name="nsr_products_file" accept=".xlsx,.csv" required />
            <p>
                <button type="submit" name="nsr_products_import_submit" class="button">Importar Produtos</button>
            </p>
        </form>

        <h2>3) Leitura de Pedido de Venda (PDF) e Bipagem de NS</h2>
        <p>Envie o PDF do pedido para extrair SKU e quantidade. Depois, realize a bipagem dos NS por SKU.</p>
        <form method="post" enctype="multipart/form-data" style="margin-bottom:16px;">
            <?php wp_nonce_field('nsr_pdf_upload', 'nsr_pdf_nonce'); ?>
            <input type="file" name="nsr_pdf_file" accept=".pdf" required />
            <button type="submit" name="nsr_pdf_upload_submit" class="button">Ler PDF</button>
        </form>

        <details style="margin-bottom:16px;border:1px solid #dcdcde;border-radius:6px;padding:12px;" open>
            <summary style="cursor:pointer;font-weight:600;">Inserir itens manualmente (use quando o PDF e imagem/scan)</summary>
            <p style="margin-top:8px;color:#555;">Digite um item por linha no formato: <code>SKU;QUANTIDADE</code><br>
            Separadores aceitos: <code>;</code> <code>|</code> <code>,</code> ou TAB. Linhas com <code>#</code> sao ignoradas.</p>
            <form method="post" style="margin-top:8px;">
                <?php wp_nonce_field('nsr_manual_items', 'nsr_manual_nonce'); ?>
                <table style="margin-bottom:10px;">
                    <tr>
                        <td style="padding-right:12px;"><label><strong>Pedido:</strong><br>
                            <input type="text" name="nsr_manual_pedido" style="width:160px;" placeholder="Ex: 12345" /></label></td>
                        <td><label><strong>Nota Fiscal:</strong><br>
                            <input type="text" name="nsr_manual_nf" style="width:160px;" placeholder="Ex: 67890" /></label></td>
                    </tr>
                </table>
                <textarea name="nsr_manual_items_text" rows="8" cols="60" placeholder="SKU-001;10&#10;SKU-002;5&#10;SKU-003;1" style="font-family:monospace;"></textarea>
                <p><button type="submit" name="nsr_manual_items_submit" class="button button-primary">Criar Sessao de Bipagem</button></p>
            </form>
        </details>

        <?php if (!empty($scan_session)) : ?>
            <?php
            $scan_token = isset($scan_session['session_token']) ? (string) $scan_session['session_token'] : '';
            $missing_skus = isset($scan_session['missing_skus']) && is_array($scan_session['missing_skus']) ? $scan_session['missing_skus'] : array();
            ?>
            <div style="padding:12px;background:#fff;border:1px solid #dcdcde;border-radius:6px;margin-bottom:16px;">
                <h3 style="margin:0 0 8px 0;">Previa da extracao do PDF</h3>
                <p style="margin-top:0;">
                    <strong>Arquivo:</strong> <?php echo esc_html(isset($scan_session['origem_arquivo']) ? $scan_session['origem_arquivo'] : ''); ?>
                    | <strong>Pedido:</strong> <?php echo esc_html(isset($scan_session['pedido']) ? $scan_session['pedido'] : ''); ?>
                    | <strong>NF:</strong> <?php echo esc_html(isset($scan_session['nota_fiscal']) ? $scan_session['nota_fiscal'] : ''); ?>
                </p>

                <table class="widefat" style="max-width:100%;margin-bottom:14px;">
                    <thead>
                        <tr>
                            <th>SKU extraido</th>
                            <th>Qtd extraida</th>
                            <th>Descricao cadastro</th>
                            <th>Status cadastro</th>
                        </tr>
                    </thead>
                    <tbody>
                        <?php foreach ($scan_session['itens'] as $sku => $item) : ?>
                            <tr>
                                <td><?php echo esc_html($sku); ?></td>
                                <td><?php echo esc_html((string) ((int) $item['quantidade'])); ?></td>
                                <td><?php echo esc_html(isset($item['descricao']) ? $item['descricao'] : ''); ?></td>
                                <td>
                                    <?php if (in_array($sku, $missing_skus, true)) : ?>
                                        <span style="color:#b32d2e;font-weight:600;">SKU nao cadastrado</span>
                                    <?php else : ?>
                                        <span style="color:#0a7d28;font-weight:600;">OK</span>
                                    <?php endif; ?>
                                </td>
                            </tr>
                        <?php endforeach; ?>
                    </tbody>
                </table>

                <h3 style="margin:0 0 8px 0;">Bipagem de NS</h3>
                <form method="post" style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:10px;">
                    <?php wp_nonce_field('nsr_scan_action', 'nsr_scan_nonce'); ?>
                    <input type="hidden" name="nsr_scan_session_token" value="<?php echo esc_attr($scan_token); ?>" />
                    <label style="display:flex;align-items:center;gap:6px;">
                        Pedido
                        <input type="text" name="nsr_pedido" value="<?php echo esc_attr(isset($scan_session['pedido']) ? $scan_session['pedido'] : ''); ?>" />
                    </label>
                    <label style="display:flex;align-items:center;gap:6px;">
                        Nota Fiscal
                        <input type="text" name="nsr_nota_fiscal" value="<?php echo esc_attr(isset($scan_session['nota_fiscal']) ? $scan_session['nota_fiscal'] : ''); ?>" />
                    </label>
                    <label style="display:flex;align-items:center;gap:6px;">
                        SKU
                        <input type="text" name="nsr_scan_sku" placeholder="SKU" required />
                    </label>
                    <label style="display:flex;align-items:center;gap:6px;">
                        NS
                        <input type="text" name="nsr_scan_ns" placeholder="Numero de Serie" required />
                    </label>
                    <button type="submit" name="nsr_scan_add_submit" class="button button-primary">Bipar NS</button>
                </form>

                <?php if (!empty($missing_skus)) : ?>
                    <div style="padding:10px;background:#fff3cd;border:1px solid #ffc107;border-radius:4px;color:#856404;margin-bottom:10px;">
                        SKU(s) sem cadastro de produto: <strong><?php echo esc_html(implode(', ', $missing_skus)); ?></strong>
                    </div>
                <?php endif; ?>

                <table class="widefat striped" style="max-width:100%;margin-bottom:10px;">
                    <thead>
                        <tr>
                            <th>SKU</th>
                            <th>Descricao</th>
                            <th>Qtd Pedido</th>
                            <th>Qtd Bipado</th>
                            <th>Status</th>
                            <th>NS bipados</th>
                        </tr>
                    </thead>
                    <tbody>
                        <?php
                        $can_finalize = true;
                        foreach ($scan_session['itens'] as $sku => $item) :
                            $expected = (int) $item['quantidade'];
                            $scanned_list = isset($item['scanned']) && is_array($item['scanned']) ? $item['scanned'] : array();
                            $scanned_count = count($scanned_list);
                            $is_ok = ($expected === $scanned_count);
                            if (!$is_ok) {
                                $can_finalize = false;
                            }
                            if (in_array($sku, $missing_skus, true)) {
                                $can_finalize = false;
                            }
                            ?>
                            <tr>
                                <td><?php echo esc_html($sku); ?></td>
                                <td><?php echo esc_html(isset($item['descricao']) ? $item['descricao'] : ''); ?></td>
                                <td><?php echo esc_html((string) $expected); ?></td>
                                <td><?php echo esc_html((string) $scanned_count); ?></td>
                                <td>
                                    <?php if ($is_ok) : ?>
                                        <span style="color:#0a7d28;font-weight:600;">OK</span>
                                    <?php else : ?>
                                        <span style="color:#b32d2e;font-weight:600;">Divergente</span>
                                    <?php endif; ?>
                                </td>
                                <td>
                                    <?php if (empty($scanned_list)) : ?>
                                        -
                                    <?php else : ?>
                                        <?php foreach ($scanned_list as $scanned_ns) : ?>
                                            <form method="post" style="display:inline-block;margin:0 6px 6px 0;">
                                                <?php wp_nonce_field('nsr_scan_action', 'nsr_scan_nonce'); ?>
                                                <input type="hidden" name="nsr_scan_session_token" value="<?php echo esc_attr($scan_token); ?>" />
                                                <input type="hidden" name="nsr_pedido" value="<?php echo esc_attr(isset($scan_session['pedido']) ? $scan_session['pedido'] : ''); ?>" />
                                                <input type="hidden" name="nsr_nota_fiscal" value="<?php echo esc_attr(isset($scan_session['nota_fiscal']) ? $scan_session['nota_fiscal'] : ''); ?>" />
                                                <input type="hidden" name="nsr_scan_remove_sku" value="<?php echo esc_attr($sku); ?>" />
                                                <input type="hidden" name="nsr_scan_remove_ns" value="<?php echo esc_attr($scanned_ns); ?>" />
                                                <button type="submit" name="nsr_scan_remove_submit" class="button button-small" title="Remover NS" style="padding:0 8px;line-height:1.6;">
                                                    <?php echo esc_html($scanned_ns); ?> x
                                                </button>
                                            </form>
                                        <?php endforeach; ?>
                                    <?php endif; ?>
                                </td>
                            </tr>
                        <?php endforeach; ?>
                    </tbody>
                </table>

                <form method="post" style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
                    <?php wp_nonce_field('nsr_scan_action', 'nsr_scan_nonce'); ?>
                    <input type="hidden" name="nsr_scan_session_token" value="<?php echo esc_attr($scan_token); ?>" />
                    <input type="hidden" name="nsr_pedido" value="<?php echo esc_attr(isset($scan_session['pedido']) ? $scan_session['pedido'] : ''); ?>" />
                    <input type="hidden" name="nsr_nota_fiscal" value="<?php echo esc_attr(isset($scan_session['nota_fiscal']) ? $scan_session['nota_fiscal'] : ''); ?>" />
                    <button type="submit" name="nsr_scan_finalize_submit" class="button button-primary" <?php disabled(!$can_finalize || empty($scan_session['pedido']) && empty($scan_session['nota_fiscal'])); ?>>
                        Finalizar e salvar NS
                    </button>
                    <button type="submit" name="nsr_scan_abort_submit" class="button">Cancelar sessao</button>
                    <?php if (!$can_finalize) : ?>
                        <span style="color:#b32d2e;">Nao e possivel finalizar: verifique SKUs e quantidades bipadas.</span>
                    <?php endif; ?>
                </form>
            </div>
        <?php endif; ?>

        <h2>4) Exportar planilha (migracao)</h2>
        <p>Baixe um <code>.csv</code> com todos os registros no mesmo layout de importacao do plugin (ideal para levar para outra hospedagem).</p>
        <form method="post" action="<?php echo esc_url(admin_url('admin-post.php')); ?>" style="margin-bottom:24px;">
            <input type="hidden" name="action" value="nsr_export_csv" />
            <?php wp_nonce_field('nsr_export_csv', 'nsr_export_nonce'); ?>
            <button type="submit" class="button button-secondary">Exportar CSV completo</button>
        </form>

        <h2>5) Teste rapido da consulta (admin)</h2>
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

        <h2>6) Consulta no site (navegador)</h2>
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
