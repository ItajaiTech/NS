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
 * Extrai SKUs PRD diretamente do binario do PDF, sem depender do pdftotext.
 *
 * Util como ultimo recurso quando o pdftotext embaralha o texto por causa
 * de layout colunar (PDFs com tabelas complexas em hospedagem Linux).
 *
 * @param string $file_path Caminho absoluto para o arquivo PDF.
 * @return array Itens no mesmo formato de nsr_extract_kdt_items_from_pdf_text().
 */
function nsr_extract_skus_from_pdf_binary($file_path) {
    $binary = @file_get_contents($file_path);
    if ($binary === false || $binary === '') {
        return array();
    }

    $items = array();

    // Busca SKUs no formato literal do PDF: (PRD00069) ou texto plano PRD00069.
    // PDFs gerados pelo Tiny/Bling costumam embutir os codigos como texto literal.
    preg_match_all('/\(PRD(\d{5})\)/i', $binary, $matches_paren, PREG_OFFSET_CAPTURE);
    preg_match_all('/(?<![A-Z0-9])(PRD\d{5})(?![A-Z0-9])/i', $binary, $matches_plain, PREG_OFFSET_CAPTURE);

    $all_skus = array();

    foreach ($matches_paren[0] as $idx => $m) {
        $all_skus[] = array(
            'sku'    => 'PRD' . strtoupper($matches_paren[1][$idx][0]),
            'offset' => $m[1],
        );
    }

    foreach ($matches_plain[0] as $idx => $m) {
        $all_skus[] = array(
            'sku'    => strtoupper($matches_plain[0][$idx][0]),
            'offset' => $m[1],
        );
    }

    foreach ($all_skus as $entry) {
        $sku    = strtoupper(preg_replace('/[^A-Z0-9]/', '', $entry['sku']));
        $offset = $entry['offset'];

        if (!preg_match('/^PRD\d{5}$/', $sku)) {
            continue;
        }

        // Ja encontrado com quantidade valida, pula duplicata.
        if (isset($items[$sku]) && $items[$sku]['quantidade'] > 0) {
            continue;
        }

        // Janela de 120 bytes apos o SKU para encontrar a quantidade.
        $window = substr($binary, $offset + strlen($sku), 120);
        $qty    = 0;

        // Padrao: quantidade entre parenteses seguida de UN/UNID/UNIDADE no PDF literal.
        if (preg_match('/\((\d{1,6}(?:[.,]\d{2})?)\)\s*(?:Tj|TJ)?\s*[^(]{0,20}\((?:UN|UNID|UNIDADE|UND|PC)\)/i', $window, $qm)) {
            $qty = (int) str_replace(array('.', ','), array('', ''), $qm[1]);
        }

        // Padrao alternativo: numero com virgula/ponto antes de UN no texto corrido.
        if ($qty <= 0 && preg_match('/(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?|\d{1,6})\s*(?:UN|UNID|UNIDADE|UND)\b/i', $window, $qm2)) {
            $raw = preg_replace('/\.(?=\d{3})/', '', $qm2[1]); // remove separador de milhar
            $raw = str_replace(',', '.', $raw);
            $qty = (int) floatval($raw);
        }

        // Fallback: primeiro numero plausivel como quantidade na janela.
        if ($qty <= 0 && preg_match_all('/\((\d{1,5})\)/', $window, $qm_all)) {
            foreach ($qm_all[1] as $qv) {
                $candidate = (int) $qv;
                if ($candidate > 0 && $candidate <= 50000) {
                    $qty = $candidate;
                    break;
                }
            }
        }

        if ($qty <= 0 || $qty > 50000) {
            continue;
        }

        $items[$sku] = array(
            'sku'        => $sku,
            'descricao'  => '',
            'quantidade' => $qty,
            'valor'      => '',
            'scanned'    => array(),
        );
    }

    return $items;
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
    // Em Linux, prioriza binario do sistema; em Windows, usa binario embarcado no plugin.
    if (function_exists('shell_exec')) {
        $safe_input = escapeshellarg((string) $file_path);
        $safe_output = escapeshellarg('-');
        $binarios = nsr_get_pdftotext_binaries();
        // CORRECAO: -raw primeiro para evitar embaralhamento de texto em PDFs com tabelas complexas.
        // O -layout agrupa por posicao visual (colunar) e mistura SKU + Qtd + Preco da mesma linha.
        // O -raw le em ordem de fluxo do documento (linha a linha), muito mais confiavel para tabelas.
        $modes = array(
            '-raw -enc UTF-8',                   // 1o: raw — leitura em ordem de fluxo (melhor para tabelas)
            '-raw -nopgbrk -enc UTF-8',          // 2o: raw sem quebra de pagina
            '-table -enc UTF-8',                 // 3o: modo tabela do Poppler (versao 0.89+)
            '-tsv -enc UTF-8',                   // 4o: TSV com coordenadas (Poppler recente)
            '-layout -enc UTF-8',                // 5o: layout como fallback
            '-layout -nopgbrk -enc UTF-8',       // 6o: layout sem quebra de pagina
        );

        foreach ($binarios as $pdftotext_bin) {
            foreach ($modes as $mode) {
                $cmd = '"' . $pdftotext_bin . '" ' . $mode . ' ' . $safe_input . ' ' . $safe_output;
                if (stripos(PHP_OS_FAMILY, 'Windows') === 0) {
                    $cmd .= ' 2>NUL';
                } else {
                    $cmd .= ' 2>/dev/null';
                }

                $result = @shell_exec($cmd);
                if ($result !== null && trim((string) $result) !== '') {
                    // Salva caminho do PDF para uso no fallback binario
                    $GLOBALS['nsr_current_pdf_path'] = $file_path;
                    return trim((string) $result);
                }
            }
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
 * Retorna possiveis caminhos do pdftotext conforme o sistema operacional.
 *
 * @return array
 */
function nsr_get_pdftotext_binaries() {
    $paths = array();
    $is_windows = (stripos(PHP_OS_FAMILY, 'Windows') === 0);

    if ($is_windows) {
        $windows_local = plugin_dir_path(__FILE__) . 'bin/pdftotext.exe';
        if (file_exists($windows_local)) {
            $paths[] = $windows_local;
        }

        $windows_common = array(
            'C:\\Program Files\\poppler\\Library\\bin\\pdftotext.exe',
            'C:\\Program Files (x86)\\poppler\\Library\\bin\\pdftotext.exe',
        );

        foreach ($windows_common as $candidate) {
            if (file_exists($candidate)) {
                $paths[] = $candidate;
            }
        }

        return array_values(array_unique($paths));
    }

    $linux_common = array(
        '/usr/bin/pdftotext',
        '/usr/local/bin/pdftotext',
        '/bin/pdftotext',
    );

    foreach ($linux_common as $candidate) {
        if (is_executable($candidate)) {
            $paths[] = $candidate;
        }
    }

    // fallback por PATH quando nao encontramos caminho fixo.
    $paths[] = 'pdftotext';

    return array_values(array_unique($paths));
}

/**
 * Identifica se token parece SKU valido.
 *
 * Regra de negocio: SKU deve seguir PRD + numeros.
 *
 * @param string $sku
 * @return bool
 */
function nsr_is_valid_prd_sku($sku) {
    $sku = strtoupper(trim((string) $sku));
    return (bool) preg_match('/^PRD\d{5}$/', $sku);
}

/**
 * Identifica se token parece SKU valido.
 *
 * @param string $token
 * @return bool
 */
function nsr_is_probable_sku($token) {
    $token = strtoupper(trim((string) $token));
    if ($token === '' || strlen($token) !== 8) {
        return false;
    }

    // Regra de negocio: SKU sempre inicia com PRD seguido de numeros.
    if (!nsr_is_valid_prd_sku($token)) {
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
 * Aplica deslocamento alfabetico em caracteres A-Z/a-z.
 *
 * @param string $text
 * @param int    $shift
 * @return string
 */
function nsr_shift_alpha_text($text, $shift) {
    $shift = (int) $shift % 26;
    if ($shift === 0) {
        return (string) $text;
    }

    $out = '';
    $len = strlen((string) $text);
    for ($i = 0; $i < $len; $i++) {
        $ch = $text[$i];
        $ord = ord($ch);

        if ($ord >= 65 && $ord <= 90) {
            $base = 65;
            $out .= chr(($ord - $base + $shift + 26) % 26 + $base);
            continue;
        }

        if ($ord >= 97 && $ord <= 122) {
            $base = 97;
            $out .= chr(($ord - $base + $shift + 26) % 26 + $base);
            continue;
        }

        $out .= $ch;
    }

    return $out;
}

/**
 * Recompacta palavras que vieram com um caractere por token no PDF.
 * Ex.: "P R D 0 0 0 6 9" -> "PRD00069".
 *
 * @param string $text
 * @return string
 */
function nsr_compact_spaced_char_tokens($text) {
    $text = (string) $text;

    return (string) preg_replace_callback(
        '/\b(?:[A-Z0-9]\s+){2,}[A-Z0-9]\b/u',
        function ($m) {
            return str_replace(' ', '', (string) $m[0]);
        },
        $text
    );
}

/**
 * Corrige substituicoes comuns de OCR/texto ruidoso em PDFs.
 *
 * Exemplos:
 * 3EDIDO -> PEDIDO | 9ENDA -> VENDA | &LIENTE -> CLIENTE | 4TD -> QTD
 *
 * @param string $text
 * @return string
 */
function nsr_decode_noisy_ocr_text($text) {
    $text = strtoupper(remove_accents((string) $text));

    $map = array(
        '&' => 'C',
        '$' => 'A',
        '3' => 'P',
        '9' => 'V',
        '4' => 'Q',
        '8' => 'U',
        '7' => 'T',
        '1|' => 'N',
        '1/' => 'N',
        '|' => 'I',
        'µ' => 'O',
        '©' => 'C',
    );

    return str_replace(array_keys($map), array_values($map), $text);
}

/**
 * Decodifica tokens com substituicoes comuns vistas em alguns PDFs na hospedagem.
 * Mantem numeros puros (quantidades/valores) sem alteracao.
 *
 * @param string $text
 * @return string
 */
function nsr_decode_obfuscated_pdf_text($text) {
    $text = strtoupper(remove_accents((string) $text));
    $tokens = preg_split('/(\s+)/u', $text, -1, PREG_SPLIT_DELIM_CAPTURE);
    if (!is_array($tokens)) {
        return $text;
    }

    // Mapeamento observado no fornecedor/hospedagem (ex.: 3RECO -> PRECO, 4TD -> QTD).
    $map = array(
        '&' => 'C',
        '*' => 'G',
        '$' => 'A',
        '©' => 'C',
        'µ' => 'O',
        '|' => 'I',
        '\\' => 'I',
    );

    $digit_map = array(
        '0' => 'O',
        '1' => 'N',
        '2' => 'O',
        '3' => 'P',
        '4' => 'Q',
        '5' => 'R',
        '6' => 'S',
        '7' => 'T',
        '8' => 'U',
        '9' => 'V',
    );

    foreach ($tokens as $idx => $token) {
        if (trim((string) $token) === '') {
            continue;
        }

        // Mantem numeros puros (inteiros/decimais) para nao quebrar quantidades.
        if (preg_match('/^\d+(?:[\.,]\d+)?$/', $token)) {
            continue;
        }

        // Mantem tokens que ja parecem SKU PRD para nao corromper digitos.
        if (preg_match('/^PRD[A-Z0-9]{2,}$/', $token)) {
            continue;
        }

        $chars = preg_split('//u', (string) $token, -1, PREG_SPLIT_NO_EMPTY);
        if (!is_array($chars)) {
            continue;
        }

        $rebuilt = '';
        $count = count($chars);

        for ($i = 0; $i < $count; $i++) {
            $ch = (string) $chars[$i];
            $prev = $i > 0 ? (string) $chars[$i - 1] : '';
            $next = $i + 1 < $count ? (string) $chars[$i + 1] : '';

            if (isset($map[$ch])) {
                $rebuilt .= $map[$ch];
                continue;
            }

            if (isset($digit_map[$ch])) {
                $prev_is_digit = preg_match('/\d/u', $prev) === 1;
                $next_is_digit = preg_match('/\d/u', $next) === 1;

                // Troca digito por letra apenas quando estiver em contexto alfanumerico,
                // evitando alterar sequencias numericas reais de quantidade/valor.
                if (!$prev_is_digit || !$next_is_digit) {
                    $prev_is_letter = preg_match('/[A-Z]/u', $prev) === 1;
                    $next_is_letter = preg_match('/[A-Z]/u', $next) === 1;
                    if ($prev_is_letter || $next_is_letter) {
                        $rebuilt .= $digit_map[$ch];
                        continue;
                    }
                }
            }

            $rebuilt .= $ch;
        }

        $tokens[$idx] = $rebuilt;
    }

    return implode('', $tokens);
}

/**
 * Gera candidatos de texto para parser de PDF (original + fallback decifrado).
 *
 * @param string $text
 * @return array
 */
function nsr_get_pdf_text_candidates($text) {
    $text = (string) $text;
    $candidates = array($text);

    $decoded_obfuscated = nsr_decode_obfuscated_pdf_text($text);
    if ($decoded_obfuscated !== '' && $decoded_obfuscated !== $text) {
        $candidates[] = $decoded_obfuscated;
    }

    $upper_original = strtoupper(remove_accents($decoded_obfuscated !== '' ? $decoded_obfuscated : $text));

    $compacted = nsr_compact_spaced_char_tokens($upper_original);
    if ($compacted !== '' && $compacted !== $upper_original) {
        $candidates[] = $compacted;
    }

    $alnum_spaced = trim((string) preg_replace('/\s+/', ' ', (string) preg_replace('/[^A-Z0-9]+/', ' ', $upper_original)));
    $alnum_compacted = nsr_compact_spaced_char_tokens($alnum_spaced);
    if ($alnum_compacted !== '' && $alnum_compacted !== $upper_original) {
        $candidates[] = $alnum_compacted;
    }

    $ocr_decoded = nsr_decode_noisy_ocr_text($upper_original);
    if ($ocr_decoded !== '' && $ocr_decoded !== $upper_original) {
        $candidates[] = $ocr_decoded;

        $ocr_compacted = nsr_compact_spaced_char_tokens($ocr_decoded);
        if ($ocr_compacted !== '' && $ocr_compacted !== $ocr_decoded) {
            $candidates[] = $ocr_compacted;
        }

        $ocr_alnum = trim((string) preg_replace('/\s+/', ' ', (string) preg_replace('/[^A-Z0-9]+/', ' ', $ocr_decoded)));
        if ($ocr_alnum !== '' && $ocr_alnum !== $ocr_decoded) {
            $candidates[] = $ocr_alnum;
        }
    }

    $upper = strtoupper(remove_accents($text));
    $encoded_keywords = array('SURGXWR', 'TXDQWLGDGH', 'FRGLJR', 'GHVFULFDR', 'SHGLGR', 'QRWD', 'ILVFDO');
    $plain_keywords = array('PRODUTO', 'QUANTIDADE', 'CODIGO', 'DESCRICAO', 'PEDIDO', 'NOTA', 'FISCAL');

    $encoded_score = 0;
    foreach ($encoded_keywords as $kw) {
        if (strpos($upper, $kw) !== false) {
            $encoded_score++;
        }
    }

    $plain_score = 0;
    foreach ($plain_keywords as $kw) {
        if (strpos($upper, $kw) !== false) {
            $plain_score++;
        }
    }

    // Alguns PDFs retornam texto com cifragem de deslocamento (+3). Nesse caso,
    // tentamos uma versao com shift -3 para recuperar cabecalhos e itens.
    if ($encoded_score > $plain_score) {
        $decoded = nsr_shift_alpha_text($text, -3);
        if ($decoded !== '' && $decoded !== $text) {
            array_unshift($candidates, $decoded);

            $decoded_compacted = nsr_compact_spaced_char_tokens(strtoupper(remove_accents($decoded)));
            if ($decoded_compacted !== '' && $decoded_compacted !== strtoupper(remove_accents($decoded))) {
                array_unshift($candidates, $decoded_compacted);
            }

            $decoded_ocr = nsr_decode_noisy_ocr_text($decoded);
            if ($decoded_ocr !== '' && $decoded_ocr !== strtoupper(remove_accents($decoded))) {
                array_unshift($candidates, $decoded_ocr);
            }

            $decoded_alnum = trim((string) preg_replace('/\s+/', ' ', (string) preg_replace('/[^A-Z0-9]+/', ' ', strtoupper(remove_accents($decoded)))));
            $decoded_alnum_compacted = nsr_compact_spaced_char_tokens($decoded_alnum);
            if ($decoded_alnum_compacted !== '' && $decoded_alnum_compacted !== strtoupper(remove_accents($decoded))) {
                array_unshift($candidates, $decoded_alnum_compacted);
            }
        }
    }

    return array_values(array_unique($candidates));
}

/**
 * Extrai itens PRD de texto ruidoso (caracteres quebrados/simbolos entre letras).
 *
 * @param string $text
 * @return array
 */
function nsr_extract_prd_items_from_noisy_text($text) {
    $items = array();
    $normalized = strtoupper(remove_accents((string) $text));
    $normalized = str_replace(array("\r", "\n", "\t"), ' ', $normalized);

    // Aceita apenas PRD literal com separadores ruidosos entre caracteres.
    if (!preg_match_all('/P\W*R\W*D\W*((?:\d\W*){5})/u', $normalized, $sku_matches, PREG_OFFSET_CAPTURE)) {
        return $items;
    }

    foreach ($sku_matches[1] as $sku_match) {
        $digits_raw = isset($sku_match[0]) ? (string) $sku_match[0] : '';
        $digits = preg_replace('/\D+/', '', $digits_raw);
        if ($digits === '' || strlen($digits) !== 5) {
            continue;
        }

        $sku = 'PRD' . $digits;
        if (!nsr_is_probable_sku($sku)) {
            continue;
        }

        $offset = isset($sku_match[1]) ? (int) $sku_match[1] : 0;
        $window = substr($normalized, $offset, 240);
        $qty = 0;

        if (preg_match('/\b(?:UNIDADE|UNID|UND|UN)\b/u', (string) $window) !== 1) {
            continue;
        }

        if (preg_match('/\b(\d{1,3}(?:\.\d{3})*(?:,\d{2})?|\d{1,6})\s*(?:UNIDADE|UNID|UND|UN)\b/u', (string) $window, $qm_unit)) {
            $qty = nsr_parse_quantity_value($qm_unit[1]);
        }

        if ($qty <= 0 && preg_match_all('/\b\d{1,14}(?:[\.,]\d{2})?\b/u', (string) $window, $num_matches)) {
            foreach ($num_matches[0] as $num_token) {
                $digits_only = preg_replace('/\D+/', '', (string) $num_token);
                if (strlen($digits_only) >= 8 && strpos((string) $num_token, ',') === false && strpos((string) $num_token, '.') === false) {
                    // Provavel GTIN/chave numerica longa, nao quantidade.
                    continue;
                }

                $candidate_qty = nsr_parse_quantity_value($num_token);
                if ($candidate_qty > 0 && $candidate_qty <= 50000) {
                    $qty = $candidate_qty;
                    break;
                }
            }
        }

        if ($qty <= 0) {
            continue;
        }

        if (!isset($items[$sku])) {
            $items[$sku] = array(
                'sku'        => $sku,
                'descricao'  => '',
                'quantidade' => 0,
                'valor'      => '',
                'scanned'    => array(),
            );
        }

        $items[$sku]['quantidade'] += $qty;
    }

    return $items;
}

/**
 * Extrai itens do layout KeepData quando o PDF retorna texto muito ruidoso.
 *
 * Estrategia:
 * - decodifica substituicoes comuns (3->P, 7->T, etc)
 * - foca no bloco entre "CODIGO" e "NUMERO DE ITENS"
 * - busca pares [codigo] + [quantidade] associados a "UN"
 *
 * @param string $text
 * @return array
 */
function nsr_extract_keepdata_items_from_noisy_text($text) {
    $items = array();

    $decoded = nsr_decode_obfuscated_pdf_text($text);
    $decoded = nsr_decode_noisy_ocr_text($decoded);
    $normalized = strtoupper(remove_accents((string) $decoded));
    $normalized = preg_replace('/[^A-Z0-9\s\.,]/', ' ', (string) $normalized);
    $normalized = preg_replace('/\s+/', ' ', (string) $normalized);
    $normalized = trim((string) $normalized);

    if ($normalized === '') {
        return $items;
    }

    if (strpos($normalized, 'CODIGO') === false || strpos($normalized, 'QTD') === false || strpos($normalized, 'UN') === false) {
        return $items;
    }

    $start_pos = strpos($normalized, 'CODIGO');
    if ($start_pos !== false) {
        $normalized = substr($normalized, $start_pos);
    }

    $end_pos = strpos($normalized, 'NUMERO DE ITENS');
    if ($end_pos !== false) {
        $normalized = substr($normalized, 0, $end_pos);
    }

    if ($normalized === '') {
        return $items;
    }

    // Insere separadores para facilitar split de blocos de item.
    $normalized = preg_replace('/\b(PRD\d{1,5}|\d{3,5})\b\s+(\d{1,3}(?:\.\d{3})*(?:,\d{2})?)\s+UN\b/', ' ||ITEM|| $1 $2 UN ', (string) $normalized);
    $chunks = array_filter(array_map('trim', explode('||ITEM||', $normalized)));

    foreach ($chunks as $chunk) {
        if (!preg_match('/\b(PRD\d{1,5}|\d{3,5})\b/', $chunk, $m_code)) {
            continue;
        }

        $code_raw = strtoupper((string) $m_code[1]);
        if (strpos($code_raw, 'PRD') === 0) {
            $digits = preg_replace('/\D+/', '', substr($code_raw, 3));
        } else {
            $digits = preg_replace('/\D+/', '', $code_raw);
        }

        if ($digits === '') {
            continue;
        }

        // KeepData trabalha com SKU PRD + 5 digitos.
        $digits = str_pad(substr($digits, -5), 5, '0', STR_PAD_LEFT);
        $sku = 'PRD' . $digits;
        if (!nsr_is_probable_sku($sku)) {
            continue;
        }

        $qty = 0;
        if (preg_match('/\b(\d{1,3}(?:\.\d{3})*(?:,\d{2})?)\s+UN\b/', $chunk, $m_qty)) {
            $qty = nsr_parse_quantity_value($m_qty[1]);
        }

        if ($qty <= 0) {
            continue;
        }

        if (!isset($items[$sku])) {
            $items[$sku] = array(
                'sku'        => $sku,
                'descricao'  => '',
                'quantidade' => 0,
                'valor'      => '',
                'scanned'    => array(),
            );
        }

        $items[$sku]['quantidade'] += $qty;
    }

    return $items;
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
    $normalized = nsr_compact_spaced_char_tokens($normalized);
    $normalized = preg_replace('/\s+/', ' ', $normalized);

    // Captura sequencias como:
    // PRD00040 7898722600065 988,00 UN 210,00 ...
    // PRD00040 988,00 PC 210,00 ...
    // PRD00040 988,00 210,00 ...
    // Grupos: (1) SKU PRD+numeros  (2) Qtd  (3) Preco unitario (opcional)
    $pattern = '/\b(PRD\d{5})\b\s+(?:\d{8,14}\s+)?(\d{1,3}(?:\.\d{3})*(?:,\d{2})?|\d{1,6})(?:\s+(?:UN(?:ID(?:AD(?:E)?)?)?|UND|PC|PCT))?\b(?:\s+(\d{1,3}(?:\.\d{3})*(?:,\d{2})?))?/u';
    if (preg_match_all($pattern, $normalized, $matches, PREG_SET_ORDER)) {
        foreach ($matches as $m) {
            $sku = strtoupper((string) $m[1]);
            $qty = nsr_parse_quantity_value($m[2]);
            $valor_str = isset($m[3]) ? trim($m[3]) : '';

            if (!nsr_is_probable_sku($sku) || $qty <= 0) {
                continue;
            }

            if (!isset($items[$sku])) {
                $items[$sku] = array(
                    'sku'       => $sku,
                    'descricao' => '',
                    'quantidade' => 0,
                    'valor'     => '',
                    'scanned'   => array(),
                );
            }

            $items[$sku]['quantidade'] += $qty;
            if ($valor_str !== '' && $items[$sku]['valor'] === '') {
                $items[$sku]['valor'] = $valor_str;
            }
        }
    }

    // Fallback quando a linha nao vem no formato completo com "UN":
    // captura PRD e tenta encontrar a quantidade mais proxima no trecho seguinte.
    if (empty($items) && preg_match_all('/\b(PRD\d{5})\b/u', $normalized, $sku_matches, PREG_OFFSET_CAPTURE)) {
        foreach ($sku_matches[1] as $sku_match) {
            $sku = strtoupper((string) $sku_match[0]);
            $offset = isset($sku_match[1]) ? (int) $sku_match[1] : 0;
            if (!nsr_is_probable_sku($sku)) {
                continue;
            }

            $window = substr($normalized, $offset + strlen($sku), 200);
            $qty = 0;

            if (preg_match('/\b(\d{1,3}(?:\.\d{3})*(?:,\d{2})?|\d{1,6})\s*(?:UNIDADE|UNID|UND|UN)\b/u', (string) $window, $qm_unit)) {
                $qty = nsr_parse_quantity_value($qm_unit[1]);
            }

            if ($qty <= 0 && preg_match_all('/\b(\d{1,3}(?:\.\d{3})*(?:,\d{2})?|\d{1,6})\b/u', (string) $window, $qm_all)) {
                foreach ($qm_all[1] as $qv) {
                    $candidate_qty = nsr_parse_quantity_value($qv);
                    if ($candidate_qty > 0 && $candidate_qty <= 50000) {
                        $qty = $candidate_qty;
                        break;
                    }
                }
            }

            if ($qty <= 0) {
                continue;
            }

            if (!isset($items[$sku])) {
                $items[$sku] = array(
                    'sku'       => $sku,
                    'descricao' => '',
                    'quantidade'=> 0,
                    'valor'     => '',
                    'scanned'   => array(),
                );
            }

            $items[$sku]['quantidade'] += $qty;
        }
    }

    if (empty($items)) {
        $items = nsr_extract_prd_items_from_noisy_text($normalized);
    }

    if (empty($items)) {
        $items = nsr_extract_keepdata_items_from_noisy_text($normalized);
    }

    // Ultimo recurso: extrai SKUs diretamente do binario do PDF.
    // Usado quando o pdftotext embaralha o texto por causa de layout colunar no Linux.
    if (empty($items) && !empty($GLOBALS['nsr_current_pdf_path'])) {
        $binary_items = nsr_extract_skus_from_pdf_binary((string) $GLOBALS['nsr_current_pdf_path']);
        if (!empty($binary_items)) {
            $items = $binary_items;
        }
    }

    // Filtro final defensivo: garante apenas SKU valido e quantidade plausivel.
    foreach (array_keys($items) as $sku_key) {
        $qty_value = isset($items[$sku_key]['quantidade']) ? (int) $items[$sku_key]['quantidade'] : 0;
        if (!nsr_is_probable_sku((string) $sku_key) || $qty_value <= 0 || $qty_value > 50000) {
            unset($items[$sku_key]);
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
function nsr_extract_order_from_pdf_text_single($text) {
    $raw_lines = preg_split('/\r\n|\r|\n/', (string) $text);
    $lines = array();

    foreach ($raw_lines as $line) {
        $clean = trim(preg_replace('/\s+/', ' ', (string) $line));
        if ($clean !== '') {
            $lines[] = $clean;
        }
    }

    $pedido      = '';
    $nota_fiscal = '';
    $data_venda  = '';

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

        // Data do pedido: "Data  28/04/2026" ou "Data: 28/04/2026"
        if ($data_venda === '' && preg_match('/\bDATA\b[^0-9]{0,5}(\d{2}\/\d{2}\/\d{4})/', $norm, $m)) {
            // Converte DD/MM/YYYY -> YYYY-MM-DD para armazenamento padrao.
            $parts = explode('/', $m[1]);
            $data_venda = $parts[2] . '-' . $parts[1] . '-' . $parts[0];
        }
    }

    $items = nsr_extract_kdt_items_from_pdf_text($text);

    // Fallback generico para outros formatos, caso o padrao KDT nao encontre itens.
    if (!empty($items)) {
        return array(
            'pedido'      => $pedido,
            'nota_fiscal' => $nota_fiscal,
            'data_venda'  => $data_venda,
            'itens'       => $items,
        );
    }

    $items = array();

    foreach ($lines as $line) {
        $norm = strtoupper(remove_accents($line));

        $sku = '';
        $qty = 0;

        if (preg_match('/\b(?:SKU|CODIGO(?:\s+DO\s+PRODUTO)?|COD\.)\b\s*[:\-]?\s*([A-Z0-9._\/-]{3,})/', $norm, $m)) {
            $sku = strtoupper($m[1]);
            if (!nsr_is_probable_sku($sku)) {
                $sku = '';
            }
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

        if ($sku === '' || !nsr_is_probable_sku($sku) || $qty <= 0) {
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
        'pedido'      => $pedido,
        'nota_fiscal' => $nota_fiscal,
        'data_venda'  => $data_venda,
        'itens'       => $items,
    );
}

/**
 * Extrai pedido, nota fiscal e itens (SKU + quantidade) do texto de PDF.
 *
 * @param string $text
 * @return array
 */
function nsr_extract_order_from_pdf_text($text) {
    $candidates = nsr_get_pdf_text_candidates($text);
    $best = null;

    foreach ($candidates as $candidate_text) {
        $parsed = nsr_extract_order_from_pdf_text_single($candidate_text);
        if (!empty($parsed['itens'])) {
            return $parsed;
        }

        if ($best === null) {
            $best = $parsed;
        }
    }

    if ($best !== null) {
        return $best;
    }

    return array(
        'pedido'      => '',
        'nota_fiscal' => '',
        'data_venda'  => '',
        'itens'       => array(),
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
    $itens = isset($session['itens']) && is_array($session['itens']) ? $session['itens'] : array();

    // Defesa final: nunca manter SKU fora do padrao PRD+numeros na sessao.
    foreach (array_keys($itens) as $sku_key) {
        if (!nsr_is_probable_sku((string) $sku_key)) {
            unset($itens[$sku_key]);
        }
    }

    $session['itens'] = $itens;
    $skus = array_keys($itens);
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
 * Busca registros para o painel admin por NS, NF ou Pedido.
 *
 * @param string $value
 * @param string $search_type ns|nf|pedido
 * @param bool $partial
 * @param int $limit
 * @return array
 */
function nsr_find_admin_records($value, $search_type = 'ns', $partial = false, $limit = 100) {
    global $wpdb;

    $value_raw = trim((string) $value);
    $value_normalized = nsr_normalize_lookup_value($value_raw);
    if ($value_raw === '' && $value_normalized === '') {
        return array();
    }

    $search_type = strtolower(trim((string) $search_type));
    $allowed_types = array(
        'ns' => 'NS',
        'nf' => 'NF',
        'pedido' => 'Pedido',
    );
    if (!isset($allowed_types[$search_type])) {
        $search_type = 'ns';
    }

    $table_name = nsr_get_table_name();
    $limit = max(1, min(absint($limit), 500));

    if ($search_type === 'ns') {
        return nsr_find_by_ns($value_raw, $partial, $limit);
    }

    $column = ($search_type === 'nf') ? 'nota_fiscal' : 'pedido';

    if ($partial) {
        $like_raw = '%' . $wpdb->esc_like($value_raw) . '%';

        if ($value_normalized !== '' && $value_normalized !== $value_raw) {
            $like_normalized = '%' . $wpdb->esc_like($value_normalized) . '%';
            $sql = $wpdb->prepare(
                "SELECT ns, nota_fiscal, pedido, sku, descricao, quantidade, valor, data_venda, origem_arquivo, updated_at
                 FROM {$table_name}
                 WHERE {$column} LIKE %s OR {$column} LIKE %s
                 ORDER BY updated_at DESC, id DESC
                 LIMIT %d",
                $like_raw,
                $like_normalized,
                $limit
            );
        } else {
            $sql = $wpdb->prepare(
                "SELECT ns, nota_fiscal, pedido, sku, descricao, quantidade, valor, data_venda, origem_arquivo, updated_at
                 FROM {$table_name}
                 WHERE {$column} LIKE %s
                 ORDER BY updated_at DESC, id DESC
                 LIMIT %d",
                $like_raw,
                $limit
            );
        }

        return $wpdb->get_results($sql, ARRAY_A);
    }

    if ($value_normalized !== '' && $value_normalized !== $value_raw) {
        $sql = $wpdb->prepare(
            "SELECT ns, nota_fiscal, pedido, sku, descricao, quantidade, valor, data_venda, origem_arquivo, updated_at
             FROM {$table_name}
             WHERE {$column} = %s OR {$column} = %s
             ORDER BY updated_at DESC, id DESC
             LIMIT %d",
            $value_raw,
            $value_normalized,
            $limit
        );
    } else {
        $sql = $wpdb->prepare(
            "SELECT ns, nota_fiscal, pedido, sku, descricao, quantidade, valor, data_venda, origem_arquivo, updated_at
             FROM {$table_name}
             WHERE {$column} = %s
             ORDER BY updated_at DESC, id DESC
             LIMIT %d",
            ($value_normalized !== '') ? $value_normalized : $value_raw,
            $limit
        );
    }

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
 * Processa cadastro manual de produto (SKU x descricao).
 *
 * @return array
 */
function nsr_handle_product_manual_submission() {
    global $wpdb;

    $messages = array(
        'success' => array(),
        'error' => array(),
    );

    if (!isset($_POST['nsr_product_manual_submit'])) {
        return $messages;
    }

    if (!current_user_can('manage_options')) {
        $messages['error'][] = 'Permissao insuficiente para cadastrar produto.';
        return $messages;
    }

    check_admin_referer('nsr_product_manual', 'nsr_product_manual_nonce');

    $sku = strtoupper(sanitize_text_field(wp_unslash($_POST['nsr_product_sku'] ?? '')));
    $descricao = sanitize_text_field(wp_unslash($_POST['nsr_product_descricao'] ?? ''));

    if ($sku === '') {
        $messages['error'][] = 'Informe o SKU para cadastrar.';
        return $messages;
    }

    if (!nsr_is_valid_prd_sku($sku)) {
        $messages['error'][] = 'SKU invalido. Use o padrao PRD + numeros (ex.: PRD00069).';
        return $messages;
    }

    $products_table = nsr_get_products_table_name();
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
        $messages['error'][] = 'Nao foi possivel salvar o produto informado.';
        return $messages;
    }

    $messages['success'][] = sprintf('Produto %s salvo com sucesso.', $sku);
    return $messages;
}

/**
 * Processa upload de PDF e fluxo de bipagem de NS.
 *
 * @return array
 */
/**
 * Extrai pedido, nota fiscal e itens (SKU + quantidade) de um XML de NF-e 4.0.
 *
 * @param string $file_path Caminho absoluto para o arquivo XML.
 * @return array|WP_Error
 */
function nsr_parse_xml_nfe($file_path) {
    $xml_content = @file_get_contents($file_path);
    if ($xml_content === false || trim($xml_content) === '') {
        return new WP_Error('nsr_xml_read_error', 'Nao foi possivel ler o arquivo XML enviado.');
    }

    libxml_use_internal_errors(true);
    $xml = simplexml_load_string($xml_content, 'SimpleXMLElement', LIBXML_NOCDATA);
    libxml_clear_errors();

    if ($xml === false) {
        return new WP_Error('nsr_xml_parse_error', 'Arquivo XML invalido ou corrompido.');
    }

    $ns = 'http://www.portalfiscal.inf.br/nfe';

    // Suporta <nfeProc><NFe> e <NFe> direto
    $nfe_node = null;
    $local_name = $xml->getName();
    if ($local_name === 'nfeProc') {
        $children = $xml->children($ns);
        $nfe_node = isset($children->NFe) ? $children->NFe : null;
        if ($nfe_node === null) {
            // tenta sem namespace
            $nfe_node = isset($xml->NFe) ? $xml->NFe : null;
        }
    } else {
        $nfe_node = $xml;
    }

    if ($nfe_node === null) {
        return new WP_Error('nsr_xml_structure_error', 'Estrutura do XML nao reconhecida como NF-e valida.');
    }

    $infNFe = $nfe_node->children($ns)->infNFe;
    if ($infNFe === null || !isset($infNFe->ide)) {
        // fallback sem namespace
        $infNFe = isset($nfe_node->infNFe) ? $nfe_node->infNFe : null;
    }
    if ($infNFe === null) {
        return new WP_Error('nsr_xml_structure_error', 'Elemento infNFe nao encontrado no XML.');
    }

    // IDE — numero da NF e data
    $ide = $infNFe->children($ns)->ide;
    if ($ide === null) {
        $ide = isset($infNFe->ide) ? $infNFe->ide : null;
    }

    $nota_fiscal = $ide !== null && isset($ide->nNF) ? trim((string) $ide->nNF) : '';
    $data_venda  = '';
    if ($ide !== null && isset($ide->dhEmi)) {
        $dt = date_create(trim((string) $ide->dhEmi));
        if ($dt !== false) {
            $data_venda = date_format($dt, 'd/m/Y');
        }
    }

    // Itens (det)
    $det_list = $infNFe->children($ns)->det;
    if ($det_list === null || count($det_list) === 0) {
        $det_list = isset($infNFe->det) ? $infNFe->det : null;
    }
    if ($det_list === null || count($det_list) === 0) {
        return new WP_Error('nsr_xml_no_items', 'Nenhum item (det) encontrado no XML da NF-e.');
    }

    $pedido = '';
    $itens  = array();

    foreach ($det_list as $det) {
        $prod = $det->children($ns)->prod;
        if ($prod === null || count($prod) === 0) {
            $prod = isset($det->prod) ? $det->prod : null;
        }
        if ($prod === null) {
            continue;
        }

        $sku       = strtoupper(trim((string) $prod->cProd));
        $qty_raw   = trim((string) $prod->qCom);
        $descricao = trim((string) $prod->xProd);
        $valor_str = trim((string) $prod->vUnCom);

        if ($pedido === '' && isset($prod->xPed) && trim((string) $prod->xPed) !== '') {
            $pedido = trim((string) $prod->xPed);
        }

        // XML pode vir com cProd sem o padrao interno PRD000XX.
        // Quando isso acontecer, tenta extrair SKU valido da descricao do item.
        if (!nsr_is_valid_prd_sku($sku)) {
            $desc_upper = strtoupper(remove_accents($descricao));
            if (preg_match('/\b(PRD\d{5})\b/', $desc_upper, $m_sku_desc)) {
                $sku = strtoupper((string) $m_sku_desc[1]);
            }
        }

        if (!nsr_is_valid_prd_sku($sku)) {
            continue;
        }

        $qty = (int) floatval(str_replace(',', '.', $qty_raw));
        if ($qty <= 0) {
            continue;
        }

        if (!isset($itens[$sku])) {
            $itens[$sku] = array(
                'sku'        => $sku,
                'descricao'  => $descricao,
                'quantidade' => 0,
                'valor'      => $valor_str,
                'scanned'    => array(),
            );
        }
        $itens[$sku]['quantidade'] += $qty;
    }

    if (empty($itens)) {
        return new WP_Error('nsr_xml_no_skus', 'Nenhum SKU/quantidade encontrado no XML da NF-e.');
    }

    return array(
        'pedido'      => $pedido,
        'nota_fiscal' => $nota_fiscal,
        'data_venda'  => $data_venda,
        'itens'       => $itens,
    );
}

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
                } elseif (!in_array(strtolower(pathinfo($safe_name, PATHINFO_EXTENSION)), array('pdf', 'xml'), true)) {
                    $messages['error'][] = 'Formato nao suportado. Envie um arquivo .pdf ou .xml (NF-e).';
                } elseif (strtolower(pathinfo($safe_name, PATHINFO_EXTENSION)) === 'xml') {
                    // --- Processamento de XML NF-e ---
                    $parsed_xml = nsr_parse_xml_nfe($tmp_name);
                    if (is_wp_error($parsed_xml)) {
                        $messages['error'][] = $parsed_xml->get_error_message();
                    } elseif (empty($parsed_xml['itens'])) {
                        $messages['error'][] = 'Nenhum item encontrado no XML da NF-e.';
                    } else {
                        $token = nsr_generate_scan_session_token();
                        $session = array(
                            'session_token'  => $token,
                            'pedido'         => (string) $parsed_xml['pedido'],
                            'nota_fiscal'    => (string) $parsed_xml['nota_fiscal'],
                            'data_venda'     => isset($parsed_xml['data_venda']) ? (string) $parsed_xml['data_venda'] : '',
                            'origem_arquivo' => $safe_name,
                            'itens'          => $parsed_xml['itens'],
                            'missing_skus'   => array(),
                        );
                        $session = nsr_recompute_scan_session_flags($session);
                        if (!nsr_save_scan_session($token, $session)) {
                            $messages['error'][] = 'Nao foi possivel salvar a sessao de bipagem.';
                        } else {
                            $active_token = $token;
                            $messages['success'][] = sprintf(
                                'XML %s lido com sucesso. %d SKU(s) encontrado(s). Inicie a bipagem dos NS.',
                                esc_html($safe_name),
                                count($parsed_xml['itens'])
                            );
                        }
                    }
                } else {
                    // --- Processamento de PDF ---
                    $GLOBALS['nsr_current_pdf_path'] = $tmp_name;
                    $text = nsr_read_pdf_text($tmp_name);
                    if (is_wp_error($text)) {
                        $messages['error'][] = $text->get_error_message();
                    } else {
                        $parsed = nsr_extract_order_from_pdf_text($text);
                        if (empty($parsed['itens'])) {
                            $candidates = nsr_get_pdf_text_candidates($text);
                            $best_sample_text = !empty($candidates) ? (string) $candidates[0] : (string) $text;
                            $sample = esc_html(substr($best_sample_text, 0, 400));
                            $messages['error'][] = 'Nao foi possivel detectar SKU e quantidade no PDF. Texto extraido (amostra): ' . $sample;
                        } else {
                            $token = nsr_generate_scan_session_token();
                            $session = array(
                                'session_token'  => $token,
                                'pedido'         => (string) $parsed['pedido'],
                                'nota_fiscal'    => (string) $parsed['nota_fiscal'],
                                'data_venda'     => isset($parsed['data_venda']) ? (string) $parsed['data_venda'] : '',
                                'origem_arquivo' => $safe_name,
                                'itens'          => $parsed['itens'],
                                'missing_skus'   => array(),
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
                if ($sku_raw === '' || !nsr_is_valid_prd_sku($sku_raw)) {
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
                $messages['success'][] = 'Aviso: existem SKUs sem cadastro de produto, mas a finalizacao sera permitida.';
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
                    $descricao   = isset($item['descricao']) ? (string) $item['descricao'] : '';
                    $expected    = (int) $item['quantidade'];
                    $valor_item  = isset($item['valor']) ? (string) $item['valor'] : '';
                    $scanned_list = isset($item['scanned']) && is_array($item['scanned']) ? $item['scanned'] : array();
                    $data_venda  = (isset($session['data_venda']) && (string) $session['data_venda'] !== '') ? (string) $session['data_venda'] : date_i18n('Y-m-d');

                    foreach ($scanned_list as $ns) {
                        $ns_normalizado = nsr_normalize_lookup_value($ns);
                        if ($ns_normalizado === '') {
                            $failed = true;
                            break;
                        }

                        $affected = nsr_upsert_ns_record(array(
                            'ns'             => $ns,
                            'ns_normalizado' => $ns_normalizado,
                            'nota_fiscal'    => (string) $session['nota_fiscal'],
                            'pedido'         => (string) $session['pedido'],
                            'sku'            => $sku,
                            'descricao'      => $descricao,
                            'quantidade'     => (string) $expected,
                            'valor'          => $valor_item,
                            'data_venda'     => $data_venda,
                            'origem_arquivo' => (string) $session['origem_arquivo'],
                            'linha_origem'   => 0,
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

                    $tiny_result = nsr_send_serials_to_tiny_order($session);
                    if (is_wp_error($tiny_result)) {
                        $messages['success'][] = 'Aviso Tiny: ' . $tiny_result->get_error_message();
                    } elseif (is_array($tiny_result) && isset($tiny_result['ok']) && $tiny_result['ok']) {
                        $messages['success'][] = sprintf(
                            'Tiny atualizado com sucesso no pedido %s (%d NS enviado(s)).',
                            isset($tiny_result['order_id']) ? (string) $tiny_result['order_id'] : '-',
                            isset($tiny_result['serials_count']) ? (int) $tiny_result['serials_count'] : 0
                        );
                    }

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

// ============================================================
// AJAX: Bipar NS (sem reload de pagina)
// ============================================================
function nsr_ajax_scan_ns() {
    if (!current_user_can('manage_options')) {
        wp_send_json_error(array('msg' => 'Permissao insuficiente.'), 403);
    }

    check_ajax_referer('nsr_ajax_scan', 'nonce');

    $token = sanitize_text_field(wp_unslash(isset($_POST['token']) ? $_POST['token'] : ''));
    $sku   = strtoupper(sanitize_text_field(wp_unslash(isset($_POST['sku']) ? $_POST['sku'] : '')));
    $ns    = sanitize_text_field(wp_unslash(isset($_POST['ns']) ? $_POST['ns'] : ''));

    $session = nsr_get_scan_session($token);
    if (empty($session)) {
        wp_send_json_error(array('msg' => 'Sessao nao encontrada ou expirada.'), 404);
    }

    $ns_normalizado = nsr_normalize_lookup_value($ns);
    if ($sku === '' || $ns_normalizado === '') {
        wp_send_json_error(array('msg' => 'Informe SKU e NS.'), 400);
    }
    if (!isset($session['itens'][$sku])) {
        wp_send_json_error(array('msg' => 'SKU ' . $sku . ' nao faz parte do pedido.'), 400);
    }

    // Checa duplicidade em toda a sessao.
    foreach ($session['itens'] as $item) {
        $scanned = isset($item['scanned']) && is_array($item['scanned']) ? $item['scanned'] : array();
        foreach ($scanned as $s) {
            if (nsr_normalize_lookup_value($s) === $ns_normalizado) {
                wp_send_json_error(array('msg' => 'NS ja bipado nesta sessao: ' . $ns), 409);
            }
        }
    }

    $expected = (int) $session['itens'][$sku]['quantidade'];
    $already  = isset($session['itens'][$sku]['scanned']) ? count($session['itens'][$sku]['scanned']) : 0;
    if ($already >= $expected) {
        wp_send_json_error(array('msg' => sprintf('SKU %s ja atingiu a quantidade esperada (%d).', $sku, $expected)), 409);
    }

    if (!isset($session['itens'][$sku]['scanned']) || !is_array($session['itens'][$sku]['scanned'])) {
        $session['itens'][$sku]['scanned'] = array();
    }
    $session['itens'][$sku]['scanned'][] = $ns;

    // Persiste pedido/NF se vieram junto.
    $pedido = sanitize_text_field(wp_unslash(isset($_POST['pedido']) ? $_POST['pedido'] : ''));
    $nf     = sanitize_text_field(wp_unslash(isset($_POST['nota_fiscal']) ? $_POST['nota_fiscal'] : ''));
    if ($pedido !== '') {
        $session['pedido'] = $pedido;
    }
    if ($nf !== '') {
        $session['nota_fiscal'] = $nf;
    }

    nsr_save_scan_session($token, $session);

    $scanned_list  = $session['itens'][$sku]['scanned'];
    $scanned_count = count($scanned_list);

    wp_send_json_success(array(
        'sku'           => $sku,
        'ns'            => $ns,
        'scanned_count' => $scanned_count,
        'expected'      => $expected,
        'is_ok'         => ($scanned_count === $expected),
        'scanned_list'  => $scanned_list,
    ));
}
add_action('wp_ajax_nsr_scan_ns', 'nsr_ajax_scan_ns');

/**
 * Separa NS em prefixo + sufixo numerico para geracao sequencial.
 *
 * @param string $ns
 * @return array|null
 */
function nsr_split_ns_sequence_parts($ns) {
    $ns = trim((string) $ns);
    if ($ns === '') {
        return null;
    }

    if (!preg_match('/^(.*?)(\d+)$/', $ns, $m)) {
        return null;
    }

    return array(
        'prefix' => (string) $m[1],
        'number' => (int) $m[2],
        'width' => strlen((string) $m[2]),
    );
}

/**
 * Monta NS sequencial com o mesmo padding do sufixo numerico.
 *
 * @param string $prefix
 * @param int    $number
 * @param int    $width
 * @return string
 */
function nsr_build_sequential_ns($prefix, $number, $width) {
    return (string) $prefix . str_pad((string) $number, (int) $width, '0', STR_PAD_LEFT);
}

/**
 * Informa se a integracao Tiny esta habilitada por constante.
 *
 * Configure no wp-config.php:
 * define('NSR_TINY_TOKEN', 'seu_token_tiny');
 *
 * @return bool
 */
function nsr_is_tiny_integration_enabled() {
    return defined('NSR_TINY_TOKEN') && trim((string) NSR_TINY_TOKEN) !== '';
}

/**
 * Retorna id do pedido alvo para envio ao Tiny.
 *
 * @param array $session
 * @return string
 */
function nsr_get_tiny_order_id_from_session($session) {
    $prefer = defined('NSR_TINY_ORDER_ID_SOURCE') ? strtolower(trim((string) NSR_TINY_ORDER_ID_SOURCE)) : 'pedido';
    $pedido = isset($session['pedido']) ? trim((string) $session['pedido']) : '';
    $nf = isset($session['nota_fiscal']) ? trim((string) $session['nota_fiscal']) : '';

    if ($prefer === 'nota_fiscal') {
        return $nf !== '' ? $nf : $pedido;
    }

    return $pedido !== '' ? $pedido : $nf;
}

/**
 * Coleta NS bipados por SKU da sessao atual.
 *
 * @param array $session
 * @return array
 */
function nsr_collect_serials_by_sku($session) {
    $result = array();
    $itens = isset($session['itens']) && is_array($session['itens']) ? $session['itens'] : array();

    foreach ($itens as $sku => $item) {
        if (!nsr_is_probable_sku((string) $sku)) {
            continue;
        }

        $scanned = isset($item['scanned']) && is_array($item['scanned']) ? $item['scanned'] : array();
        if (empty($scanned)) {
            continue;
        }

        $seen = array();
        $values = array();
        foreach ($scanned as $ns) {
            $ns = trim((string) $ns);
            if ($ns === '') {
                continue;
            }

            $norm = nsr_normalize_lookup_value($ns);
            if ($norm === '' || isset($seen[$norm])) {
                continue;
            }

            $seen[$norm] = true;
            $values[] = $ns;
        }

        if (!empty($values)) {
            $result[$sku] = $values;
        }
    }

    return $result;
}

/**
 * Monta texto de observacao para o Tiny com os NS bipados.
 *
 * @param array $session
 * @param array $serials_by_sku
 * @return string
 */
function nsr_build_tiny_obs_text($session, $serials_by_sku) {
    $lines = array();
    $pedido = isset($session['pedido']) ? trim((string) $session['pedido']) : '';
    $nf = isset($session['nota_fiscal']) ? trim((string) $session['nota_fiscal']) : '';

    $lines[] = 'NS bipados via NS Rastreio';
    $lines[] = 'Pedido: ' . ($pedido !== '' ? $pedido : '-');
    $lines[] = 'NF: ' . ($nf !== '' ? $nf : '-');

    foreach ($serials_by_sku as $sku => $serials) {
        $lines[] = $sku . ': ' . implode(', ', $serials);
    }

    $text = implode("\n", $lines);
    $max_len = defined('NSR_TINY_OBS_MAX_LEN') ? max(300, (int) NSR_TINY_OBS_MAX_LEN) : 1800;

    if (strlen($text) > $max_len) {
        $text = substr($text, 0, $max_len - 15) . '...(truncado)';
    }

    return $text;
}

/**
 * Envia observacao com NS bipados para o Tiny (pedido.alterar API 2.0).
 *
 * @param array $session
 * @return array|WP_Error
 */
function nsr_send_serials_to_tiny_order($session) {
    if (!nsr_is_tiny_integration_enabled()) {
        return array(
            'skipped' => true,
            'message' => 'Integracao Tiny nao configurada.',
        );
    }

    $order_id = nsr_get_tiny_order_id_from_session($session);
    if ($order_id === '') {
        return new WP_Error('nsr_tiny_missing_order', 'Pedido/NF nao informado para envio ao Tiny.');
    }

    $serials_by_sku = nsr_collect_serials_by_sku($session);
    if (empty($serials_by_sku)) {
        return new WP_Error('nsr_tiny_empty_serials', 'Nao ha NS bipados para enviar ao Tiny.');
    }

    $obs_text = nsr_build_tiny_obs_text($session, $serials_by_sku);
    $dados_pedido = array(
        'obs_interna' => $obs_text,
    );

    if (defined('NSR_TINY_WRITE_OBS_PUBLIC') && NSR_TINY_WRITE_OBS_PUBLIC) {
        $dados_pedido['obs'] = $obs_text;
    }

    $payload = array(
        'dados_pedido' => $dados_pedido,
    );

    $endpoint = defined('NSR_TINY_PEDIDO_ALTERAR_URL') && trim((string) NSR_TINY_PEDIDO_ALTERAR_URL) !== ''
        ? (string) NSR_TINY_PEDIDO_ALTERAR_URL
        : 'https://api.tiny.com.br/api2/pedido.alterar.php';

    $response = wp_remote_post($endpoint, array(
        'timeout' => 25,
        'body' => array(
            'token' => (string) NSR_TINY_TOKEN,
            'id' => $order_id,
            'dados_pedido' => wp_json_encode($payload),
            'formato' => 'json',
        ),
    ));

    if (is_wp_error($response)) {
        return new WP_Error('nsr_tiny_http_error', 'Erro de comunicacao com Tiny: ' . $response->get_error_message());
    }

    $body = wp_remote_retrieve_body($response);
    $data = json_decode((string) $body, true);
    if (!is_array($data) || !isset($data['retorno'])) {
        return new WP_Error('nsr_tiny_invalid_response', 'Retorno inesperado do Tiny ao atualizar pedido.');
    }

    $retorno = $data['retorno'];
    $status = isset($retorno['status']) ? strtoupper(trim((string) $retorno['status'])) : '';
    if ($status !== 'OK') {
        $erro_msg = 'Falha ao atualizar observacao no Tiny.';
        if (isset($retorno['erros'][0]['erro']) && trim((string) $retorno['erros'][0]['erro']) !== '') {
            $erro_msg = trim((string) $retorno['erros'][0]['erro']);
        }
        return new WP_Error('nsr_tiny_api_error', $erro_msg);
    }

    return array(
        'ok' => true,
        'order_id' => $order_id,
        'serials_count' => array_sum(array_map('count', $serials_by_sku)),
    );
}

/**
 * AJAX: Gerar NS sequenciais a partir de um NS inicial para o SKU ativo.
 */
function nsr_ajax_generate_ns_sequence() {
    if (!current_user_can('manage_options')) {
        wp_send_json_error(array('msg' => 'Permissao insuficiente.'), 403);
    }

    check_ajax_referer('nsr_ajax_scan', 'nonce');

    $token = sanitize_text_field(wp_unslash(isset($_POST['token']) ? $_POST['token'] : ''));
    $sku = strtoupper(sanitize_text_field(wp_unslash(isset($_POST['sku']) ? $_POST['sku'] : '')));
    $ns_start = sanitize_text_field(wp_unslash(isset($_POST['ns_start']) ? $_POST['ns_start'] : ''));
    $qty_requested = absint(wp_unslash(isset($_POST['qty']) ? $_POST['qty'] : 0));

    $session = nsr_get_scan_session($token);
    if (empty($session)) {
        wp_send_json_error(array('msg' => 'Sessao nao encontrada ou expirada.'), 404);
    }

    if ($sku === '' || $ns_start === '' || $qty_requested <= 0) {
        wp_send_json_error(array('msg' => 'Informe SKU, NS inicial e quantidade valida.'), 400);
    }

    if (!isset($session['itens'][$sku])) {
        wp_send_json_error(array('msg' => 'SKU ' . $sku . ' nao faz parte do pedido.'), 400);
    }

    $sequence = nsr_split_ns_sequence_parts($ns_start);
    if (empty($sequence)) {
        wp_send_json_error(array('msg' => 'NS inicial deve terminar com numero para gerar sequencia.'), 400);
    }

    $expected = (int) $session['itens'][$sku]['quantidade'];
    $already = isset($session['itens'][$sku]['scanned']) ? count($session['itens'][$sku]['scanned']) : 0;
    $available = max(0, $expected - $already);
    if ($available <= 0) {
        wp_send_json_error(array('msg' => sprintf('SKU %s ja atingiu a quantidade esperada (%d).', $sku, $expected)), 409);
    }

    $qty_target = min($qty_requested, $available);
    if (!isset($session['itens'][$sku]['scanned']) || !is_array($session['itens'][$sku]['scanned'])) {
        $session['itens'][$sku]['scanned'] = array();
    }

    $existing_normalized = array();
    foreach ($session['itens'] as $item) {
        $scanned = isset($item['scanned']) && is_array($item['scanned']) ? $item['scanned'] : array();
        foreach ($scanned as $s) {
            $norm = nsr_normalize_lookup_value($s);
            if ($norm !== '') {
                $existing_normalized[$norm] = true;
            }
        }
    }

    $generated = 0;
    $skipped_duplicates = 0;
    $preview_added = array();
    $current_number = (int) $sequence['number'];
    $max_attempts = max($qty_target * 5, 200);
    $attempts = 0;

    while ($generated < $qty_target && $attempts < $max_attempts) {
        $attempts++;
        $candidate = nsr_build_sequential_ns($sequence['prefix'], $current_number, $sequence['width']);
        $current_number++;

        $norm_candidate = nsr_normalize_lookup_value($candidate);
        if ($norm_candidate === '') {
            continue;
        }

        if (isset($existing_normalized[$norm_candidate])) {
            $skipped_duplicates++;
            continue;
        }

        $session['itens'][$sku]['scanned'][] = $candidate;
        $existing_normalized[$norm_candidate] = true;
        $generated++;

        if (count($preview_added) < 5) {
            $preview_added[] = $candidate;
        }
    }

    if ($generated <= 0) {
        wp_send_json_error(array('msg' => 'Nenhum NS novo foi gerado (duplicidade ou limite atingido).'), 409);
    }

    // Persiste pedido/NF se vieram junto.
    $pedido = sanitize_text_field(wp_unslash(isset($_POST['pedido']) ? $_POST['pedido'] : ''));
    $nf = sanitize_text_field(wp_unslash(isset($_POST['nota_fiscal']) ? $_POST['nota_fiscal'] : ''));
    if ($pedido !== '') {
        $session['pedido'] = $pedido;
    }
    if ($nf !== '') {
        $session['nota_fiscal'] = $nf;
    }

    nsr_save_scan_session($token, $session);

    $scanned_list = $session['itens'][$sku]['scanned'];
    $scanned_count = count($scanned_list);

    $msg = sprintf('Gerados %d NS sequenciais para %s.', $generated, $sku);
    if ($qty_requested > $qty_target) {
        $msg .= sprintf(' Ajustado para o saldo disponivel (%d).', $qty_target);
    }
    if ($skipped_duplicates > 0) {
        $msg .= sprintf(' Duplicados ignorados: %d.', $skipped_duplicates);
    }

    wp_send_json_success(array(
        'sku' => $sku,
        'scanned_count' => $scanned_count,
        'expected' => $expected,
        'is_ok' => ($scanned_count === $expected),
        'scanned_list' => $scanned_list,
        'generated' => $generated,
        'preview_added' => $preview_added,
        'msg' => $msg,
    ));
}
add_action('wp_ajax_nsr_generate_ns_sequence', 'nsr_ajax_generate_ns_sequence');

// AJAX: Remover NS
function nsr_ajax_remove_ns() {
    if (!current_user_can('manage_options')) {
        wp_send_json_error(array('msg' => 'Permissao insuficiente.'), 403);
    }

    check_ajax_referer('nsr_ajax_scan', 'nonce');

    $token = sanitize_text_field(wp_unslash(isset($_POST['token']) ? $_POST['token'] : ''));
    $sku   = strtoupper(sanitize_text_field(wp_unslash(isset($_POST['sku']) ? $_POST['sku'] : '')));
    $ns    = sanitize_text_field(wp_unslash(isset($_POST['ns']) ? $_POST['ns'] : ''));

    $session = nsr_get_scan_session($token);
    if (empty($session)) {
        wp_send_json_error(array('msg' => 'Sessao nao encontrada.'), 404);
    }

    $removed = false;
    if (isset($session['itens'][$sku]['scanned']) && is_array($session['itens'][$sku]['scanned'])) {
        $idx = array_search($ns, $session['itens'][$sku]['scanned'], true);
        if ($idx !== false) {
            unset($session['itens'][$sku]['scanned'][$idx]);
            $session['itens'][$sku]['scanned'] = array_values($session['itens'][$sku]['scanned']);
            $removed = true;
        }
    }

    if (!$removed) {
        wp_send_json_error(array('msg' => 'NS nao encontrado.'), 404);
    }

    nsr_save_scan_session($token, $session);

    $scanned_count = count($session['itens'][$sku]['scanned']);
    $expected      = (int) $session['itens'][$sku]['quantidade'];

    wp_send_json_success(array(
        'sku'           => $sku,
        'ns'            => $ns,
        'scanned_count' => $scanned_count,
        'expected'      => $expected,
        'is_ok'         => ($scanned_count === $expected),
        'scanned_list'  => $session['itens'][$sku]['scanned'],
    ));
}
add_action('wp_ajax_nsr_remove_ns', 'nsr_ajax_remove_ns');

// AJAX: Finalizar sessao
function nsr_ajax_finalize_session() {
    if (!current_user_can('manage_options')) {
        wp_send_json_error(array('msg' => 'Permissao insuficiente.'), 403);
    }

    check_ajax_referer('nsr_ajax_scan', 'nonce');

    $token = sanitize_text_field(wp_unslash(isset($_POST['token']) ? $_POST['token'] : ''));
    $session = nsr_get_scan_session($token);
    if (empty($session)) {
        wp_send_json_error(array('msg' => 'Sessao nao encontrada.'), 404);
    }

    $pedido = sanitize_text_field(wp_unslash(isset($_POST['pedido']) ? $_POST['pedido'] : ''));
    $nf     = sanitize_text_field(wp_unslash(isset($_POST['nota_fiscal']) ? $_POST['nota_fiscal'] : ''));
    if ($pedido !== '') {
        $session['pedido'] = $pedido;
    }
    if ($nf !== '') {
        $session['nota_fiscal'] = $nf;
    }

    if (empty($session['pedido']) && empty($session['nota_fiscal'])) {
        wp_send_json_error(array('msg' => 'Informe ao menos Pedido ou Nota Fiscal para finalizar.'), 400);
    }

    $session = nsr_recompute_scan_session_flags($session);
    $missing_skus = !empty($session['missing_skus']) && is_array($session['missing_skus']) ? $session['missing_skus'] : array();

    foreach ($session['itens'] as $sku => $item) {
        $expected = (int) $item['quantidade'];
        $scanned  = isset($item['scanned']) && is_array($item['scanned']) ? count($item['scanned']) : 0;
        if ($expected !== $scanned) {
            wp_send_json_error(array('msg' => sprintf('SKU %s: esperado %d, bipado %d.', $sku, $expected, $scanned)), 400);
        }
    }

    global $wpdb;
    $wpdb->query('START TRANSACTION');
    $saved  = 0;
    $failed = false;

    foreach ($session['itens'] as $sku => $item) {
        $descricao    = isset($item['descricao']) ? (string) $item['descricao'] : '';
        $expected     = (int) $item['quantidade'];
        $valor_item   = isset($item['valor']) ? (string) $item['valor'] : '';
        $scanned_list = isset($item['scanned']) && is_array($item['scanned']) ? $item['scanned'] : array();
        $origem       = isset($session['origem_arquivo']) ? (string) $session['origem_arquivo'] : '';
        $data_venda   = (isset($session['data_venda']) && (string) $session['data_venda'] !== '') ? (string) $session['data_venda'] : date_i18n('Y-m-d');

        foreach ($scanned_list as $ns) {
            $ns_normalizado = nsr_normalize_lookup_value($ns);
            if ($ns_normalizado === '') {
                $failed = true;
                break 2;
            }

            $affected = nsr_upsert_ns_record(array(
                'ns'            => $ns,
                'ns_normalizado' => $ns_normalizado,
                'nota_fiscal'   => (string) $session['nota_fiscal'],
                'pedido'        => (string) $session['pedido'],
                'sku'           => $sku,
                'descricao'     => $descricao,
                'quantidade'    => (string) $expected,
                'valor'         => $valor_item,
                'data_venda'    => $data_venda,
                'origem_arquivo'=> $origem,
                'linha_origem'  => 0,
            ));

            if ($affected === false) {
                $failed = true;
                break 2;
            }
            $saved++;
        }
    }

    if ($failed) {
        $wpdb->query('ROLLBACK');
        wp_send_json_error(array('msg' => 'Erro ao salvar registros. Nenhum dado foi gravado.'), 500);
    }

    $wpdb->query('COMMIT');

    $tiny_result = nsr_send_serials_to_tiny_order($session);

    nsr_delete_scan_session($token);

    $msg = sprintf('%d NS salvos com sucesso. Pedido: %s | NF: %s', $saved, $session['pedido'], $session['nota_fiscal']);
    if (!empty($missing_skus)) {
        $msg .= ' | Aviso: SKU(s) sem cadastro: ' . implode(', ', $missing_skus);
    }

    if (is_wp_error($tiny_result)) {
        $msg .= ' | Aviso Tiny: ' . $tiny_result->get_error_message();
    } elseif (is_array($tiny_result) && isset($tiny_result['ok']) && $tiny_result['ok']) {
        $msg .= sprintf(
            ' | Tiny atualizado no pedido %s (%d NS).',
            isset($tiny_result['order_id']) ? (string) $tiny_result['order_id'] : '-',
            isset($tiny_result['serials_count']) ? (int) $tiny_result['serials_count'] : 0
        );
    }

    wp_send_json_success(array(
        'msg'   => $msg,
        'saved' => $saved,
        'missing_skus' => $missing_skus,
    ));
}
add_action('wp_ajax_nsr_finalize_session', 'nsr_ajax_finalize_session');

// AJAX: Cancelar sessao
function nsr_ajax_cancel_session() {
    if (!current_user_can('manage_options')) {
        wp_send_json_error(array('msg' => 'Permissao insuficiente.'), 403);
    }
    check_ajax_referer('nsr_ajax_scan', 'nonce');
    $token = sanitize_text_field(wp_unslash(isset($_POST['token']) ? $_POST['token'] : ''));
    nsr_delete_scan_session($token);
    wp_send_json_success(array('msg' => 'Sessao cancelada.'));
}
add_action('wp_ajax_nsr_cancel_session', 'nsr_ajax_cancel_session');

function nsr_render_admin_page() {
    global $wpdb;

    $messages = array(
        'success' => array(),
        'error' => array(),
    );
    $messages = nsr_merge_messages($messages, nsr_handle_import_submission());
    $messages = nsr_merge_messages($messages, nsr_handle_products_import_submission());
    $messages = nsr_merge_messages($messages, nsr_handle_product_manual_submission());
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
    $admin_search_type = isset($_GET['nsr_admin_tipo']) ? sanitize_key(wp_unslash($_GET['nsr_admin_tipo'])) : 'ns';
    if (!in_array($admin_search_type, array('ns', 'nf', 'pedido'), true)) {
        $admin_search_type = 'ns';
    }
    $admin_search_label_map = array(
        'ns' => 'NS',
        'nf' => 'NF',
        'pedido' => 'Pedido',
    );
    $admin_search_label = $admin_search_label_map[$admin_search_type];
    $admin_is_partial = (isset($_GET['nsr_admin_partial']) && $_GET['nsr_admin_partial'] === '1');
    $admin_results = array();
    if ($admin_search_value !== '') {
        $admin_results = nsr_find_admin_records($admin_search_value, $admin_search_type, $admin_is_partial, 200);
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

        <h3>2.1) Cadastro manual rapido de produto</h3>
        <p>Use este formulario quando chegar SKU novo no pedido (ex.: PRD000XX) e ele ainda nao estiver na base.</p>
        <form method="post" style="margin-bottom:24px;display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap;">
            <?php wp_nonce_field('nsr_product_manual', 'nsr_product_manual_nonce'); ?>
            <label style="display:flex;flex-direction:column;gap:4px;">
                SKU
                <input type="text" name="nsr_product_sku" placeholder="Ex: PRD00016" required style="min-width:180px;" />
            </label>
            <label style="display:flex;flex-direction:column;gap:4px;flex:1;min-width:260px;">
                Descricao
                <input type="text" name="nsr_product_descricao" placeholder="Descricao do produto" />
            </label>
            <button type="submit" name="nsr_product_manual_submit" class="button button-primary">Salvar produto</button>
        </form>

        <h2>3) Leitura de Pedido de Venda (PDF) e Bipagem de NS</h2>
        <p>Envie o PDF do pedido para extrair SKU e quantidade. Depois, realize a bipagem dos NS por SKU.</p>
        <form method="post" enctype="multipart/form-data" style="margin-bottom:16px;">
            <?php wp_nonce_field('nsr_pdf_upload', 'nsr_pdf_nonce'); ?>
            <input type="file" name="nsr_pdf_file" accept=".pdf,.xml" required />
            <button type="submit" name="nsr_pdf_upload_submit" class="button">Ler PDF / XML</button>
        </form>
        <p style="margin-top:-10px;color:#666;font-size:12px;">Aceita PDF do pedido de venda ou XML da NF-e (NF-e 4.0).</p>

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
            $scan_token   = isset($scan_session['session_token']) ? (string) $scan_session['session_token'] : '';
            $missing_skus = isset($scan_session['missing_skus']) && is_array($scan_session['missing_skus']) ? $scan_session['missing_skus'] : array();
            $ajax_nonce   = wp_create_nonce('nsr_ajax_scan');
            ?>
            <div id="nsr-scan-panel" style="background:#fff;border:1px solid #dcdcde;border-radius:6px;padding:16px;margin-bottom:16px;">
                <h3 style="margin:0 0 10px 0;">
                    Previa do pedido &mdash;
                    Arquivo: <span style="font-weight:normal;"><?php echo esc_html(isset($scan_session['origem_arquivo']) ? $scan_session['origem_arquivo'] : ''); ?></span>
                    | Pedido: <span style="font-weight:normal;" id="nsr-hdr-pedido"><?php echo esc_html(isset($scan_session['pedido']) ? $scan_session['pedido'] : ''); ?></span>
                    | NF: <span style="font-weight:normal;" id="nsr-hdr-nf"><?php echo esc_html(isset($scan_session['nota_fiscal']) ? $scan_session['nota_fiscal'] : ''); ?></span>
                </h3>

                <?php if (!empty($missing_skus)) : ?>
                    <div style="padding:8px 12px;background:#fff3cd;border:1px solid #ffc107;border-radius:4px;color:#856404;margin-bottom:10px;">
                        SKU(s) sem cadastro: <strong><?php echo esc_html(implode(', ', $missing_skus)); ?></strong>
                    </div>
                <?php endif; ?>

                <!-- Tabela de SKUs clicavel -->
                <table class="widefat" id="nsr-sku-table" style="margin-bottom:14px;cursor:pointer;">
                    <thead>
                        <tr>
                            <th></th>
                            <th>SKU</th>
                            <th>Descricao</th>
                            <th>Qtd Pedido</th>
                            <th>Qtd Bipado</th>
                            <th>Status</th>
                            <th>NS bipados</th>
                        </tr>
                    </thead>
                    <tbody>
                        <?php foreach ($scan_session['itens'] as $sku => $item) :
                            $expected      = (int) $item['quantidade'];
                            $scanned_list  = isset($item['scanned']) && is_array($item['scanned']) ? $item['scanned'] : array();
                            $scanned_count = count($scanned_list);
                            $is_ok         = ($expected === $scanned_count);
                            static $row_idx = 0;
                            $row_idx++;
                            $shortcut_key = $row_idx <= 9 ? $row_idx : null;
                            ?>
                            <tr data-sku="<?php echo esc_attr($sku); ?>"
                                data-expected="<?php echo esc_attr((string) $expected); ?>"
                                onclick="nsrSelectSku(this)"
                                style="transition:background .15s;">
                                <td style="text-align:center;font-weight:600;width:30px;">
                                    <?php if ($shortcut_key) : ?>
                                        <span style="background:#2271b1;color:#fff;border-radius:3px;padding:2px 6px;font-size:12px;display:inline-block;"><?php echo esc_html((string) $shortcut_key); ?></span>
                                    <?php endif; ?>
                                </td>
                                <td><strong><?php echo esc_html($sku); ?></strong></td>
                                <td><?php echo esc_html(isset($item['descricao']) ? $item['descricao'] : ''); ?></td>
                                <td><?php echo esc_html((string) $expected); ?></td>
                                <td class="nsr-count"><?php echo esc_html((string) $scanned_count); ?></td>
                                <td class="nsr-status">
                                    <?php if ($is_ok) : ?>
                                        <span style="color:#0a7d28;font-weight:600;">&#10003; OK</span>
                                    <?php else : ?>
                                        <span style="color:#b32d2e;font-weight:600;">Pendente</span>
                                    <?php endif; ?>
                                </td>
                                <td class="nsr-ns-list">
                                    <?php foreach ($scanned_list as $sn) : ?>
                                        <button type="button" class="button button-small nsr-remove-ns"
                                            data-sku="<?php echo esc_attr($sku); ?>"
                                            data-ns="<?php echo esc_attr($sn); ?>"
                                            style="margin:1px 2px;"
                                            onclick="event.stopPropagation();nsrRemoveNs(this)">
                                            <?php echo esc_html($sn); ?> &times;
                                        </button>
                                    <?php endforeach; ?>
                                </td>
                            </tr>
                        <?php endforeach; ?>
                    </tbody>
                </table>

                <!-- Painel de bipagem AJAX -->
                <div style="background:#f6f7f7;border:1px solid #dcdcde;border-radius:6px;padding:14px;margin-bottom:12px;">
                    <div style="display:flex;flex-wrap:wrap;gap:10px;align-items:flex-end;margin-bottom:10px;">
                        <label style="display:flex;flex-direction:column;gap:3px;font-size:13px;">
                            Pedido
                            <input type="text" id="nsr-inp-pedido" value="<?php echo esc_attr(isset($scan_session['pedido']) ? $scan_session['pedido'] : ''); ?>" style="width:120px;" />
                        </label>
                        <label style="display:flex;flex-direction:column;gap:3px;font-size:13px;">
                            Nota Fiscal
                            <input type="text" id="nsr-inp-nf" value="<?php echo esc_attr(isset($scan_session['nota_fiscal']) ? $scan_session['nota_fiscal'] : ''); ?>" style="width:120px;" />
                        </label>
                        <label style="display:flex;flex-direction:column;gap:3px;font-size:13px;">
                            SKU ativo
                            <input type="text" id="nsr-inp-sku" value="" readonly placeholder="Clique na tabela" style="width:140px;background:#fff;font-weight:600;" />
                        </label>
                        <label style="display:flex;flex-direction:column;gap:3px;font-size:13px;flex:1;min-width:200px;">
                            Numero de Serie (NS)
                            <input type="text" id="nsr-inp-ns" placeholder="Bipagem aqui..." autocomplete="off"
                                style="font-size:15px;padding:6px;border:2px solid #2271b1;"
                                onkeydown="if(event.key==='Enter'){event.preventDefault();nsrScanNs();}" />
                        </label>
                        <button type="button" class="button button-primary" onclick="nsrScanNs()" style="height:36px;align-self:flex-end;">
                            Bipar NS
                        </button>
                        <label style="display:flex;flex-direction:column;gap:3px;font-size:13px;min-width:150px;">
                            Qtd sequencial
                            <input type="number" id="nsr-inp-seq-qty" min="1" step="1" value="" placeholder="Ex: 80" style="width:120px;" />
                        </label>
                        <label style="display:flex;align-items:center;gap:6px;font-size:12px;padding-bottom:8px;">
                            <input type="checkbox" id="nsr-inp-seq-auto" />
                            Completar saldo restante automaticamente
                        </label>
                        <button type="button" class="button" onclick="nsrGenerateSequentialNs()" style="height:36px;align-self:flex-end;">
                            Gerar sequencial
                        </button>
                    </div>

                    <p style="margin:0 0 8px 0;color:#555;font-size:12px;">
                        Dica: informe o primeiro NS em <strong>Numero de Serie (NS)</strong>, defina a <strong>Qtd sequencial</strong> e clique em <strong>Gerar sequencial</strong>.
                    </p>

                    <div id="nsr-scan-feedback" style="min-height:28px;font-weight:600;padding:4px 8px;border-radius:4px;display:none;"></div>
                </div>

                <!-- Botoes finalizar/cancelar -->
                <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;">
                    <button type="button" class="button button-primary" id="nsr-btn-finalize" onclick="nsrFinalize()">
                        Finalizar e salvar NS
                    </button>
                    <button type="button" class="button" id="nsr-btn-cancel" onclick="nsrCancel()">
                        Cancelar sessao
                    </button>
                    <span id="nsr-finalize-msg" style="color:#b32d2e;"></span>
                </div>
            </div>

            <script>
            (function(){
                var TOKEN   = <?php echo wp_json_encode($scan_token); ?>;
                var NONCE   = <?php echo wp_json_encode($ajax_nonce); ?>;
                var AJAXURL = <?php echo wp_json_encode(admin_url('admin-ajax.php')); ?>;
                var activeSku = null;

                // ----- SKU selection -----
                window.nsrSelectSku = function(row) {
                    document.querySelectorAll('#nsr-sku-table tbody tr').forEach(function(r) {
                        r.style.background = '';
                        var badge = r.querySelector('td:first-child span');
                        if(badge) badge.style.opacity = '0.5';
                    });
                    row.style.background = '#e5f0fb';
                    var badge = row.querySelector('td:first-child span');
                    if(badge) badge.style.opacity = '1';
                    activeSku = row.dataset.sku;
                    document.getElementById('nsr-inp-sku').value = activeSku;
                    document.getElementById('nsr-inp-ns').focus();
                };

                // Auto-select first SKU that's not done
                (function autoSelect(){
                    var rows = document.querySelectorAll('#nsr-sku-table tbody tr');
                    for(var i=0;i<rows.length;i++){
                        var expected = parseInt(rows[i].dataset.expected,10);
                        var count    = parseInt(rows[i].querySelector('.nsr-count').textContent,10);
                        if(count < expected){ window.nsrSelectSku(rows[i]); return; }
                    }
                    if(rows.length) window.nsrSelectSku(rows[0]);
                })();

                // Atalhos de teclado: Num 1-9 para selecionar SKUs rapidamente
                document.addEventListener('keydown', function(e){
                    // Ignora se estiver digitando no input de NS
                    if(document.activeElement.id === 'nsr-inp-ns') return;
                    
                    var key = e.key;
                    if(key >= '1' && key <= '9'){
                        var idx = parseInt(key, 10) - 1;
                        var rows = document.querySelectorAll('#nsr-sku-table tbody tr');
                        if(idx < rows.length){
                            e.preventDefault();
                            window.nsrSelectSku(rows[idx]);
                        }
                    }
                });

                function showFeedback(msg, type) {
                    var el = document.getElementById('nsr-scan-feedback');
                    el.textContent = msg;
                    el.style.display = 'block';
                    el.style.background = type === 'ok' ? '#d1fae5' : (type === 'warn' ? '#fff3cd' : '#fee2e2');
                    el.style.color      = type === 'ok' ? '#065f46' : (type === 'warn' ? '#92400e' : '#7f1d1d');
                    el.style.border     = '1px solid ' + (type === 'ok' ? '#6ee7b7' : (type === 'warn' ? '#fcd34d' : '#fca5a5'));
                    clearTimeout(el._t);
                    if (type === 'ok') {
                        el._t = setTimeout(function(){ el.style.display='none'; }, 2500);
                    }
                }

                function updateRow(data) {
                    var row = document.querySelector('#nsr-sku-table tbody tr[data-sku="'+data.sku+'"]');
                    if (!row) return;
                    row.querySelector('.nsr-count').textContent   = data.scanned_count;
                    var statusEl = row.querySelector('.nsr-status');
                    statusEl.innerHTML = data.is_ok
                        ? '<span style="color:#0a7d28;font-weight:600;">&#10003; OK</span>'
                        : '<span style="color:#b32d2e;font-weight:600;">Pendente</span>';
                    // Rebuild NS list
                    var nsList = row.querySelector('.nsr-ns-list');
                    nsList.innerHTML = '';
                    data.scanned_list.forEach(function(sn){
                        var btn = document.createElement('button');
                        btn.type = 'button';
                        btn.className = 'button button-small nsr-remove-ns';
                        btn.dataset.sku = data.sku;
                        btn.dataset.ns  = sn;
                        btn.style.margin = '1px 2px';
                        btn.textContent  = sn + ' \u00d7';
                        btn.onclick = function(e){ e.stopPropagation(); window.nsrRemoveNs(btn); };
                        nsList.appendChild(btn);
                    });
                    // Auto advance to next pending SKU
                    if (data.is_ok && activeSku === data.sku) {
                        var rows = document.querySelectorAll('#nsr-sku-table tbody tr');
                        for(var i=0;i<rows.length;i++){
                            if(parseInt(rows[i].querySelector('.nsr-count').textContent,10) <
                               parseInt(rows[i].dataset.expected,10)){
                                window.nsrSelectSku(rows[i]);
                                return;
                            }
                        }
                    }
                }

                // ----- Bipar NS -----
                window.nsrScanNs = function() {
                    var ns = document.getElementById('nsr-inp-ns').value.trim();
                    if (!activeSku || !ns) {
                        showFeedback(activeSku ? 'Digite o NS.' : 'Selecione um SKU primeiro.', 'error');
                        return;
                    }
                    var pedido = document.getElementById('nsr-inp-pedido').value.trim();
                    var nf     = document.getElementById('nsr-inp-nf').value.trim();
                    var fd = new FormData();
                    fd.append('action',      'nsr_scan_ns');
                    fd.append('nonce',       NONCE);
                    fd.append('token',       TOKEN);
                    fd.append('sku',         activeSku);
                    fd.append('ns',          ns);
                    fd.append('pedido',      pedido);
                    fd.append('nota_fiscal', nf);
                    fetch(AJAXURL, {method:'POST', body:fd})
                        .then(function(r){ return r.json(); })
                        .then(function(res){
                            if (res.success) {
                                document.getElementById('nsr-inp-ns').value = '';
                                document.getElementById('nsr-inp-ns').focus();
                                updateRow(res.data);
                                showFeedback('NS ' + res.data.ns + ' bipado (' + res.data.scanned_count + '/' + res.data.expected + ')', 'ok');
                                document.getElementById('nsr-hdr-pedido').textContent = pedido;
                                document.getElementById('nsr-hdr-nf').textContent     = nf;
                            } else {
                                showFeedback(res.data.msg, 'error');
                                document.getElementById('nsr-inp-ns').select();
                            }
                        })
                        .catch(function(){ showFeedback('Erro de comunicacao com o servidor.', 'error'); });
                };

                // ----- Gerar sequencia de NS -----
                window.nsrGenerateSequentialNs = function() {
                    var nsStart = document.getElementById('nsr-inp-ns').value.trim();
                    var autoQty = document.getElementById('nsr-inp-seq-auto').checked;
                    var qtyRaw  = document.getElementById('nsr-inp-seq-qty').value;
                    var qty     = parseInt(qtyRaw, 10);

                    if (!activeSku) {
                        showFeedback('Selecione um SKU primeiro.', 'error');
                        return;
                    }
                    if (!nsStart) {
                        showFeedback('Informe o NS inicial no campo de bipagem.', 'error');
                        document.getElementById('nsr-inp-ns').focus();
                        return;
                    }

                    if (autoQty) {
                        var activeRow = document.querySelector('#nsr-sku-table tbody tr[data-sku="' + activeSku + '"]');
                        if (activeRow) {
                            var expected = parseInt(activeRow.dataset.expected, 10) || 0;
                            var scanned  = parseInt(activeRow.querySelector('.nsr-count').textContent, 10) || 0;
                            qty = Math.max(0, expected - scanned);
                            document.getElementById('nsr-inp-seq-qty').value = qty > 0 ? String(qty) : '';
                        }
                    }

                    if (!qty || qty <= 0) {
                        showFeedback(autoQty ? 'Este SKU nao possui saldo restante para gerar.' : 'Informe uma quantidade sequencial valida.', 'error');
                        if (!autoQty) {
                            document.getElementById('nsr-inp-seq-qty').focus();
                        }
                        return;
                    }

                    var pedido = document.getElementById('nsr-inp-pedido').value.trim();
                    var nf     = document.getElementById('nsr-inp-nf').value.trim();
                    var fd = new FormData();
                    fd.append('action',      'nsr_generate_ns_sequence');
                    fd.append('nonce',       NONCE);
                    fd.append('token',       TOKEN);
                    fd.append('sku',         activeSku);
                    fd.append('ns_start',    nsStart);
                    fd.append('qty',         String(qty));
                    fd.append('pedido',      pedido);
                    fd.append('nota_fiscal', nf);

                    fetch(AJAXURL, {method:'POST', body:fd})
                        .then(function(r){ return r.json(); })
                        .then(function(res){
                            if (res.success) {
                                updateRow(res.data);
                                showFeedback(res.data.msg, 'ok');
                                document.getElementById('nsr-hdr-pedido').textContent = pedido;
                                document.getElementById('nsr-hdr-nf').textContent     = nf;
                                document.getElementById('nsr-inp-ns').focus();
                                document.getElementById('nsr-inp-ns').select();
                            } else {
                                showFeedback(res.data.msg, 'error');
                            }
                        })
                        .catch(function(){ showFeedback('Erro de comunicacao com o servidor.', 'error'); });
                };

                // ----- Bipar múltiplos NSs (cola Ctrl+V) -----
                window.nsrScanMultipleNs = function(nsList) {
                    if (!activeSku) {
                        showFeedback('Selecione um SKU primeiro.', 'error');
                        return;
                    }
                    if (!Array.isArray(nsList) || nsList.length === 0) {
                        showFeedback('Nenhum NS para bipar.', 'error');
                        return;
                    }

                    var pedido = document.getElementById('nsr-inp-pedido').value.trim();
                    var nf     = document.getElementById('nsr-inp-nf').value.trim();
                    var processed = 0;
                    var successful = 0;
                    var failed = 0;

                    function processNext() {
                        if (processed >= nsList.length) {
                            showFeedback(successful + ' NS bipados, ' + failed + ' erros.', failed === 0 ? 'ok' : 'warn');
                            document.getElementById('nsr-inp-ns').focus();
                            return;
                        }

                        var ns = nsList[processed++].trim();
                        if (!ns) {
                            setTimeout(processNext, 100);
                            return;
                        }

                        var fd = new FormData();
                        fd.append('action',      'nsr_scan_ns');
                        fd.append('nonce',       NONCE);
                        fd.append('token',       TOKEN);
                        fd.append('sku',         activeSku);
                        fd.append('ns',          ns);
                        fd.append('pedido',      pedido);
                        fd.append('nota_fiscal', nf);

                        fetch(AJAXURL, {method:'POST', body:fd})
                            .then(function(r){ return r.json(); })
                            .then(function(res){
                                if (res.success) {
                                    updateRow(res.data);
                                    successful++;
                                    document.getElementById('nsr-hdr-pedido').textContent = pedido;
                                    document.getElementById('nsr-hdr-nf').textContent     = nf;
                                } else {
                                    failed++;
                                }
                                setTimeout(processNext, 200);
                            })
                            .catch(function(){ failed++; setTimeout(processNext, 200); });
                    }

                    document.getElementById('nsr-inp-ns').value = '';
                    showFeedback('Bipando ' + nsList.length + ' NS(s)...', 'warn');
                    processNext();
                };

                // Event listener: paste Ctrl+V no campo de NS
                document.getElementById('nsr-inp-ns').addEventListener('paste', function(e){
                    e.preventDefault();
                    var text = (e.clipboardData || window.clipboardData).getData('text');
                    
                    // Split por quebra de linha, espaço, vírgula, ponto-e-vírgula, pipe
                    var nsList = text
                        .split(/[\r\n,;|\s]+/)
                        .map(function(s){ return s.trim(); })
                        .filter(function(s){ return s.length > 0; });

                    if (nsList.length === 1) {
                        // Se só tem um NS, comporta-se normalmente
                        document.getElementById('nsr-inp-ns').value = nsList[0];
                    } else if (nsList.length > 1) {
                        // Múltiplos NSs: bipar tudo
                        window.nsrScanMultipleNs(nsList);
                    }
                });


                // ----- Remover NS -----
                window.nsrRemoveNs = function(btn) {
                    if (!confirm('Remover NS ' + btn.dataset.ns + '?')) return;
                    var fd = new FormData();
                    fd.append('action', 'nsr_remove_ns');
                    fd.append('nonce',  NONCE);
                    fd.append('token',  TOKEN);
                    fd.append('sku',    btn.dataset.sku);
                    fd.append('ns',     btn.dataset.ns);
                    fetch(AJAXURL, {method:'POST', body:fd})
                        .then(function(r){ return r.json(); })
                        .then(function(res){
                            if (res.success) {
                                updateRow(res.data);
                                showFeedback('NS ' + res.data.ns + ' removido.', 'warn');
                            } else {
                                showFeedback(res.data.msg, 'error');
                            }
                        });
                };

                // ----- Finalizar -----
                window.nsrFinalize = function() {
                    var pedido = document.getElementById('nsr-inp-pedido').value.trim();
                    var nf     = document.getElementById('nsr-inp-nf').value.trim();
                    if (!pedido && !nf) {
                        showFeedback('Informe Pedido ou Nota Fiscal.', 'error');
                        return;
                    }
                    if (!confirm('Finalizar e salvar todos os NS?')) return;
                    var fd = new FormData();
                    fd.append('action',      'nsr_finalize_session');
                    fd.append('nonce',       NONCE);
                    fd.append('token',       TOKEN);
                    fd.append('pedido',      pedido);
                    fd.append('nota_fiscal', nf);
                    fetch(AJAXURL, {method:'POST', body:fd})
                        .then(function(r){ return r.json(); })
                        .then(function(res){
                            var panel = document.getElementById('nsr-scan-panel');
                            if (res.success) {
                                panel.innerHTML = '<div style="padding:16px;background:#d1fae5;border:1px solid #6ee7b7;border-radius:6px;color:#065f46;font-weight:600;font-size:15px;">' +
                                    '&#10003; ' + res.data.msg + '</div>';
                            } else {
                                showFeedback(res.data.msg, 'error');
                            }
                        });
                };

                // ----- Cancelar -----
                window.nsrCancel = function() {
                    if (!confirm('Cancelar a sessao? Os NS bipados serao descartados.')) return;
                    var fd = new FormData();
                    fd.append('action', 'nsr_cancel_session');
                    fd.append('nonce',  NONCE);
                    fd.append('token',  TOKEN);
                    fetch(AJAXURL, {method:'POST', body:fd})
                        .then(function(r){ return r.json(); })
                        .then(function(){ document.getElementById('nsr-scan-panel').remove(); });
                };
            })();
            </script>
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
            <select name="nsr_admin_tipo">
                <option value="ns" <?php selected($admin_search_type, 'ns'); ?>>NS</option>
                <option value="nf" <?php selected($admin_search_type, 'nf'); ?>>NF</option>
                <option value="pedido" <?php selected($admin_search_type, 'pedido'); ?>>Pedido</option>
            </select>
            <input type="text" name="nsr_admin_ns" value="<?php echo esc_attr($admin_search_value); ?>" placeholder="Digite NS, NF ou Pedido" style="flex:1;" />
            <label style="display:flex;align-items:center;gap:6px;">
                <input type="checkbox" name="nsr_admin_partial" value="1" <?php checked($admin_is_partial); ?> />
                Busca parcial
            </label>
            <button type="submit" class="button">Buscar</button>
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
                        <strong>⚠ <?php echo esc_html($admin_search_label); ?> não encontrado</strong><br/>
                        Nenhum resultado encontrado para <?php echo esc_html($admin_search_label); ?>: <strong><?php echo esc_html($admin_search_value); ?></strong>
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
