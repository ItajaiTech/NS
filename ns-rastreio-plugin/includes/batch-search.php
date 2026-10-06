<?php
defined('ABSPATH') || exit;

/** Parse pasted serials without dropping short or alphabetic identifiers. */
function nsr_parse_batch_serials($text) {
    $serials = array();
    foreach (preg_split('/[\s,;|]+/u', trim((string) $text)) as $value) {
        $key = nsr_normalize_lookup_value($value);
        if ($key !== '' && !isset($serials[$key])) {
            $serials[$key] = $value;
        }
    }
    return $serials;
}

/** Fetch all invoices for each serial, preserving the pasted order. */
function nsr_find_batch_serials($serials) {
    global $wpdb;
    $groups = array_fill_keys(array_keys($serials), array());
    foreach (array_chunk(array_keys($serials), 100) as $chunk) {
        $placeholders = implode(',', array_fill(0, count($chunk), '%s'));
        $sql = $wpdb->prepare(
            'SELECT ns_normalizado, ns, nota_fiscal, pedido, sku, descricao, data_venda FROM ' . nsr_get_table_name() .
            " WHERE ns_normalizado IN ($placeholders) ORDER BY updated_at DESC, id DESC",
            $chunk
        );
        $rows = $wpdb->get_results($sql, ARRAY_A);
        if ($rows === null || $wpdb->last_error !== '') {
            return new WP_Error('nsr_batch_query', 'Nao foi possivel consultar os NS. Tente novamente.');
        }
        foreach (nsr_enrich_product_records($rows) as $row) {
            $key = nsr_normalize_lookup_value($row['ns_normalizado']);
            if (isset($groups[$key])) {
                $groups[$key][] = $row;
            }
        }
    }
    return $groups;
}

/** Strict dates; month-end anniversaries stay in the target month. */
function nsr_format_sale_date($value) {
    foreach (array('!d/m/Y', '!Y-m-d', '!d-m-Y', '!Y-m-d H:i:s', '!d/m/Y H:i:s') as $format) {
        $date = DateTimeImmutable::createFromFormat($format, trim((string) $value));
        $errors = DateTimeImmutable::getLastErrors();
        if ($date && (!$errors || (!$errors['warning_count'] && !$errors['error_count']))) {
            return $date->format('d/m/Y');
        }
    }
    return (string) $value;
}

function nsr_batch_expiration($value) {
    $date = false;
    foreach (array('!d/m/Y', '!Y-m-d', '!d-m-Y', '!Y-m-d H:i:s', '!d/m/Y H:i:s') as $format) {
        $candidate = DateTimeImmutable::createFromFormat($format, trim((string) $value));
        $errors = DateTimeImmutable::getLastErrors();
        if ($candidate && (!$errors || (!$errors['warning_count'] && !$errors['error_count']))) {
            $date = $candidate;
            break;
        }
    }
    if (!$date) {
        return 'Data de venda ausente ou invalida';
    }
    $target = $date->modify('first day of this month')->modify('+2 years');
    $day = min((int) $date->format('d'), (int) $target->format('t'));
    return $target->setDate((int) $target->format('Y'), (int) $target->format('m'), $day)->format('d/m/Y');
}

/** Shared admin and shortcode batch consultation. Only reads existing records. */
function nsr_render_batch_search($context) {
    $submitted = isset($_POST['nsr_batch_context']) && $_POST['nsr_batch_context'] === $context;
    $text = '';
    $error = '';
    $groups = null;
    $serials = array();
    if ($submitted) {
        $nonce = isset($_POST['nsr_batch_nonce']) && is_string($_POST['nsr_batch_nonce']) ? sanitize_text_field(wp_unslash($_POST['nsr_batch_nonce'])) : '';
        $text = isset($_POST['nsr_batch_serials']) && is_string($_POST['nsr_batch_serials']) ? sanitize_textarea_field(wp_unslash($_POST['nsr_batch_serials'])) : '';
        if (!wp_verify_nonce($nonce, 'nsr_batch_search_' . $context)) {
            $error = 'A consulta expirou. Recarregue a pagina e tente novamente.';
        } elseif (strlen($text) > 50000) {
            $error = 'A lista e muito grande. Consulte ate 500 NS por vez.';
        } else {
            $serials = nsr_parse_batch_serials($text);
            if (!$serials || count($serials) > 500) {
                $error = 'Informe de 1 a 500 NS diferentes por consulta.';
            } else {
                $groups = nsr_find_batch_serials($serials);
                if (is_wp_error($groups)) {
                    $error = $groups->get_error_message();
                    $groups = null;
                }
            }
        }
    }
    $id = 'nsr-batch-' . $context;
    ?>
    <div class="nsr-batch-search" style="margin-top:24px;">
        <h3>Consultar varios NS</h3>
        <p>Cole a lista ou use o leitor com Enter ao final de cada NS. Aceita linhas, espacos, virgulas e ponto e virgula. NS repetidos sao consultados uma vez. Busca exata, ate 500 NS por consulta.</p>
        <form method="post">
            <input type="hidden" name="nsr_batch_context" value="<?php echo esc_attr($context); ?>" />
            <?php wp_nonce_field('nsr_batch_search_' . $context, 'nsr_batch_nonce'); ?>
            <label for="<?php echo esc_attr($id); ?>">Lista de numeros de serie</label><br />
            <textarea id="<?php echo esc_attr($id); ?>" name="nsr_batch_serials" rows="8" style="width:100%;max-width:760px;" required placeholder="Um NS por linha"><?php echo esc_textarea($text); ?></textarea>
            <p>Garantia: 2 anos a partir da data da compra/nota fiscal registrada em Data Venda.</p>
            <button type="submit" class="button button-primary">Consultar lista de NS</button>
        </form>
        <?php if ($error !== '') : ?>
            <p role="alert"><?php echo esc_html($error); ?></p>
        <?php elseif (is_array($groups)) : ?>
            <?php $found = count(array_filter($groups)); ?>
            <p role="status"><strong><?php echo esc_html((string) count($groups)); ?></strong> NS consultados; <strong><?php echo esc_html((string) $found); ?></strong> encontrados; <strong><?php echo esc_html((string) (count($groups) - $found)); ?></strong> nao encontrados.</p>
            <p>Expiracao: data da compra/nota fiscal + 2 anos. Todas as notas encontradas para cada NS aparecem abaixo.</p>
            <div style="overflow-x:auto;">
            <table class="widefat striped" style="width:100%;text-align:left;">
                <thead><tr><th>NS consultado</th><th>Resultado</th><th>NS encontrado</th><th>Nota fiscal</th><th>Pedido</th><th>SKU</th><th>Descricao</th><th>GTIN/EAN</th><th>Data de venda</th><th>Expiracao</th></tr></thead>
                <tbody>
                <?php foreach ($groups as $key => $rows) : ?>
                    <?php if (!$rows) : ?>
                        <tr><td><?php echo esc_html($serials[$key]); ?></td><td>Nao encontrado</td><td colspan="8">&mdash;</td></tr>
                    <?php else : foreach ($rows as $row) : ?>
                        <tr>
                            <td><?php echo esc_html($serials[$key]); ?></td><td>Encontrado</td>
                            <?php foreach (array('ns', 'nota_fiscal', 'pedido', 'sku', 'descricao', 'gtin', 'data_venda') as $column) : ?>
                                <td><?php echo esc_html($row[$column] !== '' ? ($column === 'data_venda' ? nsr_format_sale_date($row[$column]) : $row[$column]) : 'Nao informado'); ?></td>
                            <?php endforeach; ?>
                            <td><?php echo esc_html(nsr_batch_expiration($row['data_venda'])); ?></td>
                        </tr>
                    <?php endforeach; endif; ?>
                <?php endforeach; ?>
                </tbody>
            </table>
            </div>
        <?php endif; ?>
    </div>
    <?php
}
