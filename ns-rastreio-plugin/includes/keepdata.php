<?php
defined('ABSPATH') || exit;

function nsr_record_origin_label($row) {
    return isset($row['origem_arquivo']) && (string) $row['origem_arquivo'] === 'keepdata-api' ? 'PY' : 'BR';
}

function nsr_get_keepdata_settings() {
    $url = defined('NSR_KEEPDATA_API_URL')
        ? (string) NSR_KEEPDATA_API_URL
        : (string) get_option('nsr_keepdata_api_url', 'https://api.keepdata.net.br:8765');
    $token = defined('NSR_KEEPDATA_API_TOKEN')
        ? (string) NSR_KEEPDATA_API_TOKEN
        : (string) get_option('nsr_keepdata_api_token', '');
    return array('url' => untrailingslashit(trim($url)), 'token' => trim($token));
}

function nsr_handle_keepdata_settings_submission() {
    $messages = array('success' => array(), 'error' => array());
    if (!isset($_POST['nsr_keepdata_save'])) {
        return $messages;
    }
    if (!current_user_can('manage_options')) {
        $messages['error'][] = 'Permissao insuficiente para configurar o Keepdata.';
        return $messages;
    }
    check_admin_referer('nsr_keepdata_settings', 'nsr_keepdata_nonce');
    if (defined('NSR_KEEPDATA_API_URL') || defined('NSR_KEEPDATA_API_TOKEN')) {
        $messages['error'][] = 'A integracao Keepdata esta definida por constantes no wp-config.php.';
        return $messages;
    }

    $url = isset($_POST['nsr_keepdata_api_url']) ? esc_url_raw(wp_unslash($_POST['nsr_keepdata_api_url'])) : '';
    $token = isset($_POST['nsr_keepdata_api_token']) ? trim((string) wp_unslash($_POST['nsr_keepdata_api_token'])) : '';
    if ($url === '' || stripos($url, 'https://') !== 0) {
        $messages['error'][] = 'Informe uma URL HTTPS valida para a API Keepdata.';
        return $messages;
    }
    update_option('nsr_keepdata_api_url', untrailingslashit($url), false);
    if ($token !== '') {
        update_option('nsr_keepdata_api_token', sanitize_text_field($token), false);
    }
    $messages['success'][] = 'Configuracao Keepdata salva.';
    return $messages;
}

function nsr_render_keepdata_settings() {
    $settings = nsr_get_keepdata_settings();
    $locked = defined('NSR_KEEPDATA_API_URL') || defined('NSR_KEEPDATA_API_TOKEN');
    ?>
    <div style="padding:12px 16px;background:#f6f7f7;border:1px solid #dcdcde;border-radius:6px;margin-bottom:20px;">
        <h3 style="margin-top:0;">Integracao Keepdata</h3>
        <p>NS ausentes na base local sao consultados na API e salvos automaticamente. Pedido e GTIN/EAN permanecem vazios quando nao forem fornecidos.</p>
        <p><strong>Status:</strong> <?php echo $settings['token'] !== '' ? 'configurada' : 'token nao configurado'; ?></p>
        <?php if (!$locked) : ?>
        <form method="post">
            <?php wp_nonce_field('nsr_keepdata_settings', 'nsr_keepdata_nonce'); ?>
            <input type="hidden" name="nsr_tab" value="tiny" />
            <p><label>URL da API<br><input type="url" name="nsr_keepdata_api_url" value="<?php echo esc_attr($settings['url']); ?>" class="regular-text" required /></label></p>
            <p><label>Bearer Token<br><input type="password" name="nsr_keepdata_api_token" value="" class="regular-text" autocomplete="new-password" placeholder="Deixe vazio para manter o token atual" /></label></p>
            <button type="submit" name="nsr_keepdata_save" class="button">Salvar integracao</button>
        </form>
        <?php else : ?>
        <p>Configuracao protegida por constantes no <code>wp-config.php</code>.</p>
        <?php endif; ?>
    </div>
    <?php
}

function nsr_keepdata_fetch_and_cache($serial) {
    $serial = trim((string) $serial);
    if (!preg_match('/^\d{4,30}$/', $serial)) {
        return new WP_Error('nsr_keepdata_invalid_serial', 'O Keepdata aceita somente NS numerico de 4 a 30 digitos.');
    }
    $settings = nsr_get_keepdata_settings();
    if ($settings['token'] === '') {
        return new WP_Error('nsr_keepdata_not_configured', 'O NS nao existe na base local e o token da integracao Keepdata nao foi configurado.');
    }

    $response = wp_remote_post($settings['url'] . '/v1/garantias/consultar', array(
        'timeout' => 90,
        'redirection' => 0,
        'sslverify' => true,
        'headers' => array(
            'Authorization' => 'Bearer ' . $settings['token'],
            'Accept' => 'application/json',
            'Content-Type' => 'application/json',
        ),
        'body' => wp_json_encode(array('serial_number' => $serial)),
    ));
    if (is_wp_error($response)) {
        return new WP_Error('nsr_keepdata_connection', 'Nao foi possivel consultar o Keepdata: ' . $response->get_error_message());
    }

    $status = (int) wp_remote_retrieve_response_code($response);
    $data = json_decode((string) wp_remote_retrieve_body($response), true);
    if (!is_array($data)) {
        return new WP_Error('nsr_keepdata_invalid_json', 'A API Keepdata retornou uma resposta invalida.');
    }
    if ($status !== 200 || empty($data['success'])) {
        $message = isset($data['error']['message']) ? sanitize_text_field((string) $data['error']['message']) : 'Falha ao consultar o Keepdata.';
        return new WP_Error('nsr_keepdata_api_error', $message);
    }
    if (empty($data['found'])) {
        return null;
    }

    $api_serial = isset($data['serial_number']) ? preg_replace('/\D+/', '', (string) $data['serial_number']) : $serial;
    $payload = array(
        'ns' => $api_serial !== '' ? $api_serial : $serial,
        'ns_normalizado' => nsr_normalize_lookup_value($api_serial !== '' ? $api_serial : $serial),
        'nota_fiscal' => isset($data['document_number']) ? sanitize_text_field((string) $data['document_number']) : '',
        'pedido' => '',
        'sku' => isset($data['product_code']) ? sanitize_text_field((string) $data['product_code']) : '',
        'descricao' => isset($data['product_description']) ? sanitize_text_field((string) $data['product_description']) : '',
        'quantidade' => '',
        'valor' => '',
        'data_venda' => isset($data['document_date']) ? sanitize_text_field((string) $data['document_date']) : '',
        'origem_arquivo' => 'keepdata-api',
        'linha_origem' => 0,
    );
    if (nsr_upsert_ns_record($payload) === false) {
        return new WP_Error('nsr_keepdata_save_failed', 'O Keepdata respondeu, mas o registro nao pode ser salvo no plugin.');
    }
    return $payload;
}
