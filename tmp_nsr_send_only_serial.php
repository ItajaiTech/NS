<?php
require 'C:/xampp/htdocs/wordpress/wp-load.php';
$session = nsr_get_scan_session('ATA9rHbujU4E6IGeeN2HDMhG');
if (empty($session)) { echo "SESSAO_NAO_ENCONTRADA\n"; exit(1); }
$res = nsr_send_serials_to_tiny_order($session, 'kdt');
if (is_wp_error($res)) {
  echo 'ERR=' . $res->get_error_message() . "\n";
} else {
  echo "OK\n";
}
