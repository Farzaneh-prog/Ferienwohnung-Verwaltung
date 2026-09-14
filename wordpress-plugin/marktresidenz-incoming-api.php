<?php
/**
 * Plugin Name: Marktresidenz — Incoming Reservations API
 * Description: Ein kleiner, eigenständiger Eingangs-Briefkasten für die Ferienwohnung-Verwaltung-Automatisierung (Cowork -> hier -> lokales apply-Skript). Rührt keine andere Seite, kein Buchungssystem, keine bestehende Datenbanktabelle an - nur eine eigene, neue Tabelle.
 * Version: 1.0.0
 * Author: Marktresidenz
 */

if (!defined('ABSPATH')) {
    exit;
}

define('MRV_TABLE_SUFFIX', 'mrv_incoming');

/**
 * Secrets are two PHP constants (MRV_WRITE_KEY, MRV_READ_KEY) — NOT in the
 * database and NOT in this file. Defined in a separate wp-content/mu-plugins/
 * file (see marktresidenz-secrets.php, never committed to git) so they load
 * automatically without touching wp-config.php. See README section at the
 * bottom of this file for the exact setup.
 */
function mrv_write_key() {
    return defined('MRV_WRITE_KEY') ? MRV_WRITE_KEY : null;
}
function mrv_read_key() {
    return defined('MRV_READ_KEY') ? MRV_READ_KEY : null;
}

register_activation_hook(__FILE__, function () {
    global $wpdb;
    $table = $wpdb->prefix . MRV_TABLE_SUFFIX;
    $charset_collate = $wpdb->get_charset_collate();
    $sql = "CREATE TABLE IF NOT EXISTS $table (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
        received_at DATETIME NOT NULL,
        run_date VARCHAR(10) NOT NULL,
        payload LONGTEXT NOT NULL,
        applied TINYINT(1) NOT NULL DEFAULT 0,
        PRIMARY KEY (id)
    ) $charset_collate;";
    require_once ABSPATH . 'wp-admin/includes/upgrade.php';
    dbDelta($sql);
});

add_action('rest_api_init', function () {
    register_rest_route('marktresidenz/v1', '/incoming', [
        [
            'methods'             => 'POST',
            'callback'            => 'mrv_handle_post_incoming',
            'permission_callback' => 'mrv_check_write_key',
        ],
        [
            'methods'             => 'GET',
            'callback'            => 'mrv_handle_get_incoming',
            'permission_callback' => 'mrv_check_read_key',
        ],
    ]);
    register_rest_route('marktresidenz/v1', '/incoming/(?P<id>\d+)/ack', [
        'methods'             => 'POST',
        'callback'            => 'mrv_handle_ack_incoming',
        'permission_callback' => 'mrv_check_read_key',
    ]);
});

function mrv_check_write_key(WP_REST_Request $request) {
    $key = $request->get_header('x-mrv-write-key');
    $expected = mrv_write_key();
    return $expected && $key && hash_equals($expected, $key);
}

function mrv_check_read_key(WP_REST_Request $request) {
    $key = $request->get_header('x-mrv-read-key');
    $expected = mrv_read_key();
    return $expected && $key && hash_equals($expected, $key);
}

// Cowork calls this once per day: writes everything it found (over-fetch is
// fine — the local apply script decides what to actually use).
function mrv_handle_post_incoming(WP_REST_Request $request) {
    global $wpdb;
    $table = $wpdb->prefix . MRV_TABLE_SUFFIX;
    $body = $request->get_json_params();

    if (!is_array($body) || empty($body['run_date'])) {
        return new WP_Error('mrv_bad_request', 'run_date ist erforderlich', ['status' => 400]);
    }

    $wpdb->insert($table, [
        'received_at' => current_time('mysql'),
        'run_date'    => sanitize_text_field($body['run_date']),
        'payload'     => wp_json_encode($body),
        'applied'     => 0,
    ]);

    return ['ok' => true, 'id' => $wpdb->insert_id];
}

// The local apply script calls this whenever the computer is next on: fetch
// everything not yet applied to the real Excel files.
function mrv_handle_get_incoming(WP_REST_Request $request) {
    global $wpdb;
    $table = $wpdb->prefix . MRV_TABLE_SUFFIX;
    $rows = $wpdb->get_results(
        "SELECT id, received_at, run_date, payload FROM $table WHERE applied = 0 ORDER BY id ASC",
        ARRAY_A
    );
    foreach ($rows as &$row) {
        $row['payload'] = json_decode($row['payload'], true);
    }
    return ['ok' => true, 'entries' => $rows];
}

// The apply script calls this right after it has safely written an entry
// into the real Excel file (and backed it up) — marks it done so it's never
// re-applied. Entries are kept (applied=1) rather than deleted, as a log.
function mrv_handle_ack_incoming(WP_REST_Request $request) {
    global $wpdb;
    $table = $wpdb->prefix . MRV_TABLE_SUFFIX;
    $id = (int) $request['id'];
    $wpdb->update($table, ['applied' => 1], ['id' => $id]);
    return ['ok' => true];
}

/*
=================================== SETUP ===================================

1. Diese Datei nach wp-content/plugins/marktresidenz-incoming-api/ hochladen
   (als marktresidenz-incoming-api.php) und im WordPress-Adminbereich unter
   "Plugins" aktivieren.

2. In wp-content/mu-plugins/ (Ordner ggf. neu anlegen) eine Datei
   marktresidenz-secrets.php hochladen mit:

   <?php
   define('MRV_WRITE_KEY', '...');   // Cowork benutzt diesen zum Schreiben
   define('MRV_READ_KEY', '...');    // das lokale apply-Skript zum Lesen

   Kein Aktivieren nötig (mu-plugins laden automatisch). Die Werte NICHT in
   dieses öffentliche Repo committen — diese Datei ist per .gitignore
   ausgeschlossen.

3. Endpunkte (Basis: https://marktresidenz-eisenach.de/wp-json/marktresidenz/v1):
   POST /incoming              Header: X-MRV-Write-Key: <WRITE_KEY>
   GET  /incoming               Header: X-MRV-Read-Key:  <READ_KEY>
   POST /incoming/{id}/ack      Header: X-MRV-Read-Key:  <READ_KEY>

===============================================================================
*/
