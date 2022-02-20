<?php /** @noinspection MultiAssignmentUsageInspection */

use Shuchkin\SimpleXLSX;

ini_set('error_reporting', E_ALL);
ini_set('display_errors', true);

require_once __DIR__.'/../src/SimpleXLSX.php';

echo '<h1>Read several sheets</h1>';
if ($xlsx = SimpleXLSX::parse('countries_and_population.xlsx')) {
    echo '<pre>'.print_r($xlsx->sheetNames(), true).'</pre>';

    echo '<table cellpadding="10">
	<tr><td valign="top">';

    // output worksheet 1 (index = 0)

    $dim = $xlsx->dimension();
    $num_cols = $dim[0];
    $num_rows = $dim[1];

    echo '<h2>'.$xlsx->sheetName(0).'</h2>';
    echo '<table border=1>';
    foreach ($xlsx->rows() as $r) {
        echo '<tr>';
        for ($i = 0; $i < $num_cols; $i ++) {
            echo '<td>' . ( ! empty($r[ $i ]) ? $r[ $i ] : '&nbsp;' ) . '</td>';
        }
        echo '</tr>';
    }
    echo '</table>';

    echo '</td><td valign="top">';

    // output worsheet 2 (index = 1)

    $dim = $xlsx->dimension(1);
    $num_cols = $dim[0];
    $num_rows = $dim[1];

    echo '<h2>'.$xlsx->sheetName(1).'</h2>';
    echo '<table border=1>';
    foreach ($xlsx->rows(1) as $r) {
        echo '<tr>';
        for ($i = 0; $i < $num_cols; $i ++) {
            echo '<td>' . ( ! empty($r[ $i ]) ? $r[ $i ] : '&nbsp;' ) . '</td>';
        }
        echo '</tr>';
    }
    echo '</table>';

    echo '</td></tr></table>';
} else {
    echo SimpleXLSX::parseError();
}
