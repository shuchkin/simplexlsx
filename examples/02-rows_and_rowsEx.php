<?php /** @noinspection ForgottenDebugOutputInspection */

use Shuchkin\SimpleXLSX;

ini_set('error_reporting', E_ALL);
ini_set('display_errors', true);

require_once __DIR__.'/../src/SimpleXLSX.php';

echo '<h1>rows() and rowsEx()</h1>';
if ($xlsx = SimpleXLSX::parse('books.xlsx')) {
    // ->rows()
    echo '<h2>$xlsx->rows()</h2>';
    echo '<pre>';
    print_r($xlsx->rows());
    echo '</pre>';

    // ->rowsEx();
    echo '<h2>$xlsx->rowsEx()</h2>';
    echo '<pre>';
    print_r($xlsx->rowsEx());
    echo '</pre>';
} else {
    echo SimpleXLSX::parseError();
}
