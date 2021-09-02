<?php
ini_set('error_reporting', E_ALL);
ini_set('display_errors', true);

require_once __DIR__.'/../src/SimpleXLSX.php';

echo '<h1>Parse books.xslx</h1><pre>';
if ( $xlsx = SimpleXLSX::parse('books.xlsx') ) {
	print_r( $xlsx->rows() );
} else {
	echo SimpleXLSX::parseError();
}
echo '<pre>';


echo '<hr>';

$columns = array(
	'isbn',
	'title',
	'author',
	'publisher',
	'ctry'
);
echo '<h1>cRows (column) Parse books.xlsx</h1><pre>';
if ( $xls = SimpleXLSX::parse('books.xlsx') ) {
	print_r( $xls->cRows($columns) );
} else {
	echo SimpleXLSX::parseError();
}
echo '<pre>';