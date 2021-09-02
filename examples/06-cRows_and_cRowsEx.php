<?php

ini_set('error_reporting', E_ALL);
ini_set('display_errors', true);

require_once __DIR__.'/../src/SimpleXLSX.php';

if ( $xlsx = SimpleXLSX::parse('books.xlsx')) {

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
	echo '</pre>';

	echo '<hr>';

	echo '<pre>';

	echo '<h2>$xlsx->cRowsEx()</h2>';
	echo '<pre>';
	print_r( $xlsx->cRowsEx($columns) );
	echo '</pre>';

} else {
	echo SimpleXLSX::parseError();
}