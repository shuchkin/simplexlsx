<?php /** @noinspection ForgottenDebugOutputInspection */
echo '<h1>rows() and rowsEx()</h1>';
if ( $xlsx = SimpleXLSX::parse('books.xlsx')) {
	echo '<h2>$xlsx->rows()</h2>';
	echo '<pre>';
	print_r( $xlsx->rows() );
	echo '</pre>';

	echo '<h2>$xlsx->rowsEx()</h2>';
	echo '<pre>';
	print_r( $xlsx->rowsEx() );
	echo '</pre>';
} else {
	echo SimpleXLSX::parseError();
}