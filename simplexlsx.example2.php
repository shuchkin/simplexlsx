<?php

require_once __DIR__ . '/simplexlsx.class.php';

if ( $xlsx = SimpleXLSX::parse('countries_and_population.xlsx')) {

	echo '<table cellpadding="10">
	<tr><td valign="top">';

	// output worsheet 1

	list( $num_cols, $num_rows ) = $xlsx->dimension();

	echo '<h1>Sheet 1</h1>';
	echo '<table>';
	foreach ( $xlsx->rows( 1 ) as $r ) {
		echo '<tr>';
		for ( $i = 0; $i < $num_cols; $i ++ ) {
			echo '<td>' . ( ( ! empty( $r[ $i ] ) ) ? $r[ $i ] : '&nbsp;' ) . '</td>';
		}
		echo '</tr>';
	}
	echo '</table>';

	echo '</td><td valign="top">';

	// output worsheet 2

	list( $num_cols, $num_rows ) = $xlsx->dimension( 2 );

	echo '<h1>Sheet 2</h1>';
	echo '<table>';
	foreach ( $xlsx->rows( 2 ) as $r ) {
		echo '<tr>';
		for ( $i = 0; $i < $num_cols; $i ++ ) {
			echo '<td>' . ( ( ! empty( $r[ $i ] ) ) ? $r[ $i ] : '&nbsp;' ) . '</td>';
		}
		echo '</tr>';
	}
	echo '</table>';

	echo '</td></tr></table>';
} else {
	echo SimpleXLSX::parse_error();
}
?>