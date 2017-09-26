<?php
echo '<html><body><h1>XLSX to HTML</h1>';

if (isset($_FILES['file'])) {
	
	require_once __DIR__ . '/simplexlsx.class.php';
	
	if ( $xlsx = SimpleXLSX::parse( $_FILES['file']['tmp_name'] ) ) {

		echo '<h2>Parsing Result</h2>';
		echo '<table border="1" cellpadding="3" style="border-collapse: collapse">';

		list( $cols, ) = $xlsx->dimension();

		foreach ( $xlsx->rows() as $k => $r ) {
			//		if ($k == 0) continue; // skip first row
			echo '<tr>';
			for ( $i = 0; $i < $cols; $i ++ ) {
				echo '<td>' . ( ( isset( $r[ $i ] ) ) ? $r[ $i ] : '&nbsp;' ) . '</td>';
			}
			echo '</tr>';
		}
		echo '</table>';
	} else {
		echo SimpleXLSX::parse_error();
	}
}
echo '<h2>Upload form</h2>
<form method="post" enctype="multipart/form-data">
*.XLSX <input type="file" name="file"  />&nbsp;&nbsp;<input type="submit" value="Parse" />
</form>';
echo '</body></html>';
