<?php

if (isset($_FILES['file'])) {
	
	require_once "simplexlsx.class.php";
	
	$xlsx = new SimpleXLSX( $_FILES['file']['tmp_name'] );
	
	echo '<h1>Parsing Result</h1>';
	echo '<table border="1" cellpadding="3" style="border-collapse: collapse">';
	
	list($cols,) = $xlsx->dimension();
	
	foreach( $xlsx->rows() as $k => $r) {
//		if ($k == 0) continue; // skip first row
		echo '<tr>';
		for( $i = 0; $i < $cols; $i++)
			echo '<td>'.( (isset($r[$i])) ? $r[$i] : '&nbsp;' ).'</td>';
		echo '</tr>';
	}
	echo '</table>';
}

?>
<h1>Upload</h1>
<form method="post" enctype="multipart/form-data">
*.XLSX <input type="file" name="file"  />&nbsp;&nbsp;<input type="submit" value="Parse" />
</form>
