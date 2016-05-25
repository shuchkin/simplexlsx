<?php
/*
	SimpleXLSX php class v0.6.8
	MS Excel 2007 workbooks reader

	Example 1: 
	$xlsx = new SimpleXLSX('book.xlsx');
	print_r( $xlsx->rows() );

	Example 2: 
	$xlsx = new SimpleXLSX('book.xlsx');
	print_r( $xlsx->rowsEx() );

	Example 3: 
	$xlsx = new SimpleXLSX('book.xlsx');
	print_r( $xlsx->rows(2) ); // second worksheet

	Example 4.1:
	$xlsx = new SimpleXLSX('book.xlsx');
	print_r( $xlsx->sheetNames() ); // array( 1 => 'Sheet 1', 3 => 'Catalog' );
	
	Example 4.2:
	$xlsx = new SimpleXLSX('book.xlsx');	
	echo 'Sheet Name 2 = '.$xlsx->sheetName(2);
	
	Example 5:
	$xlsx = new SimpleXLSX('book.xlsx');
	if ($xslx->success())
		print_r( $xlsx->rows() );
	else
		echo 'xlsx error: '.$xslx->error();
	
	Example 6:
	$xslx = new SimpleXLSX( file_get_contents('http://www.example.com/example.xlsx'), true);
	list($num_cols, $num_rows) = $xlsx->dimension(2);
	echo $xlsx->sheetName(2).':'.$num_cols.'x'.$num_rows;

	v0.6.9 (2015-04-12) $xlsx->datetime_format to force dates out
	v0.6.8 (2013-10-13) fixed dimension() where 1 row only, fixed rowsEx() empty cells indexes (Daniel Stastka)
	v0.6.7 (2013-08-10) fixed unzip (mac), added $debug param to _constructor to display errors
	v0.6.6 (2013-06-03) +entryExists(),
	v0.6.5 (2013-03-18) fixed sheetName()
	v0.6.4 (2013-03-13) rowsEx(), _parse(): fixed date column type & format detection
	v0.6.3 (2013-03-13) rowsEx(): fixed formulas, added date type 'd', added format 'format'
						dimension(): fixed empty sheet dimension
	                    + sheetNames() - returns array( sheet_id => sheet_name, sheet_id2 => sheet_name2 ...)
	v0.6.2 (2012-10-04) fixed empty cells, rowsEx() returns type and formulas now
	v0.6.1 (2012-09-14) removed "raise exception" and fixed _unzip
	v0.6 (2012-09-13) success(), error(), __constructor( $filename, $is_data = false )
	v0.5.1 (2012-09-13) sheetName() fixed
	v0.5 (2012-09-12) sheetName()
	v0.4 sheets(), sheetsCount(), unixstamp( $excelDateTime )
	v0.3 - fixed empty cells (Gonzo patch)
*/

class SimpleXLSX {
	// Don't remove this string! Created by Sergey Schuchkin from http://www.sibvision.ru - professional php developers team 2010-2013
	private $workbook;
	private $sheets = array();
	private $styles;
	private $hyperlinks;
	private $package = array(
		'filename' => '',
		'mtime' => 0,
		'size' => 0,
		'comment' => '',
		'entries' => array()
	);
	private $datasec;
	private $sharedstrings;
	private $error = false;
	private $debug = false;
	// scheme
	const SCHEMA_REL_OFFICEDOCUMENT  =  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
	const SCHEMA_REL_SHAREDSTRINGS =  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
	const SCHEMA_REL_WORKSHEET =  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
	const SCHEMA_REL_STYLES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
	
	public $workbook_cell_formats = array();

/*
	private $date_formats = array(
		0xe => "d/m/Y",
		0xf => "d-M-Y",
		0x10 => "d-M",
		0x11 => "M-Y",
		0x12 => "h:i a",
		0x13 => "h:i:s a",
		0x14 => "H:i",
		0x15 => "H:i:s",
		0x16 => "d/m/Y H:i",
		0x2d => "i:s",
		0x2e => "H:i:s",
		0x2f => "i:s.S"
	);
	private $number_formats = array(
		0x1 => "%1.0f",     // "0"
		0x2 => "%1.2f",     // "0.00",
		0x3 => "%1.0f",     //"#,##0",
		0x4 => "%1.2f",     //"#,##0.00",
		0x5 => "%1.0f",     //"$#,##0;($#,##0)",
		0x6 => '$%1.0f',    //"$#,##0;($#,##0)",
		0x7 => '$%1.2f',    //"$#,##0.00;($#,##0.00)",
		0x8 => '$%1.2f',    //"$#,##0.00;($#,##0.00)",
		0x9 => '%1.0f%%',   //"0%"
		0xa => '%1.2f%%',   //"0.00%"
		0xb => '%1.2f',     //"0.00E00",
		0x25 => '%1.0f',    //"#,##0;(#,##0)",
		0x26 => '%1.0f',    //"#,##0;(#,##0)",
		0x27 => '%1.2f',    //"#,##0.00;(#,##0.00)",
		0x28 => '%1.2f',    //"#,##0.00;(#,##0.00)",
		0x29 => '%1.0f',    //"#,##0;(#,##0)",
		0x2a => '$%1.0f',   //"$#,##0;($#,##0)",
		0x2b => '%1.2f',    //"#,##0.00;(#,##0.00)",
		0x2c => '$%1.2f',   //"$#,##0.00;($#,##0.00)",
		0x30 => '%1.0f');   //"##0.0E0";
	// }}}
*/
	private $cell_formats = array(
		0 => 'General',
		1 => '0',
		2 => '0.00',
		3 => '#,##0',
		4 => '#,##0.00',
		9 => '0%',
		10 => '0.00%',
		11 => '0.00E+00',
		12 => '# ?/?',
		13 => '# ??/??',
		14 => 'mm-dd-yy',
		15 => 'd-mmm-yy',
		16 => 'd-mmm',
		17 => 'mmm-yy',
		18 => 'h:mm AM/PM',
		19 => 'h:mm:ss AM/PM',
		20 => 'h:mm',
		21 => 'h:mm:ss',
		22 => 'm/d/yy h:mm',
		
		37 => '#,##0 ;(#,##0)',
		38 => '#,##0 ;[Red](#,##0)',
		39 => '#,##0.00;(#,##0.00)',
		40 => '#,##0.00;[Red](#,##0.00)',
		
		44 => '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)',
		45 => 'mm:ss',
		46 => '[h]:mm:ss',
		47 => 'mmss.0',
		48 => '##0.0E+0',
		49 => '@',
		
		27 => '[$-404]e/m/d',
		30 => 'm/d/yy',
		36 => '[$-404]e/m/d',
		50 => '[$-404]e/m/d',
		57 => '[$-404]e/m/d',
		
		59 => 't0',
		60 => 't0.00',
		61 => 't#,##0',
		62 => 't#,##0.00',
		67 => 't0%',
		68 => 't0.00%',
		69 => 't# ?/?',
		70 => 't# ??/??',
	);
	public $datetime_format = 'Y-m-d H:i:s';
	
	function __construct( $filename, $is_data = false, $debug = false ) {
		$this->debug = $debug;
		if ( $this->_unzip( $filename, $is_data ) ) {
			$this->_parse();
		}
	}
	// sheets numeration: 1,2,3....
	public function rows( $worksheet_id = 1 ) {

		if (($ws = $this->worksheet( $worksheet_id)) === false)
			return false;

		$rows = array();
		$curR = 0;

		list($cols,) = $this->dimension( $worksheet_id );
		/* @var SimpleXMLElement $ws */
		foreach ($ws->sheetData->row as $row) {
//			echo 'row<br />';
			foreach ($row->c as $c) {
				list($curC,) = $this->_columnIndex((string) $c['r']);

				$rows[ $curR ][ $curC ] = $this->value($c);
			}
			for ($i = 0; $i < $cols; $i++)
				if (!isset($rows[$curR][$i]))
					$rows[ $curR ][ $i ] = '';

			ksort( $rows[ $curR ] );

			$curR++;
		}
		return $rows;
	}
	public function rowsEx( $worksheet_id = 1 ) {

		if (($ws = $this->worksheet( $worksheet_id)) === false)
			return false;

		$rows = array();
		$curR = 0;
		list($cols,) = $this->dimension( $worksheet_id );
		/* @var SimpleXMLElement $ws */
		foreach ($ws->sheetData->row as $row) {

			foreach ($row->c as $c) {
				list($curC,) = $this->_columnIndex((string) $c['r']);
				$t = (string)$c['t'];
				$s = (int)$c['s'];
				if ($s > 0 && isset($this->workbook_cell_formats[ $s ])) {
					$format = $this->workbook_cell_formats[ $s ]['format'];
					if ( strpos($format,'m') !== false )
						$t = 'd';
				} else
					$format = '';

				$rows[ $curR ][ $curC ] = array(
					'type' => $t,
					'name' => (string) $c['r'],
					'value' => $this->value($c, $format ),
					'href' => $this->href( $c ),
					'f' => (string) $c->f,
					'format' => $format
				);
			}
			for ($i = 0; $i < $cols; $i++) {

				if ( !isset($rows[$curR][$i]) ) {

					// 0.6.8
					for ($c = '', $j = $i; $j >= 0; $j = intval($j / 26) - 1)
						$c = chr( $j % 26 + 65 ) . $c;

					$rows[ $curR ][$i] = array(
						'type' => '',
//						'name' => chr($i + 65).($curR+1),
						'name' => $c.($curR+1),
						'value' => '',
						'href' => '',
						'f' => '',
						'format' => ''
					);
				}
			}

			ksort( $rows[ $curR ] );

			$curR++;
		}
		return $rows;

	}
	public function sheets() {
		return $this->sheets;
	}
	public function sheetsCount() {
		return count($this->sheets);
	}
	public function sheetName( $worksheet_id) {
		if ( !isset( $this->workbook->sheets->sheet) ) return false;
		foreach( $this->workbook->sheets->sheet as $s ) {
			/* @var SimpleXMLElement $s */
			if ( $s->attributes('r',true)->id == 'rId'.$worksheet_id)
				return (string) $s['name'];

		}
		return false;
	}
	public function sheetNames() {
		
		$result = array();

		foreach( $this->workbook->sheets->sheet as $s ) {
			/* @var SimpleXMLElement $s */
			$result[ substr( $s->attributes('r',true)->id, 3) ] = (string) $s['name'];

		}
		return $result;
	}
	public function worksheet( $worksheet_id ) {

		if ( isset( $this->sheets[ $worksheet_id ] ) ) {
			$ws = $this->sheets[ $worksheet_id ];
			
			if (isset($ws->hyperlinks)) {
				$this->hyperlinks = array();
				foreach( $ws->hyperlinks->hyperlink as $hyperlink ) {
					$this->hyperlinks[ (string) $hyperlink['ref'] ] = (string) $hyperlink['display'];
				}
			}
			
			return $ws;
		} else {
			$this->error( 'Worksheet '.$worksheet_id.' not found.');
			return false;
		}
	}
	public function dimension( $worksheet_id = 1 ) {
		
		if (($ws = $this->worksheet( $worksheet_id)) === false)
			return false;
		/* @var SimpleXMLElement $ws */
		$ref = (string) $ws->dimension['ref'];
		
		if( strpos( $ref, ':') !== false ) {
			$d = explode(':', $ref);
			if(!isset($d[1]))
				return array(0,0);
			$index = $this->_columnIndex( $d[1] );		
			return array( $index[0]+1, $index[1]+1);
		} else if ( strlen( $ref ) ) { // 0.6.8
			$index = $this->_columnIndex( $ref );		
			return array( $index[0]+1, $index[1]+1);
		} else
			return array(0,0);

	}

	public function value( $cell, $format = null ) {
		// Determine data type
		$dataType = (string) $cell['t'];

		if ( $format === NULL ) {
			$s = (int)$cell['s'];
			if ($s > 0 && isset($this->workbook_cell_formats[$s])) {
				$format = $this->workbook_cell_formats[$s]['format'];
			}
		}
		if ( strpos($format,'m') !== false ) {
			$dataType = 'd';
		}

		switch ($dataType) {
			case "s":
				// Value is a shared string
				if ((string)$cell->v != '') {
					$value = $this->sharedstrings[intval($cell->v)];
				} else {
					$value = '';
				}

				break;
				
			case "b":
				// Value is boolean
				$value = (string)$cell->v;
				if ($value == '0') {
					$value = false;
				} else if ($value == '1') {
					$value = true;
				} else {
					$value = (bool)$cell->v;
				}

				break;
				
			case "inlineStr":
				// Value is rich text inline
				$value = $this->_parseRichText($cell->is);
							
				break;
				
			case "e":
				// Value is an error message
				if ((string)$cell->v != '') {
					$value = (string)$cell->v;
				} else {
					$value = '';
				}

				break;
			case 'd':
				// Value is a date
				$value = ( $this->datetime_format ) ? gmdate( $this->datetime_format, $this->unixstamp( (float)$cell->v )) : (float) $cell->v ;
				break;


			default:
				// Value is a string
				$value = (string)$cell->v;
								
				// Check for numeric values
				if (is_numeric($value) && $dataType != 's') {
					if ($value == (int)$value) $value = (int)$value;
					elseif ($value == (float)$value) $value = (float)$value;
				}
		}
		return $value;
	}
	public function href( $cell ) {
		return isset( $this->hyperlinks[ (string) $cell['r'] ] ) ? $this->hyperlinks[ (string) $cell['r'] ] : '';
	}
	public function styles() {
		return $this->styles;
	}
	private function _unzip( $filename, $is_data = false ) {
		
		// Clear current file
		$this->datasec = array();
			
		if ($is_data) {

			$this->package['filename'] = 'default.xlsx';
			$this->package['mtime'] = time();
			$this->package['size'] = strlen( $filename );
			
			$vZ = $filename;
		} else {
			
			if (!is_readable($filename)) {
				$this->error( 'File not found '.$filename );
				return false;
			}
			
			// Package information
			$this->package['filename'] = $filename;
			$this->package['mtime'] = filemtime( $filename );
			$this->package['size'] = filesize( $filename );

			// Read file
			$vZ = file_get_contents( $filename );
		}
		// Cut end of central directory
/*		$aE = explode("\x50\x4b\x05\x06", $vZ);
		
		if (count($aE) == 1) {
			$this->error('Unknown format');
			return false;
		}
*/
		if ( ($pcd = strrpos( $vZ, "\x50\x4b\x05\x06" )) === false ) {
			$this->error('Unknown format');
			return false;
		}
		$aE = array(
			0 => substr( $vZ, 0, $pcd ),
			1 => substr( $vZ, $pcd + 3 )
		);

		// Normal way
		$aP = unpack('x16/v1CL', $aE[1]);
		$this->package['comment'] = substr($aE[1], 18, $aP['CL']);

		// Translates end of line from other operating systems
		$this->package['comment'] = strtr($this->package['comment'], array("\r\n" => "\n", "\r" => "\n"));

		// Cut the entries from the central directory
		$aE = explode("\x50\x4b\x01\x02", $vZ);
		// Explode to each part
		$aE = explode("\x50\x4b\x03\x04", $aE[0]);
		// Shift out spanning signature or empty entry
		array_shift($aE);

		// Loop through the entries
		foreach ($aE as $vZ) {
			$aI = array();
			$aI['E']  = 0;
			$aI['EM'] = '';
			// Retrieving local file header information
//			$aP = unpack('v1VN/v1GPF/v1CM/v1FT/v1FD/V1CRC/V1CS/V1UCS/v1FNL', $vZ);
			$aP = unpack('v1VN/v1GPF/v1CM/v1FT/v1FD/V1CRC/V1CS/V1UCS/v1FNL/v1EFL', $vZ);

			// Check if data is encrypted
//			$bE = ($aP['GPF'] && 0x0001) ? TRUE : FALSE;
			$bE = false;
			$nF = $aP['FNL'];
			$mF = $aP['EFL'];

			// Special case : value block after the compressed data
			if ($aP['GPF'] & 0x0008) {
				$aP1 = unpack('V1CRC/V1CS/V1UCS', substr($vZ, -12));

				$aP['CRC'] = $aP1['CRC'];
				$aP['CS']  = $aP1['CS'];
				$aP['UCS'] = $aP1['UCS'];
				// 2013-08-10
				$vZ = substr($vZ, 0, -12);
				if (substr($vZ,-4) === "\x50\x4b\x07\x08")
					$vZ = substr($vZ, 0, -4);
			}

			// Getting stored filename
			$aI['N'] = substr($vZ, 26, $nF);

			if (substr($aI['N'], -1) == '/') {
				// is a directory entry - will be skipped
				continue;
			}

			// Truncate full filename in path and filename
			$aI['P'] = dirname($aI['N']);
			$aI['P'] = $aI['P'] == '.' ? '' : $aI['P'];
			$aI['N'] = basename($aI['N']);

			$vZ = substr($vZ, 26 + $nF + $mF);

			if ( strlen($vZ) != $aP['CS'] ) { // check only if availabled
			  $aI['E']  = 1;
			  $aI['EM'] = 'Compressed size is not equal with the value in header information.';
			} else {
				if ($bE) {
					$aI['E']  = 5;
					$aI['EM'] = 'File is encrypted, which is not supported from this class.';
				} else {
					switch($aP['CM']) {
						case 0: // Stored
							// Here is nothing to do, the file ist flat.
							break;
						case 8: // Deflated
							$vZ = gzinflate($vZ);
							break;
						case 12: // BZIP2
							if (extension_loaded('bz2')) {
								$vZ = bzdecompress($vZ);
							} else {
								$aI['E']  = 7;
								$aI['EM'] = "PHP BZIP2 extension not available.";
							}
							break;
						default:
						  $aI['E']  = 6;
						  $aI['EM'] = "De-/Compression method {$aP['CM']} is not supported.";
					}
					if (! $aI['E']) {
						if ($vZ === FALSE) {
							$aI['E']  = 2;
							$aI['EM'] = 'Decompression of data failed.';
						} else {
							if (strlen($vZ) != $aP['UCS']) {
								$aI['E']  = 3;
								$aI['EM'] = 'Uncompressed size is not equal with the value in header information.';
							} else {
								if (crc32($vZ) != $aP['CRC']) {
									$aI['E']  = 4;
									$aI['EM'] = 'CRC32 checksum is not equal with the value in header information.';
								}
							}
						}
					}
				}
			}

			$aI['D'] = $vZ;

			// DOS to UNIX timestamp
			$aI['T'] = mktime(($aP['FT']  & 0xf800) >> 11,
							  ($aP['FT']  & 0x07e0) >>  5,
							  ($aP['FT']  & 0x001f) <<  1,
							  ($aP['FD']  & 0x01e0) >>  5,
							  ($aP['FD']  & 0x001f),
							  (($aP['FD'] & 0xfe00) >>  9) + 1980);

			//$this->Entries[] = &new SimpleUnzipEntry($aI);
			$this->package['entries'][] = array(
				'data' => $aI['D'],
				'error' => $aI['E'],
				'error_msg' => $aI['EM'],
				'name' => $aI['N'],
				'path' => $aI['P'],
				'time' => $aI['T']
			);

		} // end for each entries
		return true;
	}
	public function getPackage() {
		return $this->package;
	}
	public function entryExists( $name ) { // 0.6.6
		$dir = dirname( $name );
		$name = basename( $name );
		foreach( $this->package['entries'] as $entry)
			if ( $entry['path'] == $dir && $entry['name'] == $name)
				return true;
		return false;
	}
	public function getEntryData( $name ) {
		$dir = dirname( $name );
		$name = basename( $name );
		foreach( $this->package['entries'] as $entry)
			if ( $entry['path'] == $dir && $entry['name'] == $name)
				return $entry['data'];
		$this->error('Unknown format');
		return false;
	}
	public function getEntryXML( $name ) {
		if ( ($entry_xml = $this->getEntryData( $name ))
			&&  ($entry_xmlobj = simplexml_load_string( $entry_xml )))
			return $entry_xmlobj;

			
		$this->error('Entry not found: '.$name );
		return false;
	}
	public function error( $set = false ) {
		if ($set) {
			$this->error = $set;
			if ($this->debug)
				trigger_error( __CLASS__.': '.$set, E_USER_WARNING );
		} else {
			return $this->error;
		}
	}
	public function success() {
		return !$this->error;
	}
	public function unixstamp( $excelDateTime ) {
		$d = floor( $excelDateTime ); // seconds since 1900
		$t = $excelDateTime - $d;
//		echo "d=".$d." t=".$t."\r\n";
//		return ($d > 0) ? ( $d - 25569 ) * 86400 + $t * 86400 : $t * 86400;
		return floor( ($d > 0) ? ( $d - 25568 ) * 86400 + $t * 86400 : $t * 86400 ); // Yuri Nunes
	}

	private function _parse() {
		// Document data holders
		$this->sharedstrings = array();
		$this->sheets = array();
//		$this->styles = array();

		// Read relations and search for officeDocument
		if ( $relations = $this->getEntryXML("_rels/.rels" ) ) {
			
			foreach ($relations->Relationship as $rel) {
				
				if ($rel["Type"] == SimpleXLSX::SCHEMA_REL_OFFICEDOCUMENT) {
					
//						echo 'workbook found<br />';
					// Found office document! Read workbook & relations...
					
					// Workbook
					if ( $this->workbook = $this->getEntryXML( $rel['Target'] )) {
						
//						echo 'workbook read<br />';
						
						if ( $workbookRelations = $this->getEntryXML( dirname($rel['Target']) . '/_rels/workbook.xml.rels' )) {
							
//							echo 'workbook relations<br />';
		
							// Loop relations for workbook and extract sheets...
							foreach ($workbookRelations->Relationship as $workbookRelation) {

								$path = dirname($rel['Target']) . '/' . $workbookRelation['Target'];
								
								if ($workbookRelation['Type'] == SimpleXLSX::SCHEMA_REL_WORKSHEET) { // Sheets
								
//									echo 'sheet<br />';
								
									if ( $sheet = $this->getEntryXML( $path ) ) {
										$this->sheets[ str_replace( 'rId', '', (string) $workbookRelation['Id']) ] = $sheet;
//										echo '<pre>'.htmlspecialchars( print_r( $sheet, true ) ).'</pre>';
									}
													
								} else if ($workbookRelation['Type'] == SimpleXLSX::SCHEMA_REL_SHAREDSTRINGS && $this->entryExists( $path )) { // 0.6.6
									
//									echo 'sharedstrings<br />';
									
									if ( $sharedStrings = $this->getEntryXML( $path ) ) {
										foreach ($sharedStrings->si as $val) {
											if (isset($val->t)) {
												$this->sharedstrings[] = (string)$val->t;
											} elseif (isset($val->r)) {
												$this->sharedstrings[] = $this->_parseRichText($val);
											}
										}
									}
								} else if ($workbookRelation['Type'] == SimpleXLSX::SCHEMA_REL_STYLES) {
									$this->styles = $this->getEntryXML( $path );
									
									$nf = array();
									if ( $this->styles->numFmts->numFmt != NULL )
										foreach( $this->styles->numFmts->numFmt as $v )
											$nf[ (int) $v['numFmtId'] ] = (string) $v['formatCode'];

									if ( $this->styles->cellXfs->xf != NULL )	
										foreach( $this->styles->cellXfs->xf as $v ) {
											$v = (array) $v->attributes();
											$v = $v['@attributes'];
											if (isset($this->cell_formats[ $v['numFmtId'] ]) ) {
												$v['format'] = $this->cell_formats[$v['numFmtId']];
											} else if (isset($nf[ $v['numFmtId'] ])) {
												$v['format'] = $nf[$v['numFmtId']];
											} else {
												$v['format'] = '';
											}
											$this->workbook_cell_formats[] = $v;
										}
//									print_r( $this->workbook_cell_formats );
								}
							}
							
							break;
						}
					}
				}
			}
		}
		// Sort sheets
		ksort($this->sheets);
	}
	// thx Gonzo
	private function _columnIndex( $cell = 'A1' ) {

		if (preg_match("/([A-Z]+)(\d+)/", $cell, $matches)) {

			$col = $matches[1];
			$row = $matches[2];

			$colLen = strlen($col);
			$index = 0;

			for ($i = $colLen-1; $i >= 0; $i--)
				$index += (ord($col{$i}) - 64) * pow(26, $colLen-$i-1);

			return array($index-1, $row-1);
		} else
			throw new Exception("Invalid cell index.");
	}
    private function _parseRichText($is = null) {
        $value = array();

        if (isset($is->t)) {
            $value[] = (string)$is->t;
        } else {
            foreach ($is->r as $run) {
                $value[] = (string)$run->t;
            }
        }

        return implode(' ', $value);
    }
}
?>