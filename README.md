# SimpleXLSX class 0.7.12 (Official)

Parse and retrieve data from Excel XLSx files. MS Excel 2007 workbooks PHP reader.

**Sergey Shuchkin** <sergey.shuchkin@gmail.com> 2010-2018

## Basic Usage
```php
if ( $xlsx = SimpleXLSX::parse('book.xlsx') ) {
	print_r( $xlsx->rows() );
} else {
	echo SimpleXLSX::parse_error();
}
```
```
Array
(
    [0] => Array
        (
            [0] => ISBN
            [1] => title
            [2] => author
            [3] => publisher
            [4] => ctry
        )

    [1] => Array
        (
            [0] => 618260307
            [1] => The Hobbit
            [2] => J. R. R. Tolkien
            [3] => Houghton Mifflin
            [4] => USA
        )

    [2] => Array
        (
            [0] => 908606664
            [1] => Slinky Malinki
            [2] => Lynley Dodd
            [3] => Mallinson Rendel
            [4] => NZ
        )

)
```
## Examples
```php
// XLSX to html table
if ( $xlsx = SimpleXLSX::parse('book.xlsx') ) {
	echo '<table>';
	foreach( $xlsx->rows() as $r ) {
		echo '<tr><td>'.implode('</td><td>', $r ).'</td></tr>';
	}
	echo '</table>';
} else {
	echo SimpleXLSX::parse_error();
}

// rowsEx() 
$xlsx = SimpleXLSX::parse('book.xlsx');
print_r( $xlsx->rowsEx() );

// Select Sheet
$xlsx = SimpleXLSX::parse('book.xlsx');
print_r( $xlsx->rows(1) ); // Sheet numeration started 0, we select second worksheet

// List sheets
$xlsx = SimpleXLSX::parse('book.xlsx');
print_r( $xlsx->sheetNames() ); // array( 0 => 'Sheet 1', 1 => 'Catalog' );

// Sheet by id
$xlsx = SimpleXLSX::parse('book.xlsx');	
echo 'Sheet Name 2 = '.$xlsx->sheetName(1);

// XLSX::parse remote data 
$data = file_get_contents('http://www.example.com/example.xlsx');
if ( $xlsx = SimpleXLSX::parse( $data, true) ) {
	list($num_cols, $num_rows) = $xlsx->dimension(1);
	echo $xlsx->sheetName(1).':'.$num_cols.'x'.$num_rows;
} else {
	echo SimpleXLSX::parse_error();
}

// Classic OOP style 
$xlsx = new SimpleXLSX('books.xlsx');
if ($xslx->success()) {
	print_r( $xlsx->rows() );
} else {
	echo 'xlsx error: '.$xlsx->error();
}
```
	

## History
```
v0.7.12 (2018-06-17) $worksheet_id to $worksheet_index, sheet numeration started 0
v0.7.11 (2018-04-25) rowsEx(), added row index "r" to cell info
v0.7.10 (2018-04-21) fixed getCell, returns NULL if not exits
v0.7.9 (2018-03-17) fixed sheetNames() (namespaced or not namespaced attr)
v0.7.8 (2018-01-15) remove namespace prefixes (hardcoded)
v0.7.7 (2017-10-02) XML External Entity (XXE) Prevention (<!ENTITY xxe SYSTEM "file: ///etc/passwd" >]>)
v0.7.6 (2017-09-26) if worksheet_id === 0 (default) then detect first sheet (for LibreOffice capabilities)
v0.7.5 (2017-09-10) ->getCell() - fixed
v0.7.4 (2017-08-22) ::parse_error() - to get last error in "static style"
v0.7.3 (2017-08-14) ->_parse fixed relations reader, added ->getCell( sheet_id, address, format ) for direct cell reading
v0.7.2 (2017-05-13) ::parse( $filename ) helper method
v0.7.1 (2017-03-29) License added
v0.6.11 (2016-07-27) fixed timestamp()
v0.6.10 (2016-06-10) fixed search entries (UPPERCASE)
v0.6.9 (2015-04-12) $xlsx->datetime_format to force dates out
v0.6.8 (2013-10-13) fixed dimension() where 1 row only, fixed rowsEx() empty cells indexes (Daniel Stastka)
v0.6.7 (2013-08-10) fixed unzip (mac), added $debug param to _constructor to display errors
v0.6.6 (2013-06-03) +entryExists()
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
```