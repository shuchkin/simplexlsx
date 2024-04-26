# SimpleXLSX class (Official)
[<img src="https://img.shields.io/packagist/dt/shuchkin/simplexlsx" />](https://packagist.org/packages/shuchkin/simplexlsx)
[<img src="https://img.shields.io/github/license/shuchkin/simplexlsx" />](https://github.com/shuchkin/simplexlsx/blob/master/license.md) [<img src="https://img.shields.io/github/stars/shuchkin/simplexlsx" />](https://github.com/shuchkin/simplexlsx/stargazers) [<img src="https://img.shields.io/github/forks/shuchkin/simplexlsx" />](https://github.com/shuchkin/simplexlsx/network) [<img src="https://img.shields.io/github/issues/shuchkin/simplexlsx" />](https://github.com/shuchkin/simplexlsx/issues)
[<img src="https://img.shields.io/opencollective/all/simplexlsx" />](https://opencollective.com/simplexlsx)
[<img src="https://img.shields.io/badge/patreon-_-_" />](https://www.patreon.com/shuchkin)

Parse and retrieve data from Excel XLSx files. MS Excel 2007 workbooks PHP reader.
No addiditional extensions need (internal unzip + standart SimpleXML parser).

See also:<br/>
[SimpleXLS](https://github.com/shuchkin/simplexls) old format MS Excel 97 php reader.<br/>
[SimpleXLSXGen](https://github.com/shuchkin/simplexlsxgen) xlsx php writer.  

*Hey, bro, please ★ the package for my motivation :) and [donate](https://opencollective.com/simplexlsx) for more motivation!*

**Sergey Shuchkin** <sergey.shuchkin@gmail.com>

## Basic Usage
```php
use Shuchkin\SimpleXLSX;

if ( $xlsx = SimpleXLSX::parse('book.xlsx') ) {
    print_r( $xlsx->rows() );
} else {
    echo SimpleXLSX::parseError();
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

)
```
## Installation
The recommended way to install this library is [through Composer](https://getcomposer.org).
[New to Composer?](https://getcomposer.org/doc/00-intro.md)

This will install the latest supported version:
```bash
$ composer require shuchkin/simplexlsx
```
or download PHP 5.5+ class [here](https://github.com/shuchkin/simplexlsx/blob/master/src/SimpleXLSX.php)

## Basic methods
```
// open
SimpleXLSX::parse( $filename, $is_data = false, $debug = false ): SimpleXLSX (or false)
SimpleXLSX::parseFile( $filename, $debug = false ): SimpleXLSX (or false)
SimpleXLSX::parseData( $data, $debug = false ): SimpleXLSX (or false)
// simple
$xlsx->rows($worksheetIndex = 0, $limit = 0): array
$xlsx->readRows($worksheetIndex = 0, $limit = 0): Generator - helps read huge xlsx
$xlsx->toHTML($worksheetIndex = 0, $limit = 0): string
// extended
$xlsx->rowsEx($worksheetIndex = 0, $limit = 0): array
$xlsx->readRowsEx($worksheetIndex = 0, $limit = 0): Generator - helps read huge xlsx with styles
$xlsx->toHTMLEx($worksheetIndex = 0, $limit = 0): string
// meta
$xlsx->dimension($worksheetIndex):array [num_cols, num_rows]
$xlsx->sheetsCount():int
$xlsx->sheetNames():array
$xlsx->sheetName($worksheetIndex):string
$xlsx->sheetMeta($worksheetIndex = null):array sheets metadata (null = all sheets)
$xlsx->isHiddenSheet($worksheetIndex):bool
$xlsx->getStyles():array
```

## Examples
### XLSX to html table
```php
echo SimpleXLSX::parse('book.xlsx')->toHTML();
```
or
```php
if ( $xlsx = SimpleXLSX::parse('book.xlsx') ) {
	echo '<table border="1" cellpadding="3" style="border-collapse: collapse">';
	foreach( $xlsx->rows() as $r ) {
		echo '<tr><td>'.implode('</td><td>', $r ).'</td></tr>';
	}
	echo '</table>';
} else {
	echo SimpleXLSX::parseError();
}
```
or styled html table
```php
if ( $xlsx = SimpleXLSX::parse('book_styled.xlsx') ) {
    echo $xlsx->toHTMLEx();
}
```
### XLSX read huge file, xlsx to csv
```php
if ( $xlsx = SimpleXLSX::parse( 'xlsx/books.xlsx' ) ) {
    $f = fopen('book.csv', 'wb');
    // fwrite($f, chr(0xEF) . chr(0xBB) . chr(0xBF)); // UTF-8 BOM
    foreach ( $xlsx->readRows() as $r ) {
        fputcsv($f, $r); // fputcsv($f, $r, ';', '"', "\\", "\r\n");
    }
    fclose($f);
} else {
    echo SimpleXLSX::parseError();
}
```
### XLSX get sheet names and sheet indexes
```php
// Sheet numeration started 0

if ( $xlsx = SimpleXLSX::parse( 'xlsx/books.xlsx' ) ) {
    print_r( $xlsx->sheetNames() );
    print_r( $xlsx->sheetName( $xlsx->activeSheet ) );
}

```
```
Array
(
    [0] => Sheet1
    [1] => Sheet2
    [2] => Sheet3
)
Sheet2
```
### Using rowsEx() to extract cell info
```php
$xlsx = SimpleXLSX::parse('book.xlsx');
print_r( $xlsx->rowsEx() );

```
```
Array
(
    [0] => Array
        (
            [0] => Array
                (
                    [type] => s
                    [name] => A1
                    [value] => ISBN
                    [href] => 
                    [f] => 
                    [format] => 
                    [s] => 0
                    [css] => color: #000000;font-family: Calibri;font-size: 17px;
                    [r] => 1
                    [hidden] =>
                    [width] => 13.7109375
                    [height] => 0
                    [comment] =>
                )
        
            [1] => Array
                (
                    [type] => 
                    [name] => B1
                    [value] => 2016-04-12 13:41:00
                    [href] => Sheet1!A1
                    [f] => 
                    [format] => m/d/yy h:mm
                    [s] => 0
                    [css] => color: #000000;font-family: Calibri;font-size: 17px;            
                    [r] => 2
                    [hidden] => 1
                    [width] => 16.5703125
                    [height] => 0
                    [comment] => Serg: See transaction history   
                    
                )
```
<!--suppress HttpUrlsUsage -->
<table>
<tr><td>type</td><td>cell <a href="http://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_CellType_topic_ID0E6NEFB.html#topic_ID0E6NEFB">type</a></td></tr>
<tr><td>name</td><td>cell name (A1, B11)</td></tr>
<tr><td>value</td><td>cell value (1233, 1233.34, 2022-02-21 00:00:00, String)</td></tr>
<tr><td>href</td><td>internal and external links</td></tr>
<tr><td>f</td><td>formula</td></tr>
<tr><td>s</td><td>style index, use <code>$xlsx->cellFormats[ $index ]</code> to get style</td></tr>
<tr><td>css</td><td>generated cell CSS</td></tr>
<tr><td>r</td><td>row index</td></tr>
<tr><td>hidden</td><td>hidden row or column</td></tr>
<tr><td>width</td><td>width in <a href="http://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html">custom units</a></td></tr>
<tr><td>height</td><td>height in points (pt, 1/72 in)</td></tr>
<tr><td>comment</td><td>Cell comment as plain text</td></tr>
</table>

### Select Sheet
```php
$xlsx = SimpleXLSX::parse('book.xlsx');
// Sheet numeration started 0, we select second worksheet
foreach( $xlsx->rows(1) as $r ) {
// ...
}
```
### Get sheet by index 
```php
$xlsx = SimpleXLSX::parse('book.xlsx');	
echo 'Sheet Name 2 = '.$xlsx->sheetName(1);
```
### XLSX::parse remote data
```php
if ( $xlsx = SimpleXLSX::parse('https://www.example.com/example.xlsx' ) ) {
	$dim = $xlsx->dimension(1); // don't trust dimension extracted from xml
	$num_cols = $dim[0];
	$num_rows = $dim[1];
	echo $xlsx->sheetName(1).':'.$num_cols.'x'.$num_rows;
} else {
	echo SimpleXLSX::parseError();
}
```
### XLSX::parse memory data
```php
// For instance $data is a data from database or cache    
if ( $xlsx = SimpleXLSX::parseData( $data ) ) {
	print_r( $xlsx->rows() );
} else {
	echo SimpleXLSX::parseError();
}
```
### Get Cell (slow)
```php
echo $xlsx->getCell(0, 'B2'); // The Hobbit
``` 
### DateTime helpers
```php
// default SimpleXLSX datetime format is YYYY-MM-DD HH:MM:SS (ISO, MySQL)
echo $xlsx->getCell(0,'C2'); // 2016-04-12 13:41:00

// custom datetime format
$xlsx->setDateTimeFormat('d.m.Y H:i');
echo $xlsx->getCell(0,'C2'); // 12.04.2016 13:41

// unixstamp
$xlsx->setDateTimeFormat('U');
$ts = $xlsx->getCell(0,'C2'); // 1460468460
echo gmdate('Y-m-d', $ts); // 2016-04-12
echo gmdate('H:i:s', $ts); // 13:41:00

// raw excel value
$xlsx->setDateTimeFormat( NULL ); // returns as excel datetime
$xd = $xlsx->getCell(0,'C2'); // 42472.570138889
echo gmdate('m/d/Y', $xlsx->unixstamp( $xd )); // 04/12/2016
echo gmdate('H:i:s', $xlsx->unixstamp( $xd )); // 13:41:00 
```
### Rows with header values as keys
```php
if ( $xlsx = SimpleXLSX::parse('books.xlsx')) {
    // Produce array keys from the array values of 1st array element
    $header_values = $rows = [];
    foreach ( $xlsx->rows() as $k => $r ) {
        if ( $k === 0 ) {
            $header_values = $r;
            continue;
        }
        $rows[] = array_combine( $header_values, $r );
    }
    print_r( $rows );
}
```
```
Array
(
    [0] => Array
        (
            [ISBN] => 618260307
            [title] => The Hobbit
            [author] => J. R. R. Tolkien
            [publisher] => Houghton Mifflin
            [ctry] => USA
        )
    [1] => Array
        (
            [ISBN] => 908606664
            [title] => Slinky Malinki
            [author] => Lynley Dodd
            [publisher] => Mallinson Rendel
            [ctry] => NZ
        )
)
```
### Debug
```php
use Shuchkin\SimpleXLSX;

ini_set('error_reporting', E_ALL );
ini_set('display_errors', 1 );

if ( $xlsx = SimpleXLSX::parseFile('books.xlsx', true ) ) {
    echo $xlsx->toHTML();
} else {
    echo SimpleXLSX::parseError();
}
```
### Classic OOP style 
```php
use SimpleXLSX;

$xlsx = new SimpleXLSX('books.xlsx'); // try...catch
if ( $xlsx->success() ) {
    foreach( $xlsx->rows() as $r ) {
        // ...
    }
} else {
    echo 'xlsx error: '.$xlsx->error();
}
```
More examples [here](https://github.com/shuchkin/simplexlsx/tree/master/examples)

### Error Codes
SimpleXLSX::ParseErrno(), $xlsx->errno()<br/>
<table>
<tr><th>code</th><th>message</th><th>comment</th></tr>
<tr><td>1</td><td>File not found</td><td>Where file? UFO?</td></tr>
<tr><td>2</td><td>Unknown archive format</td><td>ZIP?</td></tr>
<tr><td>3</td><td>XML-entry parser error</td><td>bad XML</td></tr>
<tr><td>4</td><td>XML-entry not found</td><td>bad ZIP archive</td></tr>
<tr><td>5</td><td>Entry not found</td><td>File not found in ZIP archive</td></tr>
<tr><td>6</td><td>Worksheet not found</td><td>Not exists</td></tr>
</table>