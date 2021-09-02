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

#### rows()

```php
if ( $xlsx = SimpleXLSX::parse('book.xlsx') ) {
	print_r( $xlsx->rows() );
} else {
	echo SimpleXLSX::parseError();
}
```
***Result***
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

---

## rowsEx()

```php
if ( $xlsx = SimpleXLSX::parse('books.xlsx')) {
	// ->rowsEx();
	echo '<h2>$xlsx->rowsEx()</h2>';
	echo '<pre>';
	print_r( $xlsx->rowsEx() );
	echo '</pre>';

} else {
	echo SimpleXLSX::parseError();
}

```
***Result***
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
                    [r] => 1
                    [hidden] => 
                )

            [1] => Array
                (
                    [type] => s
                    [name] => B1
                    [value] => title
                    [href] => 
                    [f] => 
                    [format] => 
                    [r] => 1
                    [hidden] => 
                )

            [2] => Array
                (
                    [type] => s
                    [name] => C1
                    [value] => author
                    [href] => 
                    [f] => 
                    [format] => 
                    [r] => 1
                    [hidden] => 
                )

            [3] => Array
                (
                    [type] => s
                    [name] => D1
                    [value] => publisher
                    [href] => 
                    [f] => 
                    [format] => 
                    [r] => 1
                    [hidden] => 
                )

            [4] => Array
                (
                    [type] => s
                    [name] => E1
                    [value] => ctry
                    [href] => 
                    [f] => 
                    [format] => 
                    [r] => 1
                    [hidden] => 
                )

        )

    [1] => Array
        (
            [0] => Array
                (
                    [type] => 
                    [name] => A2
                    [value] => 618260307
                    [href] => 
                    [f] => 
                    [format] => 
                    [r] => 2
                    [hidden] => 
                )

            [1] => Array
                (
                    [type] => s
                    [name] => B2
                    [value] => The Hobbit
                    [href] => 
                    [f] => 
                    [format] => 
                    [r] => 2
                    [hidden] => 
                )

            [2] => Array
                (
                    [type] => s
                    [name] => C2
                    [value] => J. R. R. Tolkien
                    [href] => 
                    [f] => 
                    [format] => 
                    [r] => 2
                    [hidden] => 
                )

            [3] => Array
                (
                    [type] => s
                    [name] => D2
                    [value] => Houghton Mifflin
                    [href] => 
                    [f] => 
                    [format] => 
                    [r] => 2
                    [hidden] => 
                )

            [4] => Array
                (
                    [type] => s
                    [name] => E2
                    [value] => USA
                    [href] => 
                    [f] => 
                    [format] => 
                    [r] => 2
                    [hidden] => 
                )

        )

    [2] => Array
        (
            [0] => Array
                (
                    [type] => 
                    [name] => A3
                    [value] => 908606664
                    [href] => 
                    [f] => 
                    [format] => 
                    [r] => 3
                    [hidden] => 
                )

            [1] => Array
                (
                    [type] => s
                    [name] => B3
                    [value] => Slinky Malinki
                    [href] => 
                    [f] => 
                    [format] => 
                    [r] => 3
                    [hidden] => 
                )

            [2] => Array
                (
                    [type] => s
                    [name] => C3
                    [value] => Lynley Dodd
                    [href] => 
                    [f] => 
                    [format] => 
                    [r] => 3
                    [hidden] => 
                )

            [3] => Array
                (
                    [type] => s
                    [name] => D3
                    [value] => Mallinson Rendel
                    [href] => 
                    [f] => 
                    [format] => 
                    [r] => 3
                    [hidden] => 
                )

            [4] => Array
                (
                    [type] => s
                    [name] => E3
                    [value] => NZ
                    [href] => 
                    [f] => 
                    [format] => 
                    [r] => 3
                    [hidden] => 
                )

        )

)
```
---

#### cRows()

```php
if ( $xlsx = SimpleXLSX::parse('book.xlsx') ) {
	print_r( $xlsx->cRows() );
} else {
	echo SimpleXLSX::parseError();
}
```
***Result***
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

OR

```php

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
```
***Result***

```
Array
(
    [0] => Array
        (
            [isbn] => 618260307
            [title] => The Hobbit
            [author] => J. R. R. Tolkien
            [publisher] => Houghton Mifflin
            [ctry] => USA
        )

    [1] => Array
        (
            [isbn] => 908606664
            [title] => Slinky Malinki
            [author] => Lynley Dodd
            [publisher] => Mallinson Rendel
            [ctry] => NZ
        )

)
```

---

## cRowEx()

```php
if ( $xlsx = SimpleXLSX::parse('books.xlsx')) {

	echo '<pre>';

	echo '<h2>$xlsx->cRowsEx()</h2>';
	echo '<pre>';
	print_r( $xlsx->cRowsEx() );
	echo '</pre>';

} else {
	echo SimpleXLSX::parseError();
}
```
***Result***
```
Array
(
    [title] => Array
        (
            [0] => ISBN
            [1] => title
            [2] => author
            [3] => publisher
            [4] => ctry
        )

    [data] => Array
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

    [details] => Array
        (
            [0] => Array
                (
                    [ISBN] => Array
                        (
                            [type] => 
                            [name] => A2
                            [value] => 618260307
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 2
                            [hidden] => 
                        )

                    [title] => Array
                        (
                            [type] => s
                            [name] => B2
                            [value] => The Hobbit
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 2
                            [hidden] => 
                        )

                    [author] => Array
                        (
                            [type] => s
                            [name] => C2
                            [value] => J. R. R. Tolkien
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 2
                            [hidden] => 
                        )

                    [publisher] => Array
                        (
                            [type] => s
                            [name] => D2
                            [value] => Houghton Mifflin
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 2
                            [hidden] => 
                        )

                    [ctry] => Array
                        (
                            [type] => s
                            [name] => E2
                            [value] => USA
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 2
                            [hidden] => 
                        )

                )

            [1] => Array
                (
                    [ISBN] => Array
                        (
                            [type] => 
                            [name] => A3
                            [value] => 908606664
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 3
                            [hidden] => 
                        )

                    [title] => Array
                        (
                            [type] => s
                            [name] => B3
                            [value] => Slinky Malinki
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 3
                            [hidden] => 
                        )

                    [author] => Array
                        (
                            [type] => s
                            [name] => C3
                            [value] => Lynley Dodd
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 3
                            [hidden] => 
                        )

                    [publisher] => Array
                        (
                            [type] => s
                            [name] => D3
                            [value] => Mallinson Rendel
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 3
                            [hidden] => 
                        )

                    [ctry] => Array
                        (
                            [type] => s
                            [name] => E3
                            [value] => NZ
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 3
                            [hidden] => 
                        )

                )

        )

)
```

OR

```php
	$columns = array(
		'isbn',
		'title',
		'author',
		'publisher',
		'ctry'
	);
    if ( $xlsx = SimpleXLSX::parse('books.xlsx')) {

	echo '<pre>';

	echo '<h2>$xlsx->cRowsEx($column)</h2>';
	echo '<pre>';
	print_r( $xlsx->cRowsEx($column) );
	echo '</pre>';

} else {
	echo SimpleXLSX::parseError();
}
```
***Result***
```
Array
(
    [title] => Array
        (
            [0] => ISBN
            [1] => title
            [2] => author
            [3] => publisher
            [4] => ctry
        )

    [data] => Array
        (
            [0] => Array
                (
                    [isbn] => 618260307
                    [title] => The Hobbit
                    [author] => J. R. R. Tolkien
                    [publisher] => Houghton Mifflin
                    [ctry] => USA
                )

            [1] => Array
                (
                    [isbn] => 908606664
                    [title] => Slinky Malinki
                    [author] => Lynley Dodd
                    [publisher] => Mallinson Rendel
                    [ctry] => NZ
                )

        )

    [details] => Array
        (
            [0] => Array
                (
                    [isbn] => Array
                        (
                            [type] => 
                            [name] => A2
                            [value] => 618260307
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 2
                            [hidden] => 
                        )

                    [title] => Array
                        (
                            [type] => s
                            [name] => B2
                            [value] => The Hobbit
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 2
                            [hidden] => 
                        )

                    [author] => Array
                        (
                            [type] => s
                            [name] => C2
                            [value] => J. R. R. Tolkien
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 2
                            [hidden] => 
                        )

                    [publisher] => Array
                        (
                            [type] => s
                            [name] => D2
                            [value] => Houghton Mifflin
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 2
                            [hidden] => 
                        )

                    [ctry] => Array
                        (
                            [type] => s
                            [name] => E2
                            [value] => USA
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 2
                            [hidden] => 
                        )

                )

            [1] => Array
                (
                    [isbn] => Array
                        (
                            [type] => 
                            [name] => A3
                            [value] => 908606664
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 3
                            [hidden] => 
                        )

                    [title] => Array
                        (
                            [type] => s
                            [name] => B3
                            [value] => Slinky Malinki
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 3
                            [hidden] => 
                        )

                    [author] => Array
                        (
                            [type] => s
                            [name] => C3
                            [value] => Lynley Dodd
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 3
                            [hidden] => 
                        )

                    [publisher] => Array
                        (
                            [type] => s
                            [name] => D3
                            [value] => Mallinson Rendel
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 3
                            [hidden] => 
                        )

                    [ctry] => Array
                        (
                            [type] => s
                            [name] => E3
                            [value] => NZ
                            [href] => 
                            [f] => 
                            [format] => 
                            [r] => 3
                            [hidden] => 
                        )

                )

        )

)
```

```
// SimpleXLSX::parse( $filename, $is_data = false, $debug = false ): SimpleXLSX (or false)
// SimpleXLSX::parseFile( $filename, $debug = false ): SimpleXLSX (or false)
// SimpleXLSX::parseData( $data, $debug = false ): SimpleXLSX (or false)
```

## Installation
The recommended way to install this library is [through Composer](https://getcomposer.org).
[New to Composer?](https://getcomposer.org/doc/00-intro.md)

This will install the latest supported version:
```bash
$ composer require shuchkin/simplexlsx
```
or download class [here](https://github.com/shuchkin/simplexlsx/blob/master/src/SimpleXLSX.php)

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
### XLSX read cells, out commas and bold headers
```php
echo '<pre>';
if ( $xlsx = SimpleXLSX::parse( 'xlsx/books.xlsx' ) ) {
	foreach ( $xlsx->rows() as $r => $row ) {
		foreach ( $row as $c => $cell ) {
			echo ($c > 0) ? ', ' : '';
			echo ( $r === 0 ) ? '<b>'.$cell.'</b>' : $cell;
		}
		echo '<br/>';
	}
} else {
	echo SimpleXLSX::parseError();
}
echo '</pre>';
```
### XLSX get sheet names and sheet indexes
```php
if ( $xlsx = SimpleXLSX::parse( 'xlsx/books.xlsx' ) ) {
	print_r( $xlsx->sheetNames() );
}
// Sheet numeration started 0
```
```
Array
(
    [0] => Sheet1
    [1] => Sheet2
    [2] => Sheet3
)
```
### Gets extend cell info by ->rowsEx()
```php
print_r( SimpleXLSX::parse('book.xlsx')->rowsEx() );
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
                    [r] => 1
                    [hidden] =>
                )

            [1] => Array
                (
                    [type] => 
                    [name] => B1
                    [value] => 2016-04-12 13:41:00
                    [href] => 
                    [f] => 
                    [format] => m/d/yy h:mm
                    [r] => 2
                    [hidden] => 1
                )
```
### Select Sheet
```php
$xlsx = SimpleXLSX::parse('book.xlsx');
print_r( $xlsx->rows(1) ); // Sheet numeration started 0, we select second worksheet
```
### Get sheet by index 
```php
$xlsx = SimpleXLSX::parse('book.xlsx');	
echo 'Sheet Name 2 = '.$xlsx->sheetName(1);
```
### XLSX::parse remote data
```php
if ( $xlsx = SimpleXLSX::parse('http://www.example.com/example.xlsx' ) ) {
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
// default SimpleXLSX datetime format YYYY-MM-DD HH:MM:SS (MySQL)
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
$xlsx = new SimpleXLSX('books.xlsx'); // try...catch
if ( $xlsx->success() ) {
	print_r( $xlsx->rows() );
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