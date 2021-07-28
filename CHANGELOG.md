# Changelog

## 0.8.25 (2021-07-28)

* ::rowsEx() - returns *hidden* flag, for hidden cells 

## 0.8.24 (2021-05-18)

* Extract internal links

## 0.8.23 (2021-04-20)

* x10 fastest getCell, thx [Jonowa](https://github.com/shuchkin/simplexlsx/issues/102)
* for xpath idea / all private methods protected now. 

## 0.8.22 (2021-04-20)

* fixed lost hash of hyperlinks 

## 0.8.21 (2021-01-11)

* libxml_disable_entity_loader and PHP 8, thx [iKlsR](https://github.com/shuchkin/simplexlsx/issues/96)

## 0.8.19 (2020-07-28)

* fixed empty shared strings xml

## 0.8.18 (2020-07-21)

* fixed hyperlinks

## 0.8.17 (2020-06-15)

* fixed version number, added relative pathes a/../b support in relations

## 0.8.16 (2020-06-14)

* fixed abs pathes in relations

## 0.8.15 (2020-04-28)

* fixed LibreOffice xml specificity, thx [stephengmatthews](https://github.com/shuchkin/simplexlsx/issues/77)

## 0.8.14 (2020-04-03)

* fixed Office for Mac relations

## 0.8.13 (2020-02-19)

* removed skipEmptyRows property (xml optimization always), added static parseFile & parseData

## 0.8.12 (2020-01-22)

* remove empty rows (opencalc)

## 0.8.11 (2020-01-20)

* changed formats source priority

## 0.8.10 (2019-11-07)

* skipEmptyRows improved

## 0.8.9 (2019-08-15)

* fixed release version

## 0.8.8 (2019-06-19)

* removed list( $x, $y ), added bool $xlsx->skipEmptyRows, $xlsx->parseFile( $filename ), $xlsx->parseData( $data ), release 0.8.8

## 0.8.7 (2019-04-18)

* empty rows fixed, release 0.8.7

## 0.8.6 (2019-04-16)

* 1900/1904 bug fixed

## 0.8.5 (2019-03-07)

* SimpleXLSX::ParseErrno(), $xlsx->errno() returns error code

## 0.8.4 (2019-02-14)

* detect datetime values, mb_string.func_overload=2 support .!. Bitrix

## 0.8.3 (2018-11-14)

* getCell - fixed empty cells and rows, safe now, but very slow

## 0.8.2 (2018-11-09)

* fix empty cells and rows in rows() and rowsEx(), added setDateTimeFormat( $see_php_date_func )

## 0.8.1

* rename simplexlsx.php to SimpleXLSX.php, rename parse_error to parseError fix _columnIndex, add ->toHTML(), GNU to MIT license

## 0.7.13 (2018-06-18)

* get sheet indexes bug fix

## 0.7.12 (2018-06-17)

* $worksheet_id to $worksheet_index, sheet numeration started 0

## 0.7.11 (2018-04-25)

* rowsEx(), added row index "r" to cell info

## 0.7.10 (2018-04-21)

* fixed getCell, returns NULL if not exits

## 0.7.9 (2018-03-17)

* fixed sheetNames() (namespaced or not namespaced attr)

## 0.7.8 (2018-01-15)

* remove namespace prefixes (hardcoded)

## 0.7.7 (2017-10-02)

* XML External Entity (XXE) Prevention (<!ENTITY xxe SYSTEM "file: ///etc/passwd" >]>)

## 0.7.6 (2017-09-26)

* if worksheet_id === 0 (default) then detect first sheet (for LibreOffice capabilities)

## 0.7.5 (2017-09-10)

* ->getCell() - fixed

## 0.7.4 (2017-08-22)

* ::parse_error() - to get last error in "static style"

## 0.7.3 (2017-08-14)

* ->_parse fixed relations reader, added ->getCell( sheet_id, address, format ) for direct cell reading

## 0.7.2 (2017-05-13)

* ::parse( $filename ) helper method

## 0.7.1 (2017-03-29)

* License added

## 0.6.11 (2016-07-27)

* fixed timestamp()

## 0.6.10 (2016-06-10)

* fixed search entries (UPPERCASE)

## 0.6.9 (2015-04-12)

* $xlsx->datetime_format to force dates out

## 0.6.8 (2013-10-13)

* fixed dimension() where 1 row only, fixed rowsEx() empty cells indexes (Daniel Stastka)

## 0.6.7 (2013-08-10)

* fixed unzip (mac), added $debug param to _constructor to display errors

## 0.6.6 (2013-06-03)

* +entryExists()

## 0.6.5 (2013-03-18)

* fixed sheetName()

## 0.6.4 (2013-03-13)

* rowsEx(), _parse(): fixed date column type & format detection

## 0.6.3 (2013-03-13)

* rowsEx(): fixed formulas, added date type 'd', added format 'format'
					dimension(): fixed empty sheet dimension
                    + sheetNames() - returns array( sheet_id => sheet_name, sheet_id2 => sheet_name2 ...)

## 0.6.2 (2012-10-04)

* fixed empty cells, rowsEx() returns type and formulas now

## 0.6.1 (2012-09-14)

* removed "raise exception" and fixed _unzip

## 0.6 (2012-09-13)

* success(), error(), __constructor( $filename, $is_data = false )

## 0.5.1 (2012-09-13)

* sheetName() fixed

## 0.5 (2012-09-12)

* sheetName()

## 0.4

* sheets(), sheetsCount(), unixstamp( $excelDateTime )

## 0.3

* fixed empty cells (Gonzo patch)