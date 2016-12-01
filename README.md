#QExcel

The Qronicle (or Quick - still haven't really decided) Excel library is an experimental Excel reader based on PHPExcel (http://phpexcel.codeplex.com/). It uses less resources by only looking at the actual content, ignoring styles etc.

This library should be helpful when importing excel files where the styling is not important. Keep in mind that it will only ever contain the Excel readers.

## Progress

Basic functionality is all in place. The optimized Excel5, Excel2003XML, Excel2007 and CSV Readers are mostly ready. Everything is tied together by the QExcel class (that can be used as a replacement to PHPExcel's IO factory).

Up next is adding more documentation, example files (the test files I use now are not open for publication) and extending the index file with more sweetness to test out. If everything works I'll probably add the OO and other Readers from PHPExcel. I also need to check for updates on the PHPExcel front, should they have improved their readers.

## Composer installation

```php
composer require qronicle/qexcel
```

## Getting started

```php
// Always include the QExcel file
// This will start the autoloader and you will probably use the QExcel class to start as well
require_once('library/QExcel/QExcel.php');

// The workbook file
$filename = 'files/test.xls';

// Load the file into a QExcel_Workbook object
$workbook = QExcel::loadWorkbook($filename);
```

## Performance

Tested with a file containing 2400 lines and 22 columns (mostly text and number fields).
PHPExcel used setReadDataOnly as true.

Note that these are quickly made tests and that they are just an indication of the amount of memory and time that is won by ignoring the (for this library) unimportant data.

### Excel2007 Reader

<table>
	<tr>
		<th></th>
		<th>PHPExcel</th>
		<th>QExcel</th>
		<th>Gain</th>
	</tr>
	<tr>
		<td>Memory usage</td>
		<td>158.76 MB</td>
		<td>8.87 MB</td>
		<td>5.6% (18 times less)</td>
	</tr>
	<tr>
		<td>Duration</td>
		<td>18.35 seconds</td>
		<td>3.01 seconds</td>
		<td>16.4% (6 times faster)</td>
	</tr>
</table>

### Excel5 Reader

<table>
	<tr>
		<th></th>
		<th>PHPExcel</th>
		<th>QExcel</th>
		<th>Gain</th>
	</tr>
	<tr>
		<td>Memory usage</td>
		<td>62.46 MB</td>
		<td>12.40 MB</td>
		<td>19.8% (5 times less)</td>
	</tr>
	<tr>
		<td>Duration</td>
		<td>7.77 seconds</td>
		<td>2.86 seconds</td>
		<td>36.8% (3 times faster)</td>
	</tr>
</table>

### Excel2003XML Reader

<table>
	<tr>
		<th></th>
		<th>PHPExcel</th>
		<th>QExcel</th>
		<th>Gain</th>
	</tr>
	<tr>
		<td>Memory usage</td>
		<td>172.29 MB</td>
		<td>6.08 MB</td>
		<td>3.5% (28 times less)</td>
	</tr>
	<tr>
		<td>Duration</td>
		<td>13.67 seconds</td>
		<td>2.66 seconds</td>
		<td>36.8% (5 times faster)</td>
	</tr>
</table>

### CSV Reader

<table>
	<tr>
		<th></th>
		<th>PHPExcel</th>
		<th>QExcel</th>
		<th>Gain</th>
	</tr>
	<tr>
		<td>Memory usage</td>
		<td>55.29 MB</td>
		<td>6.31 MB</td>
		<td>11.4% (9 times less)</td>
	</tr>
	<tr>
		<td>Duration</td>
		<td>6.80 seconds</td>
		<td>0.42 seconds</td>
		<td>6.1% (16 times faster)</td>
	</tr>
</table>
