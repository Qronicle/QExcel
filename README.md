#QExcel

The Qronicle (or Quick - whatever) Excel library is an experimental Excel reader based on PHPExcel (http://phpexcel.codeplex.com/). Hopefully it will use less resources by only looking at the actual content, ignoring styles etc.

This library should be helpful when importing excel files where the styling is not important. This library will only ever contain the Excel readers.

## Progress

Currently the Excel5, Excel2003XML, Excel2007 and CSV Readers are mostly ready. 
The IOFactory needs a rewrite and a more unified configuration system (per reader) should be built.

For the moment the indiviual readers seem to work well, but note I didn't test a lot of different files per reader.

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