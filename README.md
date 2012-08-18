#QExcel

Experimental extension of PHPExcel (http://phpexcel.codeplex.com/) that will hopefully use less resources by only looking at the actual content, ignoring styles etc.

## Progress

I currently stripped both the Excel2007 and Excel5 Readers. After CSV and Excel2003XML Readers I'll work on stripping the total package and putting the contents up on GitHub.

## Performance

Tested with a file containing 2400 lines and 22 columns (mostly text and number fields).
PHPExcel used setReadDataOnly as true.

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
		<td>36.8% (2.7 times faster)</td>
	</tr>
</table>