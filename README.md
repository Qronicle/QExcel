QExcel
======

Experimental extension of PHPExcel (http://phpexcel.codeplex.com/) that will hopefully use less resources by only looking at the actual content.

Progress
--------

I currently stripped both the Excel2007 and Excel5 Readers. After CSV and Excel2003XML I'll work on stripping the total package and putting the contents up on GitHub.

Performance
-----------

Tested with a file with 2400 lines and 22 columns with mostly text and number fields. PHPExcel used setReadDataOnly as true.

*Excel2007 Reader*

	PHPExcel	QExcel
Memory usage	158.76 MB	8.87 MB	5.6% (18 times less)
Total time	18.35 seconds	3.01 seconds	16.4% (6 times faster)

*Excel5 Reader*

	PHPExcel	QExcel
Memory usage	62.46 MB	12.40 MB	19.8% (5 times less)
Total time	7.77 seconds	2.86 seconds	36.8% (2.7 times faster)