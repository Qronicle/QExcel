<?php

require_once('common.php');
require_once('library/QExcel/QExcel.php');

switch (1)
{
    case 1:
        $excelFile = WEB_DIR . '/test/files/example1.xlsx';
        $readerType = 'Excel2007';
        break;

    case 2:
        $excelFile = WEB_DIR . '/test/files/example2.xls';
        $readerType = 'Excel5';
        break;

    case 3:
        $excelFile = WEB_DIR . '/test/files/export.csv';
        $readerType = 'CSV';
        break;

    case 0:
    default:
        $excelFile = WEB_DIR . '/test/files/example2.xml';
        $readerType = 'Excel2003XML';
}

$workbook = QExcel::loadWorkbook($excelFile/*, $readerType*/);

tdump('Workbook', $workbook);