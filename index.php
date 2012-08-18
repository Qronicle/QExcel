<?php

require_once('common.php');
require_once('library/QExcel/QExcel.php');

switch (3)
{
    case 1:
        $excelFile = WEB_DIR . '/test/files/example1.xlsx';
        $reader = new QExcel_Reader_Excel2007();
        break;

    case 2:
        $excelFile = WEB_DIR . '/test/files/example2.xls';
        $reader = new QExcel_Reader_Excel5();
        break;

    case 3:
        $excelFile = WEB_DIR . '/test/files/export.csv';
        $reader = new QExcel_Reader_CSV();
        break;

    case 0:
    default:
        $excelFile = WEB_DIR . '/test/files/example2.xml';
        $reader = new QExcel_Reader_Excel2003XML();
}



$workbook = $reader->load($excelFile);
tdump('Workbook', $workbook);