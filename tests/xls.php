<?php

namespace Tests;

require_once '../vendor/autoload.php';

use QExcel\QExcel;

$workbook = QExcel::loadWorkbook('files/test.xls');
print_r($workbook->getSheet('Jommeke'));