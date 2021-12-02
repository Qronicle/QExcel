<?php

namespace Tests;

require_once '../vendor/autoload.php';

use QExcel\QExcel;

$workbook = QExcel::loadWorkbook('files/test.csv');
print_r($workbook->getSheet(0));