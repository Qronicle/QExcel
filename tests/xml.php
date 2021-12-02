<?php

namespace Tests;

require_once '../vendor/autoload.php';

use QExcel\QExcel;

$workbook = QExcel::loadWorkbook('files/test.xml');
print_r($workbook->getSheet('Jommeke'));