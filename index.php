<?php

namespace exceltest;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require('vendor/autoload.php');



$spreadsheet = new Spreadsheet();
$spreadsheet->addSheet(new Worksheet(null, 'a'), 0);
$spreadsheet->addSheet(new Worksheet(null, 'b'), 1);
$spreadsheet->addSheet(new Worksheet(null, 'c'), 2);
$spreadsheet->addSheet(new Worksheet(null, 'd'), 3);

$workSheet = $spreadsheet->getActiveSheet();

$workSheet->setCellValue('A1', 'test');

// title
$workSheet->setTitle('Test Title');

$writer = new Xlsx($spreadsheet);
$writer->save('test/a.xlsx');
