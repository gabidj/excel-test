<?php

namespace exceltest;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require('vendor/autoload.php');


// initially has a worksheet, will be the last one!
$spreadsheet = new Spreadsheet();
/*
$sheetA = new Worksheet(null, 'a');
$sheetB = new Worksheet(null, 'b');
$sheetC = new Worksheet(null, 'c');
$sheetD = new Worksheet(null, 'd');
$sheetE = new Worksheet(null, 'e');

$spreadsheet->addSheet($sheetA, 0);
$spreadsheet->addSheet($sheetB, 1);
$spreadsheet->addSheet($sheetC, 2);
$spreadsheet->addSheet($sheetD, 3);
$spreadsheet->addSheet($sheetE, 4);
//*/

$xlsConfig = [
    'header' => [
        'size' => [
            'width' => '10',
            'height' => '3'
        ],
        'start' => [
            'col' => 'A',
            'row' => '1'
        ],
        'text' => 'Header Test',
        'fill' => [
            'type' => 'solid',
            // ARGB
            'startColor' => 'FF004499',
            'endColor' => 'FF004499'
        ],
        'font' => [
            // argb
            'color' => 'FFFFFFFF',
        ],
    ]
];


function incrementLetter(string $letter, $times = 1) : string {
    for ($i=0; $i<$times; $i++) {
        $letter++;
    }
    return $letter;
}

$startCol = $xlsConfig['header']['start']['col'];
$startRow = $xlsConfig['header']['start']['row'];
$endCol = incrementLetter($startCol, $xlsConfig['header']['size']['width'] - 1);
$endRow = $startRow + $xlsConfig['header']['size']['height'] - 1;

$workSheet = $spreadsheet->getActiveSheet();

// header range
$headerRange = $startCol.$startRow.':'.$endCol.$endRow;

// set header text
$workSheet->setCellValue($startCol.$startRow, $xlsConfig['header']['text']);

# $cell = $workSheet->getCell('E2');
$cell = $workSheet->mergeCells($headerRange);;

// style get
$style = $cell->getStyle($startCol.$startRow);

// style apply
$style->getFont()->setColor(new Color($xlsConfig['header']['font']['color']))->setBold(true);
$style->getFill()->getStartColor()->setARGB($xlsConfig['header']['fill']['startColor']);
$style->getFill()->getEndColor()->setARGB($xlsConfig['header']['fill']['endColor']);

$style->getFill()->setFillType(Fill::FILL_SOLID);
$style->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
$style->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);


$style->getFont()->setSize('24');





// title
$workSheet->setTitle('Test Title');

$writer = new Xlsx($spreadsheet);
$writer->save('test/a.xlsx');
