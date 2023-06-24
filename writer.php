<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$activeWorksheet = $spreadsheet->getActiveSheet();
$activeWorksheet->setCellValue('A1', 'Hello World !');

$spreadsheet->getActiveSheet()->setCellValue('A2', 'PhpSpreadsheet');
$spreadsheet->getActiveSheet()->setCellValue('A3', 12345.6789);
$spreadsheet->getActiveSheet()->setCellValue('A4', TRUE);
$spreadsheet->getActiveSheet()->setCellValue(
    'B2',
    '=IF(A4, CONCATENATE(A2, " ", A3), CONCATENATE(A2, " ", A3))'
);

$spreadsheet->getActiveSheet()->setCellValue(
    'A5',
    '=IF(A4, CONCATENATE(A2, " ", A3), CONCATENATE(A1, " ", A2))'
);
$spreadsheet->getActiveSheet()->getCell('A4')
    ->getStyle()->setQuotePrefix(true);


$dateTimeNow = time();
$excelDateValue = \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel( $dateTimeNow );

$spreadsheet->getActiveSheet()->setCellValue(
    'A6',
    $excelDateValue
);
$spreadsheet->getActiveSheet()->getStyle('A6')
    ->getNumberFormat()
    ->setFormatCode(
        \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_DATETIME
    );

$spreadsheet->getActiveSheet()->setCellValueExplicit(
    'A8',
    "01513789642",
    \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING
);

$spreadsheet->getActiveSheet()->setCellValue('A9', 1513789642);
$spreadsheet->getActiveSheet()->getStyle('A9')
    ->getNumberFormat()
    ->setFormatCode(
        '00000000000'
    );


$arrayData = [
    [NULL, 2010, 2011, 2012],
    ['Q1',   12,   15,   21],
    ['Q2',   56,   73,   86],
    ['Q3',   52,   61,   69],
    ['Q4',   30,   32,    0],
    ['Q1',   12,   15,   21],
    ['Q2',   56,   73,   86],
    ['Q3',   52,   61,   69],
    ['Q4',   30,   32,    0],
    ['Q1',   12,   15,   21],
    ['Q2',   56,   73,   86],
    ['Q3',   52,   61,   69],
    ['Q4',   30,   32,    0],
];
$spreadsheet->getActiveSheet()
    ->fromArray(
        $arrayData,  // The data to set
        NULL,        // Array values with this value will not be set
        'D1'         // Top left coordinate of the worksheet range where we want to set these values (default is A1)
    );

$writer = new Xlsx($spreadsheet);
$writer->save('docs/hello world.xlsx');