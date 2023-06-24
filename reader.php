<?php

include 'vendor/autoload.php';

$sheetname = 'SalesOrders';

$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
$reader->setReadDataOnly(TRUE);
$reader->setLoadSheetsOnly($sheetname);
$spreadsheet = $reader->load("docs/SampleData.xlsx");

$worksheet = $spreadsheet->getActiveSheet();

$excelDataArray;
$count = 0;

foreach ($worksheet->getRowIterator() as $row) {
    $cellIterator = $row->getCellIterator();
    $cellIterator->setIterateOnlyExistingCells(FALSE);

    foreach ($cellIterator as $cell) {
        $excelDataArray[$count][] = $cell->getValue();
    }

    ++$count;
}

var_dump($excelDataArray);

