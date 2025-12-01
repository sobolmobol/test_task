<?php

if (!file_exists(__DIR__ . '/vendor/autoload.php')) {
    die("Ошибка: библиотека PhpSpreadsheet не установлена.");
}
require __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

foreach(range('A', 'J') as $columnId){
    foreach(range(1, 10) as $rawId){
        $cellId = $columnId . $rawId;
        $randomNumber = rand(1, 100);

        $sheet->setCellValue($cellId, $randomNumber);
    }
    //$sheet->getColumnDimension($columnId)->setAutoSize(true);
}
$sheet->getStyle('A1:J10')->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
$sheet->getStyle('A1:J10')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
$sheet->getStyle('A1:J10')->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);

$resultDir = __DIR__ . '/result';
if (!is_dir($resultDir)) {
    mkdir($resultDir, 0777, true);
}

$fileName = realpath($resultDir . '/random_numbers.xlsx');

try {
    $writer = new Xlsx($spreadsheet);
    $writer->save($fileName);
    echo "Файл " . $fileName . " успешно сохранён." . PHP_EOL;
} catch (Exception $e) {
    die("Ошибка при сохранении файла: " . $e->getMessage());
}
