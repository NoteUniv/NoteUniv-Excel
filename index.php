<?php

require 'vendor/autoload.php';

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
$spreadsheet = $reader->load('grades_template.xlsx');

// Get sheet Select values
$sheet = $spreadsheet->getSheetByName('Select values');

// Select data from row 2
$data = $sheet->rangeToArray('B2:N2')[0];
$alphas = range('B', 'N');

$promo = 'LPDWEB';
$promoCell = $alphas[array_search($promo, $data)];
$subjects = ['Math', 'English', 'Science', 'History', 'Geography', 'Biology', 'Chemistry', 'Physics', 'Music', 'Art', 'PE', 'Health', 'Sports', 'Other'];

// Insert subject into row B starting row 3
$row = 3;
foreach ($subjects as $subject) {
    $sheet->setCellValue($promoCell . $row, $subject);
    $row++;
}

$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
$writer->save('grades_template_full.xlsx');
