<?php

require 'vendor/autoload.php';
ini_set('memory_limit', '128M');                    // 设置内存使用，防止多文件时挂掉。  也可以通过配置php.ini文件调整
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;

// 可多文件读取
$filePath = 'data/';
$fileNames = glob($filePath.'*.json');
foreach($fileNames as $singleName){
    // 去除后缀 
    $excelName = basename($singleName,".json");

    $contents = file_get_contents($singleName);
    // $contents = iconv('gbk','utf-8',$contents);    gbk 转 utf-8

    $contents = json_decode($contents);
    // echo json_last_error();               // 检测json格式是否正确

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setCellValue('A1', '大学名称');
    $sheet->setCellValue('B1', '录取人数');
    $sheet->setCellValue('C1', '备注');
    $sheet->setCellValue('D1', '专业名称');
        
    // Set the number format mask so that the excel timestamp will be displayed as a human-readable date/time
    $spreadsheet->getActiveSheet()->getStyle('A6')->getNumberFormat()->setFormatCode(
        \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_DATETIME
    );

    // Set cell A8 with a numeric value, but tell PhpSpreadsheet it should be treated as a string
    // $spreadsheet->getActiveSheet()->setCellValueExplicit('A8',"01513789642",\PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);

    $spreadsheet->getActiveSheet()->getStyle('D')->getAlignment()->setWrapText(true);  // 设置行内文字换行

    foreach($contents as $key=>$lineData){
        $codekey = $key+2;
        $spreadsheet->getActiveSheet()->setCellValue("A$codekey", $lineData->name);
        $spreadsheet->getActiveSheet()->setCellValue("B$codekey", $lineData->qty);
        $spreadsheet->getActiveSheet()->setCellValue("C$codekey", $lineData->info);
        
        $majors = '';
        foreach($lineData->majors as $SingleMajor){
            $majors = $SingleMajor->name."\n".$SingleMajor->qty."\n". $SingleMajor->info;
        }
        $spreadsheet->getActiveSheet()->setCellValue("D$codekey", $majors);
    }

    // csv文件写入
    $CsvWriter = IOFactory::createWriter($spreadsheet, 'Csv');
    $CsvWriter->setDelimiter(';');                         // 每个元素之间的分割符
    $CsvWriter->setEnclosure('');                          // 元素的设置格式   如:ex   "ex"
    $CsvWriter->setLineEnding("\r\n");                     // 换行格式
    $CsvWriter->setSheetIndex(0);                          // 第几行开始执行
    $CsvWriter->setPreCalculateFormulas(false);            // 表格中的公式是否执行
    // $writer->setUseBOM(true);                        // 是否增加BOM;默认是无BOM的，可通过设置增加
    $CsvWriter->save("$excelName.csv");

    // xlsx文件写入
    $XlsxWriter = IOFactory::createWriter($spreadsheet, "Xlsx");
    $XlsxWriter->save("$excelName.xlsx");
}

