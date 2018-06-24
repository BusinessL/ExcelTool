<?php

require 'vendor/autoload.php';
ini_set('memory_limit', '128M');                    // 设置内存使用，防止多文件时挂掉。  也可以通过配置php.ini文件调整

use PhpOffice\PhpSpreadsheet\IOFactory;

// 多文件读取xlsx文件
$filePath = 'data/';
$fileNames = glob($filePath.'*.xlsx');

foreach($fileNames as $singleName){
    $spreadsheet = IOFactory::load($singleName);
    $sheetData = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);
    var_dump($sheetData);
}

