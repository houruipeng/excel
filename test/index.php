<?php
/**
 * The following code, none of which has BUG.
 *
 * @author: BD<565792893@qq.com>
 * @date: 2020/5/13 21:31
 */
require_once '../vendor/autoload.php';

use hrp\Excel;

$excel = new Excel();

//要导入的表头,必须和excel有效列数量一直
$header = ['name', 'sex', 'department', 'phone', 'card'];
$rows = $excel->readExcel('./demo.xlsx', $header);
print_r($rows);

//导出excel
$outData = [
	['id' => 1, 'name' => 'foo'],
	['id' => 12, 'name' => 'foo1'],
	['id' => 13, 'name' => 'foo2'],
	['id' => 14, 'name' => 'foo3'],
	['id' => 15, 'name' => 'foo4'],
];
$excelHeader = ['id' => '编号', 'name' => '名称'];
$exportDir = '.';
$fileName = '123';
$res = $excel->exportToExcel($excelHeader, $outData, $exportDir, $fileName);
// .\123.xlsx
print_r($res);

