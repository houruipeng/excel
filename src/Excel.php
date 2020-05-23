<?php
/**
 * The following code, none of which has BUG.
 *
 * @author: BD<565792893@qq.com>
 * @date: 2020/5/13 21:07
 */

namespace hrp;

class Excel{

	/**
	 * 读取excel 数据
	 * excel 格式需要保持一致 uid字段需要做处理(导出或在excel里面处理)
	 *
	 * @param string $path
	 * @param array  $keys
	 * @return array 返回excel的内容
	 * @throws \Exception
	 */
	public function readExcel(string $path, array $keys){
		$excelData = [];
		$objPHPExcel = $this->getExcelData($path);

		$sheet = $objPHPExcel->getSheet(0);
		$highestRow = $sheet->getHighestRow(); //总行数
		$highestColumn = $sheet->getHighestColumn(); //总列数

		if($highestRow <= 0) throw new \Exception('excel没有数据');

		//第一行不读数据 总行数加1读取最后一行

		for($row = 2; $row < $highestRow + 1; $row++){
			$rowData = $sheet->rangeToArray('A'.$row.':'.$highestColumn.$row, null, true, false);
			if($rowData[0] > 0){
				foreach($rowData as $key => $value){
					//array_pop($value);
					$excelData[] = array_combine($keys, array_filter($value));
				}
			}
		}
		return $excelData;
	}

	/**
	 * @param array  $excelHeader id=>'主键',name=>'昵称'
	 * @param array  $xlsData 需要导出的数据
	 * @param string $exportPath 保存的路径
	 * @param string $fileName
	 * @return string 导出的文件名
	 * @throws \PHPExcel_Exception
	 * @throws \PHPExcel_Writer_Exception
	 */
	public function exportToExcel(array $excelHeader, array $xlsData, string $exportPath = '.', $fileName = ''){
		error_reporting(E_ALL);
		ini_set('display_errors', true);
		ini_set('display_startup_errors', true);
		date_default_timezone_set('PRC');

		$objExcel = new \PHPExcel();
		// 设置文档信息，这个文档信息windows系统可以右键文件属性查看
		//$objWriter = \PHPExcel_IOFactory::createWriter($objExcel, 'Excel5');

		$len = count($excelHeader);
		$objActSheet = $objExcel->getActiveSheet();
		$letter = $this->getLetter($len);
		$keys = array_keys($excelHeader);
		$headers = array_values($excelHeader);
		//填充表头信息
		for($i = 0; $i < $len; $i++){
			$objActSheet->setCellValue("$letter[$i]1", "$headers[$i]");
		};

		//填充表格信息
		foreach($xlsData as $k => $v){
			$k += 2;
			for($i = 0; $i < $len; $i++){
				$objActSheet->setCellValue($letter[$i].$k, $v[$keys[$i]]);
			}
			// 表格高度
			$objActSheet->getRowDimension($k)->setRowHeight(20);
		}
		$fileName = $fileName ?? md5(date('YmdHis'));
		$outfile = $exportPath.DIRECTORY_SEPARATOR.$fileName.'.xlsx';

		(new \PHPExcel_Writer_Excel2007($objExcel))->save($outfile);
		chmod($outfile, 0755);
		return $outfile;
	}

	private function getLetter($len){
		$letter = [
			'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
			'V', 'W', 'X', 'Y', 'Z',
		];
		return array_slice($letter, 0, $len);
	}

	/**
	 * 读取excel 信息
	 *
	 * @param $path
	 * @return \PHPExcel
	 * @throws \PHPExcel_Reader_Exception
	 */
	private function getExcelData($path){
		if(!file_exists($path)) throw new \Exception('excel不存在');
		$inputFileType = \PHPExcel_IOFactory::identify($path);
		$objReader = \PHPExcel_IOFactory::createReader($inputFileType);
		return $objReader->load($path);
	}
}
