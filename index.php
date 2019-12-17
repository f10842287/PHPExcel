<?php
include "PHPExcel/Classes/PHPExcel.php";
require_once "PHPExcel/Classes/PHPExcel.php";
require_once "PHPExcel/Classes/PHPExcel/IOFactory.php";
$objPHPExcel = new PHPExcel();

$objPHPExcel->getActiveSheet()->setCellValue('A1','日期');
$objPHPExcel->getActiveSheet()->setCellValue('B1','時間');
$objPHPExcel->getActiveSheet()->setCellValue('C1','餐點');
$objPHPExcel->getActiveSheet()->setCellValue('D1','花費');
$objPHPExcel->getActiveSheet()->setCellValue('E1','總開銷');
$objPHPExcel->getActiveSheet()->setCellValue('F1','當天總花費');

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(20);

date_default_timezone_set('Asia/Taipei');
$data = ["早餐","午餐","晚餐"];

for($i = 1;$i<31;$i++){
	if(date("m",strtotime("+$i day")) == date("m",strtotime("+1 months"))){break;}
	for($j = 1;$j<=3;$j++){
		$objPHPExcel->getActiveSheet()->setCellValue('A'.(($i-1)*3+$j+1),date("Y-m-d",strtotime("+$i day")));
		$objPHPExcel->getActiveSheet()->setCellValue('B'.(($i-1)*3+$j+1),$data[$j-1]);
		if(($i == 1) && ($j == 1)){continue;}
		$objPHPExcel->getActiveSheet()->setCellValue('E'.(($i-1)*3+$j+1),"=SUM(D".(($i-1)*3+$j+1).",E".(($i-1)*3+$j).")");
	}
	$sum = ($i-1)*3+4;
	$objPHPExcel->getActiveSheet()->setCellValue('F'.$sum,"=SUM(D".($sum-2).",D".($sum-1).",D".($sum).")");
}

$current_month = date("m");
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007'); 
$objWriter->save($current_month.'月份-記帳.xlsx');
?>