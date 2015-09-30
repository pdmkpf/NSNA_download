<?php
/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Asia/Seoul');
if (PHP_SAPI == 'cli')
	die('This example should only be run from a Web Browser');

/** Include PHPExcel */
require_once '../Classes/PHPExcel.php';

// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');
// If you're serving to IE over SSL, then the following may be needed
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0
header('Content-Disposition: attachment;filename="degree.xls"');

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
// Set document properties
$objPHPExcel->getProperties()->setCreator("NSNA")
							 ->setLastModifiedBy("NSNA")
							 ->setTitle("NewsSource XLS")
							 ->setSubject("NewsSource XLS")
							 ->setDescription("NewsSource XLS")
							 ->setKeywords("NewsSource XLS")
							 ->setCategory("NewsSource XLS");

function trim_str($val) {
	$ret = (string) $val;
	$ret = trim($ret);
	return $ret;
}

//$type = $_GET['type'];
$keyword = urlencode($_GET['keyword']);// 검색 키워드를 의미. 주의사항: 다중키워드 검색을 할 경우 키워드 토큰을 white space가 아닌 +기호로 바꿔주세요. 예) 박근혜+대통령
$begin = $_GET['begin']; //검색 시작점. 양식 YYYYMMDD
$end = $_GET['end'];
$period = $_GET['period'];
$providers= "";
if (isset($_GET['providers'])) 
	$providers = urlencode($_GET['providers']); //언론사 필터링

$unit = "1";

if (isset($_GET['unit'])) {
	$unit = $_GET['unit']; //출력 날짜 범위
}

//http://147.47.123.2:9999/NSNA_ExpertFrontEnd/JSP/DownloadExcel.jsp?keyword=%EA%B7%BC%ED%98%9C&begin=20130601&end=20130614&providers=%EA%B2%BD%EC%9D%B8%EC%9D%BC%EB%B3%B4&unit=7
//$url = "http://147.47.123.2:8080/NSNA_ExpertFrontEnd/JSP/DownloadExcel.jsp?keyword=".$keyword."&begin=".$begin."&end=".$end."&type=second&providers=".$providers."&unit=".$unit;
$url = "http://147.47.125.161:9999/NSNA_FrontEnd/JSP/DownloadExcel.jsp?keyword=".$keyword."&period=".$period."&begin=".$begin."&end=".$end."&type=second&providers=".$providers."&unit=".$unit;
$xml = simplexml_load_file($url);

foreach($xml -> UNIT  as $unit_child) {
        $active_sheet = $objPHPExcel->setActiveSheetIndex(0)->setTitle('Degree');
        $row_num = 1;

        foreach($unit_child -> DEGREE -> children() as $child) {
                $child_name = $child -> getName();
                $cellNum = 0;
                foreach($child as $key => $value) {
                        $active_sheet->setCellValueByColumnAndRow($cellNum, $row_num, trim_str($value));
                        $cellNum += 1;
                }
                $row_num += 1;

        }

        break;

}

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
exit;
?>
