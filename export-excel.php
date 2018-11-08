<?php
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="ten_file_se_duoc_tai_ve.xlsx"');
header('Cache-Control: max-age=0');
  //Include thư viện PHPExcel_IOFactory vào
include 'PHPExcel-1.8/Classes/PHPExcel/IOFactory.php';
include'conn.php';
mysqli_set_charset($conn, 'UTF8');
// Loại file cần ghi là file excel phiên bản 2007 trở đi
$fileType = 'Excel2007';
// Tên file cần ghi
$fileName = 'product_import.xlsx';
 
// Load file product_import.xlsx lên để tiến hành ghi file
$objPHPExcel = PHPExcel_IOFactory::load("product_import.xlsx");
 
// Giả sử chúng ta có mảng dữ liệu cần ghi như sau
$query = mysqli_query($conn,"SELECT * FROM customer_detail");

$array_data=array();
$k=0;
while($row = mysqli_fetch_array($query))
{
	
	$array_data[$k]=array('id'=>$row['ID_Customer'],'name'=>$row['Customer_Name'],'mobile'=>$row['Customer_Mobile']);
	$k++;
}

 
// Thiết lập tên các cột dữ liệu
$objPHPExcel->setActiveSheetIndex(0)
                            ->setCellValue('A1', "ID")
                            ->setCellValue('B1', "Name")
                            ->setCellValue('C1', "MOBILE");
                            
 
// Lặp qua các dòng dữ liệu trong mảng $array_data và tiến hành ghi dữ liệu vào file excel
$i = 2;
foreach ($array_data as $value) {
	
	$objPHPExcel->setActiveSheetIndex(0)
								->setCellValue("A$i", $value['id'])
								->setCellValue("B$i", $value['name'])
	                            ->setCellValue("C$i", $value['mobile']);
	                           
	$i++;
}
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, $fileType);
 //Tiến hành ghi file
$objWriter->save('php://output');
