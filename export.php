<?php 
require('Connect/con_db.php');
require('Classes/PHPExcel.php');
if(isset($_POST['btnExport'])) {
    // khởi tạo 1 đối tượng excel
    $objExcel = new PHPExcel;
    // tạo sheet đầu tiên trong excel
    $objExcel->setActiveSheetIndex(0);
    $sheet = $objExcel->getActiveSheet()->setTitle('D14_TH01');
    // Găn data cho các cột và dòng trong sheet D14_TH02
    $rowCount =1;
    $sheet->setCellValue('A'.$rowCount , 'Họ Tên');
    $sheet->setCellValue('B'.$rowCount , 'Toán');
    $sheet->setCellValue('C'.$rowCount , 'Lý');
    $sheet->setCellValue('D'.$rowCount , 'Hóa');

    $result = $con->query("SELECT core.name,toan,ly,hoa FROM core INNER JOIN class ON class.id = core.id_class WHERE class.name = 'D14_TH01' ");
    while ($row = mysqli_fetch_array($result) ) {
        //print_r($row);
        $rowCount++;
        $sheet->setCellValue('A'.$rowCount , $row['name']);
        $sheet->setCellValue('B'.$rowCount , $row['toan']);
        $sheet->setCellValue('C'.$rowCount , $row['ly']);
        $sheet->setCellValue('D'.$rowCount , $row['hoa']);
    }
    $objWriter = new PHPExcel_Writer_Excel2007($objExcel);
    $file_name = "export.xlsx";
    //print_r($file_name); exit();
    $objWriter->save($file_name);
    // tra về file kiểu attachment
    header('Content-Disposition: attachment;filename="'.$file_name.'"');
    // trả về dữ liệu kiểu excel và có đuôi .xlsx 
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Length:'.filesize($file_name)); // trả về file size
    header("Content-Transfer-Encoding: binary");
    header('Cache-Control: must-revalidate');
    header("Pragma: no-cache");
    readfile($file_name); // đọc file
    return ;

}


 ?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Export Data</title>
</head>
<body>
    <form action="" method="POST" role="form">
        <button type="submit" name="btnExport" class="btn btn-primary">Export</button>
    </form>
</body>
</html>