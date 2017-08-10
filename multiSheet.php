<?php 	
require('Connect/con_db.php');
require('Classes/PHPExcel.php');
if(isset($_POST['btnsend'])) {
    $file       = $_FILES['file']['tmp_name']; // Biến này lưu tên tạm của file.
    // Tạo đối tượng Reader
    $objReader  = PHPExcel_IOFactory::createReaderForFile($file);
    // thông tin các name sheet
    $listWorkSheets = $objReader->listWorksheetNames($file);
    
    foreach ($listWorkSheets as $name ) {
        $sql = "INSERT INTO class(name) VALUES('$name')";
        $con->query($sql);
        $id_class = $con->insert_id; // id của sheet, tương đương id_class
        
        //Load 1 Sheet từ file excel
        $objReader->setLoadSheetsOnly($name);
        // nhận thông tin của sheet
        $objExcel   = $objReader->load($file);
        //chuyển thông tin của sheet sang mảng.
        $sheetdata  = $objExcel->getActiveSheet()->toArray('null', true, true, true);
        //print_r($sheetdata);
        //stt dòng cuối cùng trong bảng excel
        $highestRow = $objExcel->setActiveSheetIndex()->getHighestRow();
        //echo $highestRow; exit();
        //đưa thông tin của sheet vào database-đưa 1 sheet vào bảng core
        for ($i = 2; $i < $highestRow ; $i++) { 
            // thông tin lấy từ excel
            $name     = $sheetdata[$i]['A'];
            $toan     = $sheetdata[$i]['B'];
            $ly       = $sheetdata[$i]['C'];
            $hoa      = $sheetdata[$i]['D'];
            //insert vào database
            $sql = "INSERT INTO core(id_class, name, toan, ly, hoa) VALUES ($id_class, '$name', $toan, $ly, $hoa) " ;
            $con->query($sql);
        }
    }
    echo "OK";
    
}

 ?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Import data to Database</title>
    <!-- Latest compiled and minified CSS -->
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">

    <!-- Optional theme -->
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap-theme.min.css" integrity="sha384-rHyoN1iRsVXV4nD0JutlnGaslCJuC7uwjduW9SVrLvRYooPp2bWYgmgJQIXwl/Sp" crossorigin="anonymous">

    <!-- Latest compiled and minified JavaScript -->
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js" integrity="sha384-Tc5IQib027qvyjSMfHjOMaLkfuWVxZxUPnCJA7l2mCWNIpG9mGCD8wGNIcPD7Txa" crossorigin="anonymous"></script>
</head>
<body >
    <form action="" method="POST" role="form" style="width: 800px; height: 600px;text-align: center;" enctype="multipart/form-data">
        <legend>Import data from Excel to Database</legend>
    
        <div class="form-group">
            <label for=""></label>
            <input type="file" name="file" class="form-control" id="" placeholder="Input field">
        </div>
        <button type="submit" name="btnsend" class="btn btn-primary">Import</button>
    </form>
</body>
</html>