<?php

$con = mysqli_connect('localhost','root','','import_excel');

if(isset($_POST['submit'])){
	   $file = $_FILES['doc']['tmp_name'];

	   $ext = pathinfo($_FILES['doc']['name'],PATHINFO_EXTENSION);

	   if($ext == 'xlsx'){

	   require('PHPExcel/PHPExcel.php');
	   require('PHPExcel/PHPExcel/IOFactory.php');



	   $obj = PHPExcel_IOFactory::load($file);

	   foreach($obj->getWorksheetIterator() as $sheet){
	   	 $getHighestRow = $sheet->getHighestRow();

	   	 for ($i=0; $i <=$getHighestRow; $i++) { 
	   	 	   $name = $sheet->getCellByColumnAndRow(0,$i)->getValue();
	   	 	   $email = $sheet->getCellByColumnAndRow(1,$i)->getValue();
           if($name!= ''){
	   	 	   mysqli_query($con,"insert into user(name, email) values('$name','$email') ");
	   	 	 }
	   	 }

   }

 }else{
 	echo "Invaild file format.";
 }

}


?>
<!DOCTYPE html>
<html>
	<head>
		<meta charset="utf-8">
		<meta http-equiv="X-UA-Compatible" content="IE=edge">
		<title></title>
		<link rel="stylesheet" href="">
		<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous">
		<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
	</head>
	<body>
		<div class="container mt-2 p-5">
			<div class="row">
				<div class="col-sm-12 col-lg-6 offset-lg-3 col-md-6 offset-md-3">

					<form action="" method="post" enctype="multipart/form-data" class="p-4 shadow-sm">
						<div class="row">
							<div class="col-sm-12 col-lg-10 col-md-10">
								<input type="file" class="form-control shadow-none border " name="doc">
							</div>
							<div class="col-sm-12 col-lg-2 col-md-2">
								<input type="submit" class="btn btn-success" name="submit" value="upload">
							</div>
						</div>
					</form>

				</div>
			</div>
			
		</div>
		
	</body>
</html>