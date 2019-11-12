<html>
<body>
<?php
error_reporting(E_ALL^E_NOTICE^E_WARNING);  //关闭错误提示

header("content-type:text/html;charset=utf-8");         //设置编码

require_once './Classes/PHPExcel.php';
require_once './Classes/PHPExcel/IOFactory.php';
require_once './Classes/PHPExcel/Reader/Excel5.php';
include './Classes/PHPExcel/Writer/Excel2007.php';
//接下来的就是查出数据或者修改，增加

/////////////////////////////////////////////////////////////文件上传//////////////////////////////////////
// 允许上传的文件后缀
$allowedExts = array("xls","xlsx","csv");
$temp = explode(".", $_FILES["file"]["name"]);

//获取文件名前缀
$firstname = $temp[0];
echo $firstname;

echo $_FILES["file"]["size"];
$extension = end($temp);     // 获取文件后缀名
if ((($_FILES["file"]["type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")  ### xls支持
|| ($_FILES["file"]["type"] == "application/vnd.ms-excel"))  ### csv支持
&& ($_FILES["file"]["size"] < 2048000)   // 小于 2000 kb
&& in_array($extension, $allowedExts))
{
	if ($_FILES["file"]["error"] > 0)
	{
		echo "错误：: " . $_FILES["file"]["error"] . "<br>";
	}
	else
	{
		echo "上传文件名: " . $_FILES["file"]["name"] . "<br>";
		echo "文件类型: " . $_FILES["file"]["type"] . "<br>";
		echo "文件大小: " . ($_FILES["file"]["size"] / 1024) . " kB<br>";
		echo "文件临时存储的位置: " . $_FILES["file"]["tmp_name"] . "<br>";

		// 判断当期目录下的 upload 目录是否存在该文件
		// 如果没有 upload 目录，你需要创建它，upload 目录权限为 777
//		if (file_exists("upload/" . $_FILES["file"]["name"]))
		if (file_exists("./upload/" . $firstname.'.'.$extension))
		{
			echo $_FILES["file"]["name"] . " 文件已经存在。 ";
//            rename('upload/'.$_FILES["file"]["name"],'/upload/2135.jpg');
            rename('./upload/'.$_FILES["file"]["name"],'./upload/'.$firstname.'.'.$extension);
			echo $_FILES["file"]["name"] . " 文件已经存在。 ";
		}
		else
		{
			// 如果 upload 目录不存在该文件则将文件上传到 upload 目录下
            // move_uploaded_file($_FILES["file"]["tmp_name"], "upload/" . $_FILES["file"]["name"]);

            //将文件从缓存空间中复制到本地文档时转换支持中文命名的文件
            move_uploaded_file($_FILES["file"]["tmp_name"], iconv("UTF-8","gb2312","upload/" .$firstname.'.'.$extension));
//            move_uploaded_file($_FILES["file"]["tmp_name"], iconv("UTF-8",$_FILES["file"]["name"]));
			echo "文件存储在: " .$_FILES["file"]["tmp_name"];
			//rename($_FILES["file"]["name"],$firstname.$extension);
			echo "文件存储在: " . "upload/" . $_FILES["file"]["name"];
		}
	}
}
else
{
	//文件格式错误时候的跳转页面
	$url  =  "./error/wrong_format.html" ;
	echo " <script language = 'javascript'  
	type = 'text/javascript'> ";
	echo "window.location.href = '$url' ";
	echo " </script> ";

}

/////////////////////////////////////////////////////////将上传的文件导入数据库/////////////////////////////////

//连接数据库
$db=new mysqli();
$db->connect('localhost','root','root','excel');

// $db->query('set names utf8');

$dir = './upload/';

//$templateName = $_FILES["file"]["name"];
$templateName = iconv("UTF-8","gb2312",$_FILES["file"]["name"]);

//实例化Excel读取类
$objReader = new PHPExcel_Reader_Excel2007();


//此处将选择要上传的文件作为要插入数据库的文件
if(!$objReader->canRead($dir.$templateName))
{
    $objReader = new PHPExcel_Reader_Excel5();
    if(!$objReader->canRead($dir.$templateName)){
        echo '无法识别的Excel文件！';
        return false;
    }
}

$objPHPExcel=$objReader->load($dir.$templateName);
$sheet=$objPHPExcel->getSheet(0);//获取第一个工作表
$highestRow=$sheet->getHighestRow();//取得总行数
$highestColumn=$sheet->getHighestColumn(); //取得总列数


//循环读取excel文件,读取一条,插入一条
for($j=1;$j<=$highestRow;$j++){//从第一行开始读取数据
    $str='';
    for($k='A';$k<=$highestColumn;$k++){            //从A列读取数据
        //这种方法简单，但有不妥，以'\\'合并为数组，再分割\\为字段值插入到数据库,实测在excel中，如果某单元格的值包含了\\导入的数据会为空
        $str.=$objPHPExcel->getActiveSheet()->getCell("$k$j")->getValue().' ';//读取单元格

    }
    //explode:函数把字符串分割为数组。
    $strs=explode(" ",$str);

	//插入到数据库中的对应的表中
    $sql="INSERT INTO student(xh, name, haha,lala) VALUES (
	 '{$strs[0]}',
	 '{$strs[1]}',
	 '{$strs[2]}',
	 '{$strs[3]}'
	)";
    $db->query($sql);//这里执行的是插入数据库操作

}
// unlink($dir.$templateName); //删除excel文件

$ip=gethostbyname($_ENV['COMPUTERNAME']);
echo '<br>'.$ip;

echo '<br>行数为::\t'.$highestColumn.'<br>';
echo '列数为:::\t'.$highestRow.'<br>';




/////////////////////////////////////////////////////////从本地文件读取上传至数据库/////////////////////////////////

//从本地文件读取上传至数据库
function insertDB(){
    //连接数据库
    $db=new mysqli();
    $db->connect('localhost','root','root','excel');


    $dir = './';
    $templateName = 'staff.xlsx';
//实例化Excel读取类
    $objReader = new PHPExcel_Reader_Excel2007();
    if(!$objReader->canRead($dir.$templateName)){
        $objReader = new PHPExcel_Reader_Excel5();
        if(!$objReader->canRead($dir.$templateName)){
            echo '无法识别的Excel文件！';
            return false;
        }
    }

    $objPHPExcel=$objReader->load($dir.$templateName);
    $sheet=$objPHPExcel->getSheet(0);//获取第一个工作表
    $highestRow=$sheet->getHighestRow();//取得总行数
    $highestColumn=$sheet->getHighestColumn(); //取得总列数


//循环读取excel文件,读取一条,插入一条
    for($j=1;$j<=$highestRow;$j++){//从第一行开始读取数据
        $list = array();
        $str='';
        $list = [];
        for($k='A';$k<=$highestColumn;$k++){            //从A列读取数据
            //这种方法简单，但有不妥，以'\\'合并为数组，再分割\\为字段值插入到数据库,实测在excel中，如果某单元格的值包含了\\导入的数据会为空
            $str.=$objPHPExcel->getActiveSheet()->getCell("$k$j")->getValue().' ';//读取单元格
//            echo $str;
        }
        //explode:函数把字符串分割为数组。
        $strs=explode(" ",$str);

        // echo '+++'.$strs;
        $sql="INSERT INTO student(xh, name, haha,lala) VALUES (
	 '{$strs[0]}',
	 '{$strs[1]}',
	 '{$strs[2]}',
	 '{$strs[3]}'
	)";
        $db->query($sql);//这里执行的是插入数据库操作
    }
    echo "<br>插入成功!!!<br>";
}

//insertDB();

/////////////////////////////////////////////从数据库读取输出,并在指定的列里面判断条件,并生成标红////////////////////////////
//   if($v['haha']>30
//从数据库读取输出,并在指定的列里面判断条件,并生成标红
function Table01(){
    //连接数据库
    $db=new mysqli();
    $db->connect('localhost','root','root','excel');
    //创建PHPExcel实例对象
    $phpexcel=new PHPExcel();
    $phpexcel->setActiveSheetIndex(0);
    $sheet=$phpexcel->getActiveSheet();
    //从数据库中读出数据
    $Query = "SELECT xh, name, haha,lala FROM student";
    $articles = mysqli_query($db, $Query);

    $arr=[
        'xh'=>"姓名",
        'name'=>"性别",
        'haha'=>"年龄",
        'lala'=>"手机号",
    ];
    array_unshift($articles,$arr);
    $currow=0;
    //将数据库中取出来的数据插入到表中
    foreach ($articles as $key=>$v){
        $currow=$key+1;
        $sheet->setCellValue('A'.$currow,$v['xh'])
            ->setCellValue('B'.$currow,$v['name'])
            ->setCellValue('C'.$currow,$v['haha'])
            ->setCellValue('D'.$currow,$v['lala']);
    }
    $phpexcel->getActiveSheet()->getStyle('A1:D'.$currow)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
    //设置单元格背景色
    $currow=0;
    foreach ($articles as $key=> $v){
        $currow=$key+1;
        if($v['haha']>30)
        {
    $phpexcel->getActiveSheet()->getStyle('A'.$currow.':D'.$currow)->getFill()->applyFromArray(array(
        'type' => PHPExcel_Style_Fill::FILL_SOLID,
        'startcolor' => array('rgb' => 'FF0000')));
		}
}

/* function cellColor($cells,$color){
     $phpexcel = new PHPExcel;
     $phpexcel->setActiveSheetIndex(0);

     $phpexcel->getActiveSheet()->getStyle($cells)->getFill()->applyFromArray(array(
         'type' => PHPExcel_Style_Fill::FILL_SOLID,
         'startcolor' => array(
             'rgb' => $color
         )
     ));
 }*/

//        cellColor('B5', 'F28A8C');
//        cellColor('G5', 'F28A8C');
//        cellColor('A7:I7', 'F28A8C');
//        cellColor('A17:I17', 'F28A8C');
//        cellColor('A30:Z30', 'F28A8C');
    /*header('Content-Type: application/vnd.ms-excel');//设置文档类型
    header('Content-Disposition: attachment;filename="员工信表.xls"');//设置文件名
    header('Cache-Control: max-age=0');*/
	date_default_timezone_set(‘PRC’);
    $temp = date("Y-m-d-H-i",time());
	$phpwriter = new PHPExcel_Writer_Excel2007($phpexcel);
    $phpwriter->save('./output/Output_'.$temp.'.xls');
    echo "<br>Table01 函数实现!!!<br>";
    return;
}
Table01();

//////////////////////////////////////从数据库读取输出,并在指定的列里面判断条件,并输出符合条件的表格////////////////////////////

//    where haha>30 && haha<34 && name like '女'   $sql = "DELETE FROM student where 1";  $db->query($sql);
function Table02(){
    //连接数据库
    $db=new mysqli();
    $db->connect('localhost','root','root','excel');
    //创建PHPExcel实例对象
    $phpexcel=new PHPExcel();
    $phpexcel->setActiveSheetIndex(0);
    $sheet=$phpexcel->getActiveSheet();
    //根据限定条件从数据库中读出数据
    $Query = "SELECT xh, name, haha,lala FROM student where haha>30 && haha<34 && name like '女'";
//    $Query = "SELECT xh, name, haha,lala FROM student where substr((string)(haha),-1)=='1'";
    $articles = mysqli_query($db, $Query);

    $arr=[
        'xh'=>"姓名",
        'name'=>"性别",
        'haha'=>"年龄",
        'lala'=>"手机号",
    ];
    array_unshift($articles,$arr);
    $currow=0;
    foreach ($articles as $key=>$v){
        $currow=$key+1;
        $sheet->setCellValue('A'.$currow,$v['xh'])
            ->setCellValue('B'.$currow,$v['name'])
            ->setCellValue('C'.$currow,$v['haha'])
            ->setCellValue('D'.$currow,$v['lala']);
    }
    $phpexcel->getActiveSheet()->getStyle('A1:D'.$currow)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
    //设置单元格背景色
    $currow=0;
    foreach ($articles as $key=> $v){
        $currow=$key+1;
        if($v['haha']>30)
        {
            $phpexcel->getActiveSheet()->getStyle('A'.$currow.':D'.$currow)->getFill()->applyFromArray(array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'startcolor' => array('rgb' => 'FF0000')));
        }
    }

    $phpwriter = new PHPExcel_Writer_Excel2007($phpexcel);
    $phpwriter->save('./Table02.xlsx');

    //使用结束删除数据库
    $sql = "DELETE FROM student where 1";
//    $db->query($sql);
    echo "<br>暂时未删除成功<br>";
    return;
}
Table02();


//测试   substr((string)($haha),-1

$haha = 1398282821;
if(substr((string)($haha),-1)=='1')
{
    echo "Hello kitty";
}else{
    print  "Are you kidding?";
}
?>

<h1>
	点击下面蓝色图标下载检测结果
</h1>
<div id="container">
<!--	<img width="680" height="433" src="./error/wrong_format.jpg">-->
	<p><a href="http://localhost/Check/tt.php"><img width="66" height="66" src="./error/download.jpg"></a></p>
</div>
<?php
$a=$_REQUEST["a"];
if ($a=="a")
{
	echo "执行程序吧";
}
?>
</body>
</html>
