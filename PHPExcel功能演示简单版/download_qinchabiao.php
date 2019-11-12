<?php
/**
 * Created by IntelliJ IDEA.
 * User: Alien
 * Date: 2019/7/8
 * Time: 17:56
 */
error_reporting(E_ALL ^ E_NOTICE ^ E_WARNING);  //关闭错误提示

header("content-type:text/html;charset=utf-8");  //设置编码

require_once './Classes/PHPExcel.php';
require_once './Classes/PHPExcel/IOFactory.php';
require_once './Classes/PHPExcel/Reader/Excel5.php';
include './Classes/PHPExcel/Writer/Excel2007.php';




function Table01(){
    //连接数据库
    $db=new mysqli();
    $db->connect('localhost','root1','root','excel');
    //创建PHPExcel实例对象
    $phpexcel=new PHPExcel();
    $phpexcel->setActiveSheetIndex(0);
    $sheet=$phpexcel->getActiveSheet();
    //从数据库中读出数据
    $Query = "SELECT xiangmu,hangshu,mianji,beizhu FROM qinchabiao";
    $articles = mysqli_query($db, $Query);

    $arr=[
        'xiangmu'=>"项目",
        'hangshu'=>"行数",
        'mianji'=>"面积",
        'beizhu'=>"备注",
    ];
    array_unshift($articles,$arr);
    $currow=0;
    //将数据库中取出来的数据插入到表中
    foreach ($articles as $key=>$v){
        $currow=$key+1;
        $sheet->setCellValue('A'.$currow,$v['xiangmu'])
            ->setCellValue('B'.$currow,$v['hangshu'])
            ->setCellValue('C'.$currow,$v['mianji'])
            ->setCellValue('D'.$currow,$v['beizhu']);
    }
    $phpexcel->getActiveSheet()->getStyle('A1:D'.$currow)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
    //设置单元格背景色

    foreach ($articles as $key=> $v){
        $currow=$key+1;
        if($v['hangshu']<5 && $v['beizhu'] > 70)
        {
        $phpexcel->getActiveSheet()->getStyle('A'.$currow.':D'.$currow)->getFill()->applyFromArray(array(
        'type' => PHPExcel_Style_Fill::FILL_SOLID,
        'startcolor' => array('rgb' => 'FF0000')));
        }
    }

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');
    header('Content-Disposition: attachment;filename="资产负债表检测结果.xls"');
    header('Cache-Control: max-age=0');
    // If you're serving to IE 9, then the following may be needed
    header('Cache-Control: max-age=1');

    // If you're serving to IE over SSL, then the following may be needed
    header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
    header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
    header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
    header('Pragma: public'); // HTTP/1.0

    $objWriter = PHPExcel_IOFactory::createWriter($phpexcel, 'Excel5');
    $objWriter->save('php://output');
    echo "<br>Table01 函数实现!!!<br>";
    return;
}
Table01();






//从数据库读取输出,并在指定的列里面判断条件,并生成标红
function Table03()
{
    //连接数据库
    $db = new mysqli();
    $db->connect('localhost', 'root1', 'root', 'excel');
    //创建PHPExcel实例对象
    $phpexcel = new PHPExcel();
    $phpexcel->setActiveSheetIndex(0);
    $sheet = $phpexcel->getActiveSheet();
    //从数据库中读出数据
    $Query = "SELECT xh, name, haha,lala FROM student";
    $articles = mysqli_query($db, $Query);

    $arr = [
        'xh' => "姓名",
        'name' => "性别",
        'haha' => "年龄",
        'lala' => "手机号",
    ];
    array_unshift($articles, $arr);
    $currow = 0;
    //将数据库中取出来的数据插入到表中
    foreach ($articles as $key => $v) {
        $currow = $key + 1;
        $sheet->setCellValue('A' . $currow, $v['xh'])
            ->setCellValue('B' . $currow, $v['name'])
            ->setCellValue('C' . $currow, $v['haha'])
            ->setCellValue('D' . $currow, $v['lala']);
    }
    $phpexcel->getActiveSheet()->getStyle('A1:D' . $currow)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
    //设置单元格背景色
    foreach ($articles as $key => $v) {
        $currow = $key + 1;
        if ($v['haha'] > 30) {
            $phpexcel->getActiveSheet()->getStyle('A' . $currow . ':D' . $currow)->getFill()->applyFromArray(array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'startcolor' => array('rgb' => 'FF0000')));
        }
    }

    /*function cellColor($cells,$color){
        $phpexcel = new PHPExcel;
        $phpexcel->setActiveSheetIndex(0);

        $phpexcel->getActiveSheet()->getStyle($cells)->getFill()->applyFromArray(array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array(
                'rgb' => $color
            )
        ));
    }

    cellColor('B5', 'F28A8C');
    cellColor('G5', 'F28A8C');
    cellColor('A7:I7', 'F28A8C');
    cellColor('A17:I17', 'F28A8C');
    cellColor('A30:Z30', 'F28A8C');*/

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');
    header('Content-Disposition: attachment;filename="jiandan.xls"');
    header('Cache-Control: max-age=0');
    // If you're serving to IE 9, then the following may be needed
    header('Cache-Control: max-age=1');

    // If you're serving to IE over SSL, then the following may be needed
    header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
    header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
    header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
    header('Pragma: public'); // HTTP/1.0

    $objWriter = PHPExcel_IOFactory::createWriter($phpexcel, 'Excel5');
    $objWriter->save('php://output');
    echo "<br>Table01 函数实现!!!<br>";
    return;
}

// Table01();



$haha = 1398282821;
if (substr((string)($haha), -1) == '1') {
    echo "Hello kitty";
} else {
    print  "Are you kidding?";
}