<?php
/**
 * Created by PhpStorm.
 * User: yifeng
 * Date: 2017/8/10
 * Time: 22:57
 */

namespace app\index\controller;
use think\db\Query;

class UserSecond extends Common
{
    //员工列表
    public function index()
    {
        $admin_name = db('manager')->where('username',session('username'))->find();
        $Unique = $admin_name['Id'];

        $result = db('qinchabiao')->where('uid',$Unique)->order('Id Desc')->paginate(15);
        $this->assign('user', $result);
        return view();
    }

//    手动添加员工
    public function add()
    {
        if (request()->isPost()) {
            $data = input('post.');
            $validate = validate("User");
            if (!$validate->scene("add")->check($data)) {
                $this->error($validate->getError());
            }
            $data['password'] = md5('12345678');
            db('user')->insert($data);
            $this->success("员工添加成功", 'user/index');
            return;
        }
        return view();
    }

    //修改员工信息
    public function edit()
    {
        if (request()->isPost()) {
            $data = input("post.");
            $validate = validate("User");
            if (!$validate->scene('edit')->check($data)) {
                $this->error($validate->getError());
            }
            if (db('user')->update($data)) {
                $this->success("更新成功", 'user/index');
            } else {
                $this->error("数据更新不成功", null, null, 1);
            }

            return;
        }
        $id = input("id");
        $res = db('user')->where("Id", $id)->find();
        if (!$res) {
            $this->error("该员工不存在", "user/index");
        }
        $this->assign('user', $res);
        return view();
    }

//    删除
    public function del()
    {
        $id = input('id');
        if (!db('user')->delete($id)) {
            $this->error('删除失败');
        }
        $this->success("删除成功", 'user/index');
        return;
    }

    //批量导入员工信息
    public function leadingin()
    {
        if (request()->isPost()) {
            // 获取表单上传文件 例如上传了001.jpg
            $file = request()->file('fileexcel');
            $fileinfo = $this->upload($file);
            if (!$fileinfo['code']) {
                return json($fileinfo['info']);
            }
            //文档处理
            $inputFileType = \PHPExcel_IOFactory::identify($fileinfo['info']);
            $objReader = \PHPExcel_IOFactory::createReader($inputFileType);
            $objPHPExcel = $objReader->load($fileinfo['info']);
            $sheet = $objPHPExcel->getSheet(0);
            $data = $this->getexceldateone($sheet);
            if (!db('qinchabiao')->insertAll($data)) {
                return json("数据导入异常");
            }
            if (file_exists($fileinfo['info'])) {
                unlink($fileinfo['info']);
            }
            return json("数据导入成功");
        }
        return view();
    }

    //  取得excel数据方法一

    protected function getexceldateone($sheet)
    {
        $admin_name = db('manager')->where('username',session('username'))->find();

        //获取当前工作表的行数
        $rows = $sheet->getHighestRow();
        //获取当前工作表的列（在这里获取到的是字母列），
        $column = $sheet->getHighestColumn();
        $columns = \PHPExcel_Cell::columnIndexFromString($column);
        $field = [
            'uid','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z',
            'AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ',
            'BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ'
        ];
        $data = [];
        for($row=6;$row<=$rows;$row++){
            $row_data=[];
            for($col=0;$col<$columns;$col++){
                $value=$sheet->getCellByColumnAndRow($col,$row)->getValue();
                $row_data[$field[$col+1]]=$value;
                $row_data[$field[0]]=$admin_name['Id'];
            }
            $data[]=$row_data;

        }
        return $data;
    }


    //    上传文件

    protected function upload($file)
    {
        // 移动到框架应用根目录/public/uploads/ 目录下
        $info = $file->move(ROOT_PATH . 'public' . DS . 'uploads');
        $msg = [];
        if ($info) {
            // 成功上传后 获取上传信息
            // 输出 20160820/42a79759f284b767dfcb2a0197904287.jpg
            $msg['code'] = 1;
            $msg['info'] = ROOT_PATH . 'public' . DS . 'uploads' . DS . $info->getSaveName();

        } else {
            // 上传失败获取错误信息
            $msg['code'] = 0;
            $msg['info'] = $file->getError();
        }
        return $msg;
    }


    //    全部导出
    public function expuser()
    {
        $admin_name = db('manager')->where('username',session('username'))->find();

        $phpexcel = new \PHPExcel();
        $phpexcel->setActiveSheetIndex(0);
        $sheet = $phpexcel->getActiveSheet();
        $res = db('qinchabiao')
            ->field("uid,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,
                           AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW,AX,AY,AZ,
                           BA,BB,BC,BD,BE,BF,BG,BH,BI,BJ
            ")
            ->where('uid',$admin_name['Id'])
            ->select();
        $arr = [
            'A' => '地区',
            'B' => '清查单位',
            'C' => '资产总计账面数',
            /* D */
            'D' => '资产总计核实数',
            'E' => '流动资产合计 账面数',
            /* F */
            'F' => '流动资产合计 核实数',
            'G' => '货币资金 账面数',
            /* H */
            'H' => '货币资金 核实数',
            'I' => '短期投资 账面数',
            'J' => '短期投资 核实数',
            'K' => '应收款项 账面数',
            /* L */
            'L' => '应收款项 核实数',
            'M' => '存款 账面数',
            'N' => '存款 核实数',
            'O' => '农业资产合计 账面数',
            'P' => '农业资产合计 核实数',
            'Q' => '牲畜资产 账面数',
            'R' => '牲畜资产 核实数',
            'S' => '林木资产 账面数',
            'T' => '林木资产 核实数',
            'U' => '长期资产合计 账面数',
            /* V */
            'V' => '长期资产合计 核实数',
            'W' => '长期投资 账面数',
            /* X */
            'X' => '长期投资 核实数',
            'Y' => '长期股权投资 账面数',
            /* Z */
            'Z'=> '长期股权投资 核实数',
            'AA' => '固定资产合计 账面数',
            /* AB */
            'AB' => '固定资产合计 核实数',
            'AC' => '固定资产原值 账面数',
            /* AD */
            'AD' => '固定资产原值 核实数',
            'AE' => '累计折扣 账面数',
            'AF' => '累计折扣 核实数',
            'AG' => '固定资产净值 账面数',
            /* AH */
            'AH' => '固定资产净值 核实数',
            'AI' => '经营性固定资产 账面数',
            'AJ' => '经营性固定资产 核实数',
            'AK' => '固定资产清理 账面数',
            /* AL */
            'AL' => '固定资产清理 核实数',
            'AM' => '在建工程 账面数',
            /* AN */
            'AN' => '在建工程 核实数',
            'AO' => '经营性在建工程 账面数',
            'AP' => '经营性在建工程 核实数',
            'AQ' => '其他资产 账面数',
            'AR' => '其他资产 核实数',
            'AS' => '无形资产 账面数',

            'AT' => '无形资产 核实数',
            'AU' => '负债和所有者权益 账面数',
            'AV' => '负债和所有者权益 核实数',
            'AW' => '流动负债合计 账面数',
            /* AX */
            'AX' => '流动负债合计 核实数',
            'AY' => '短期借款 账面数',
            /* AZ */
            'AZ' => '短期借款 核实数',
            'BA' => '应付款项 账面数',
            /* BB */
            'BB' => '应付款项 核实数',
            'BC' => '应付工资 账面数',
            /* BD */
            'BD' => '应付工资 核实数',
            'BE' => '应付福利费 账面数',
            /* BF */
            'BF' => '应付福利费 核实数',
            'BG' => '长期负债合计 账面数',
            /* BH */
            'BH' => '长期负债合计 核实数',
            'BI' => '长期借款及应付款 账面数',
            /* BJ */
            'BJ' => '长期借款及应付款 核实数',
        ];
        array_unshift($res, $arr);
        $currow = 0;
        foreach ($res as $key => $v) {
            $currow = $key + 1;
            $sheet->setCellValue('A' . $currow, $v['A'])
                ->setCellValue('B' . $currow, $v['B'])
                ->setCellValue('C' . $currow, $v['C'])
                ->setCellValue('D' . $currow, $v['D'])
                ->setCellValue('E' . $currow, $v['E'])
                ->setCellValue('F' . $currow, $v['F'])
                ->setCellValue('G' . $currow, $v['G'])
                ->setCellValue('H' . $currow, $v['H'])

                ->setCellValue('I' . $currow, $v['I'])
                ->setCellValue('J' . $currow, $v['J'])
                ->setCellValue('K' . $currow, $v['K'])
                ->setCellValue('L' . $currow, $v['L'])
                ->setCellValue('M' . $currow, $v['M'])
                ->setCellValue('N' . $currow, $v['N'])
                ->setCellValue('O' . $currow, $v['O'])
                ->setCellValue('P' . $currow, $v['P'])
                ->setCellValue('Q' . $currow, $v['Q'])
                ->setCellValue('R' . $currow, $v['R'])
                ->setCellValue('S' . $currow, $v['S'])
                ->setCellValue('T' . $currow, $v['T'])
                ->setCellValue('U' . $currow, $v['U'])
                ->setCellValue('V' . $currow, $v['V'])
                ->setCellValue('W' . $currow, $v['W'])
                ->setCellValue('X' . $currow, $v['X'])
                ->setCellValue('Y' . $currow, $v['Y'])
                ->setCellValue('Z' . $currow, $v['Z'])

                ->setCellValue('AA' . $currow, $v['AA'])
                ->setCellValue('AB' . $currow, $v['AB'])
                ->setCellValue('AC' . $currow, $v['AC'])
                ->setCellValue('AD' . $currow, $v['AD'])
                ->setCellValue('AE' . $currow, $v['AE'])
                ->setCellValue('AF' . $currow, $v['AF'])
                ->setCellValue('AG' . $currow, $v['AG'])
                ->setCellValue('AH' . $currow, $v['AH'])
                ->setCellValue('AI' . $currow, $v['AI'])
                ->setCellValue('AJ' . $currow, $v['AJ'])
                ->setCellValue('AK' . $currow, $v['AK'])
                ->setCellValue('AL' . $currow, $v['AL'])
                ->setCellValue('AM' . $currow, $v['AM'])
                ->setCellValue('AN' . $currow, $v['AN'])
                ->setCellValue('AO' . $currow, $v['AO'])
                ->setCellValue('AP' . $currow, $v['AP'])
                ->setCellValue('AQ' . $currow, $v['AQ'])
                ->setCellValue('AR' . $currow, $v['AR'])
                ->setCellValue('AS' . $currow, $v['AS'])
                ->setCellValue('AT' . $currow, $v['AT'])
                ->setCellValue('AU' . $currow, $v['AU'])
                ->setCellValue('AV' . $currow, $v['AV'])
                ->setCellValue('AW' . $currow, $v['AW'])
                ->setCellValue('AX' . $currow, $v['AX'])
                ->setCellValue('AY' . $currow, $v['AY'])
                ->setCellValue('AZ' . $currow, $v['AZ'])

                ->setCellValue('BA' . $currow, $v['BA'])
                ->setCellValue('BB' . $currow, $v['BB'])
                ->setCellValue('BC' . $currow, $v['BC'])
                ->setCellValue('BD' . $currow, $v['BD'])
                ->setCellValue('BE' . $currow, $v['BE'])
                ->setCellValue('BF' . $currow, $v['BF'])
                ->setCellValue('BG' . $currow, $v['BG'])
                ->setCellValue('BH' . $currow, $v['BH'])
                ->setCellValue('BI' . $currow, $v['BI'])
                ->setCellValue('BJ' . $currow, $v['BJ'])
            ;
        }
        $phpexcel->getActiveSheet()->getStyle('A1:BJ' . $currow)->getBorders()->getAllBorders()->setBorderStyle(\PHPExcel_Style_Border::BORDER_THIN);
        header('Content-Type: application/vnd.ms-excel');//设置文档类型
        header('Content-Disposition: attachment;filename="清查表全部信息输出.xlsx"');//设置文件名
        header('Cache-Control: max-age=0');
        $phpwriter = new \PHPExcel_Writer_Excel2007($phpexcel);
        $phpwriter->save('php://output');
        return;
    }

    //    带标记原表导出
    public function conditionAll()
    {
        $admin_name = db('manager')->where('username',session('username'))->find();

        $phpexcel = new \PHPExcel();
        $phpexcel->setActiveSheetIndex(0);
        $sheet = $phpexcel->getActiveSheet();
        $map['uid'] = $admin_name['Id'];
        $res = db('qinchabiao')
            ->field("uid,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,
                           AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW,AX,AY,AZ,
                           BA,BB,BC,BD,BE,BF,BG,BH,BI,BJ
            ")
            ->where($map)
            ->select();
        $arr = [
              'A' => '地区',
              'B' => '清查单位',
              /* C */
              'C' => '集体土地总面积',
              'D' => '备注',
              /* E */
              'E' => '农用地面积',
              'F' => '备注',
              'G' => '耕地面积',
              'H' => '备注',
              'I' => '其中未承包到户面积',
              'J' => '备注',
              'K' => '园地面积',
              'L' => '备注',
              'M' => '未承包到户面积',
              'N' => '备注',
              /* O */
              'O' => '林地面积',
              'P' => '备注',
              'Q' => '其中未承包到户面积',
              'R' => '其中未承包到户面积备注',
              'S' => '草地面积',
              'T' => '备注',
              'U' => '其中未承包到户面积',
              'V' => '备注',
              'W' => '农田水利设施用地',
              'X' => '备注',
              'Y' => '养殖水面面积',
              'Z' => '备注',
              'AA' => '其中未承包面积',
              'AB' => '其中未承包到户面积备注',
              'AC' => '其他农用地',
              'AD' => '备注',
              'AE' => '建设用地面积',
              'AF' => '备注',
              'AG' => '工矿仓储面积',
              'AH' => '备注',
              'AI' => '商服用地面积',
              'AJ' => '备注',
                /* AK */
              'AK' => '农村宅基地面积',
              'AL' => '备注',
              /* AM */
              'AM' => '公共管理与公共服务用地面积',
              'AN' => '备注',
              'AO' => '交通运输和水利设施用地面积',
              'AP' => '备注',
              'AQ' => '其他建设用地面积',
              'AR' => '备注',
              'AS' => '未利用地面积',
            
              'AT' => '备注',
              'AU' => '四荒地面积',
              'AV' => '备注',
              'AW' => '待界定土地面积',
              'AX' => '备注',
              'AY' => '待界定农用地面积',
              'AZ' => '备注',
              'BA' => '待界定建设用地面积',
              'BB' => '备注',
              'BC' => '待界定未利用地面积',
              'BD' => '备注',
              /* BE */
              'BE' => '林地面积',
              /* BF */
              'BF' => '备注',
              'BG' => '公益林(立方米)面积',
              'BH' => '备注',
              'BI' => '商品林(立方米)面积',
              'BJ' => '备注',
        ];
        array_unshift($res, $arr);
        $currow = 0;
        foreach ($res as $key => $v) {
            $currow = $key + 1;
            $sheet->setCellValue('A' . $currow, $v['A'])
                ->setCellValue('B' . $currow, $v['B'])
                ->setCellValue('C' . $currow, $v['C'])
                ->setCellValue('D' . $currow, $v['D'])
                ->setCellValue('E' . $currow, $v['E'])
                ->setCellValue('F' . $currow, $v['F'])
                ->setCellValue('G' . $currow, $v['G'])
                ->setCellValue('H' . $currow, $v['H'])

                ->setCellValue('I' . $currow, $v['I'])
                ->setCellValue('J' . $currow, $v['J'])
                ->setCellValue('K' . $currow, $v['K'])
                ->setCellValue('L' . $currow, $v['L'])
                ->setCellValue('M' . $currow, $v['M'])
                ->setCellValue('N' . $currow, $v['N'])
                ->setCellValue('O' . $currow, $v['O'])
                ->setCellValue('P' . $currow, $v['P'])
                ->setCellValue('Q' . $currow, $v['Q'])
                ->setCellValue('R' . $currow, $v['R'])
                ->setCellValue('S' . $currow, $v['S'])
                ->setCellValue('T' . $currow, $v['T'])
                ->setCellValue('U' . $currow, $v['U'])
                ->setCellValue('V' . $currow, $v['V'])
                ->setCellValue('W' . $currow, $v['W'])
                ->setCellValue('X' . $currow, $v['X'])
                ->setCellValue('Y' . $currow, $v['Y'])
                ->setCellValue('Z' . $currow, $v['Z'])

                ->setCellValue('AA' . $currow, $v['AA'])
                ->setCellValue('AB' . $currow, $v['AB'])
                ->setCellValue('AC' . $currow, $v['AC'])
                ->setCellValue('AD' . $currow, $v['AD'])
                ->setCellValue('AE' . $currow, $v['AE'])
                ->setCellValue('AF' . $currow, $v['AF'])
                ->setCellValue('AG' . $currow, $v['AG'])
                ->setCellValue('AH' . $currow, $v['AH'])
                ->setCellValue('AI' . $currow, $v['AI'])
                ->setCellValue('AJ' . $currow, $v['AJ'])
                ->setCellValue('AK' . $currow, $v['AK'])
                ->setCellValue('AL' . $currow, $v['AL'])
                ->setCellValue('AM' . $currow, $v['AM'])
                ->setCellValue('AN' . $currow, $v['AN'])
                ->setCellValue('AO' . $currow, $v['AO'])
                ->setCellValue('AP' . $currow, $v['AP'])
                ->setCellValue('AQ' . $currow, $v['AQ'])
                ->setCellValue('AR' . $currow, $v['AR'])
                ->setCellValue('AS' . $currow, $v['AS'])
                ->setCellValue('AT' . $currow, $v['AT'])
                ->setCellValue('AU' . $currow, $v['AU'])
                ->setCellValue('AV' . $currow, $v['AV'])
                ->setCellValue('AW' . $currow, $v['AW'])
                ->setCellValue('AX' . $currow, $v['AX'])
                ->setCellValue('AY' . $currow, $v['AY'])
                ->setCellValue('AZ' . $currow, $v['AZ'])

                ->setCellValue('BA' . $currow, $v['BA'])
                ->setCellValue('BB' . $currow, $v['BB'])
                ->setCellValue('BC' . $currow, $v['BC'])
                ->setCellValue('BD' . $currow, $v['BD'])
                ->setCellValue('BE' . $currow, $v['BE'])
                ->setCellValue('BF' . $currow, $v['BF'])
                ->setCellValue('BG' . $currow, $v['BG'])
                ->setCellValue('BH' . $currow, $v['BH'])
                ->setCellValue('BI' . $currow, $v['BI'])
                ->setCellValue('BJ' . $currow, $v['BJ'])
            ;
        }
        $phpexcel->getActiveSheet()->getStyle('A1:BJ' . $currow)->getBorders()->getAllBorders()->setBorderStyle(\PHPExcel_Style_Border::BORDER_THIN);
        ///添加判断条件

        foreach ($res as $key=> $v){
            //记录被标记的行数(出错的个数)
            $counter = 0;
            $singleFlag = False;
            $currow=$key+1;
            if($currow!=1 && $v['C']==0)   //1
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('C'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '99FFFF')));
            }
            if ($currow!=1 && $v['E']==0)   //2
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('E'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '99FFCC')));
            }
            if ($currow!=1 && $v['AK']==0)  //3
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AK'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '99FF66')));
            }
            if ($currow!=1 && $v['AM']==0)  //4
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AM'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => '99FF00')));
            }

            if ($currow!=1 && (($v['O']==0 && $v['BE'] != 0) || ($v['O']!=0 && $v['BE'] == 0)))   //5
//            if ($currow!=1 && !(($v['O']==0 && $v['BE'] == 0)))   //5
            {
                $singleFlag = True;
                $phpexcel->getActiveSheet()->getStyle('AY'.$currow.':BB'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CCCC99')));
            }

            //对最后的错误计数统计并作出处理
            if($counter>0 && $counter <2)
            {
                $phpexcel->getActiveSheet()->getStyle('BF'.$currow.':BJ'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '99FFFF')));
            }elseif ($counter>=2 && $counter<4)
            {
                $phpexcel->getActiveSheet()->getStyle('BF'.$currow.':BJ'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => 'CC6600')));
            }elseif($counter==4)
            {
                //  BF:BJ   标记四个检测项的错误程度

                $phpexcel->getActiveSheet()->getStyle('BF'.$currow.':BJ'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => 'FF3300')));
            }

            if($singleFlag)
            {
                $phpexcel->getActiveSheet()->getStyle('O'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => 'FFFF33')));
                $phpexcel->getActiveSheet()->getStyle('BE'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => 'FFFF33')));
            }
        }

        header('Content-Type: application/vnd.ms-excel');//设置文档类型
        header('Content-Disposition: attachment;filename="清查表全部标记输出.xlsx"');//设置文件名
        header('Cache-Control: max-age=0');
        $phpwriter = new \PHPExcel_Writer_Excel2007($phpexcel);
        $phpwriter->save('php://output');
        return;
    }

    //导出被标记子表
    public function conditionChildren()
    {
        $admin_name = db('manager')->where('username',session('username'))->find();

        $phpexcel = new \PHPExcel();
        $phpexcel->setActiveSheetIndex(0);
        $sheet = $phpexcel->getActiveSheet();
        $together['C']=array('eq',0);
        $together['E']=array('eq',0);
        $together['AK']=array('eq',0);
        $together['AM']=array('eq',0);
        $another['O']=array('eq',0);
        $another['BE']=array('eq',0);

        $map['uid'] = $admin_name['Id'];
        $res = db('qinchabiao')
            ->field("uid,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,
                           AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW,AX,AY,AZ,
                           BA,BB,BC,BD,BE,BF,BG,BH,BI,BJ
            ")
            //根据条件直接从数据库面拿出数据
            ->where($map)
            ->whereor($together)
            ->whereXor($another)
            ->select();
        $arr = [
            'A' => '地区',
            'B' => '清查单位',
            /* C */
            'C' => '集体土地总面积',
            'D' => '备注',
            /* E */
            'E' => '农用地面积',
            'F' => '备注',
            'G' => '耕地面积',
            'H' => '备注',
            'I' => '其中未承包到户面积',
            'J' => '备注',
            'K' => '园地面积',
            'L' => '备注',
            'M' => '未承包到户面积',
            'N' => '备注',
            /* O */
            'O' => '林地面积',
            'P' => '备注',
            'Q' => '其中未承包到户面积',
            'R' => '其中未承包到户面积备注',
            'S' => '草地面积',
            'T' => '备注',
            'U' => '其中未承包到户面积',
            'V' => '备注',
            'W' => '农田水利设施用地',
            'X' => '备注',
            'Y' => '养殖水面面积',
            'Z' => '备注',
            'AA' => '其中未承包面积',
            'AB' => '其中未承包到户面积备注',
            'AC' => '其他农用地',
            'AD' => '备注',
            'AE' => '建设用地面积',
            'AF' => '备注',
            'AG' => '工矿仓储面积',
            'AH' => '备注',
            'AI' => '商服用地面积',
            'AJ' => '备注',
            /* AK */
            'AK' => '农村宅基地面积',
            'AL' => '备注',
            /* AM */
            'AM' => '公共管理与公共服务用地面积',
            'AN' => '备注',
            'AO' => '交通运输和水利设施用地面积',
            'AP' => '备注',
            'AQ' => '其他建设用地面积',
            'AR' => '备注',
            'AS' => '未利用地面积',

            'AT' => '备注',
            'AU' => '四荒地面积',
            'AV' => '备注',
            'AW' => '待界定土地面积',
            'AX' => '备注',
            'AY' => '待界定农用地面积',
            'AZ' => '备注',
            'BA' => '待界定建设用地面积',
            'BB' => '备注',
            'BC' => '待界定未利用地面积',
            'BD' => '备注',
            /* BE */
            'BE' => '林地面积',
            /* BF */
            'BF' => '备注',
            'BG' => '公益林(立方米)面积',
            'BH' => '备注',
            'BI' => '商品林(立方米)面积',
            'BJ' => '备注',
        ];
        array_unshift($res, $arr);
        $currow = 0;
        foreach ($res as $key => $v) {
            $currow = $key + 1;
            $sheet->setCellValue('A' . $currow, $v['A'])
                ->setCellValue('B' . $currow, $v['B'])
                ->setCellValue('C' . $currow, $v['C'])
                ->setCellValue('D' . $currow, $v['D'])
                ->setCellValue('E' . $currow, $v['E'])
                ->setCellValue('F' . $currow, $v['F'])
                ->setCellValue('G' . $currow, $v['G'])
                ->setCellValue('H' . $currow, $v['H'])

                ->setCellValue('I' . $currow, $v['I'])
                ->setCellValue('J' . $currow, $v['J'])
                ->setCellValue('K' . $currow, $v['K'])
                ->setCellValue('L' . $currow, $v['L'])
                ->setCellValue('M' . $currow, $v['M'])
                ->setCellValue('N' . $currow, $v['N'])
                ->setCellValue('O' . $currow, $v['O'])
                ->setCellValue('P' . $currow, $v['P'])
                ->setCellValue('Q' . $currow, $v['Q'])
                ->setCellValue('R' . $currow, $v['R'])
                ->setCellValue('S' . $currow, $v['S'])
                ->setCellValue('T' . $currow, $v['T'])
                ->setCellValue('U' . $currow, $v['U'])
                ->setCellValue('V' . $currow, $v['V'])
                ->setCellValue('W' . $currow, $v['W'])
                ->setCellValue('X' . $currow, $v['X'])
                ->setCellValue('Y' . $currow, $v['Y'])
                ->setCellValue('Z' . $currow, $v['Z'])

                ->setCellValue('AA' . $currow, $v['AA'])
                ->setCellValue('AB' . $currow, $v['AB'])
                ->setCellValue('AC' . $currow, $v['AC'])
                ->setCellValue('AD' . $currow, $v['AD'])
                ->setCellValue('AE' . $currow, $v['AE'])
                ->setCellValue('AF' . $currow, $v['AF'])
                ->setCellValue('AG' . $currow, $v['AG'])
                ->setCellValue('AH' . $currow, $v['AH'])
                ->setCellValue('AI' . $currow, $v['AI'])
                ->setCellValue('AJ' . $currow, $v['AJ'])
                ->setCellValue('AK' . $currow, $v['AK'])
                ->setCellValue('AL' . $currow, $v['AL'])
                ->setCellValue('AM' . $currow, $v['AM'])
                ->setCellValue('AN' . $currow, $v['AN'])
                ->setCellValue('AO' . $currow, $v['AO'])
                ->setCellValue('AP' . $currow, $v['AP'])
                ->setCellValue('AQ' . $currow, $v['AQ'])
                ->setCellValue('AR' . $currow, $v['AR'])
                ->setCellValue('AS' . $currow, $v['AS'])
                ->setCellValue('AT' . $currow, $v['AT'])
                ->setCellValue('AU' . $currow, $v['AU'])
                ->setCellValue('AV' . $currow, $v['AV'])
                ->setCellValue('AW' . $currow, $v['AW'])
                ->setCellValue('AX' . $currow, $v['AX'])
                ->setCellValue('AY' . $currow, $v['AY'])
                ->setCellValue('AZ' . $currow, $v['AZ'])

                ->setCellValue('BA' . $currow, $v['BA'])
                ->setCellValue('BB' . $currow, $v['BB'])
                ->setCellValue('BC' . $currow, $v['BC'])
                ->setCellValue('BD' . $currow, $v['BD'])
                ->setCellValue('BE' . $currow, $v['BE'])
                ->setCellValue('BF' . $currow, $v['BF'])
                ->setCellValue('BG' . $currow, $v['BG'])
                ->setCellValue('BH' . $currow, $v['BH'])
                ->setCellValue('BI' . $currow, $v['BI'])
                ->setCellValue('BJ' . $currow, $v['BJ'])
            ;
        }
        $phpexcel->getActiveSheet()->getStyle('A1:BJ' . $currow)->getBorders()->getAllBorders()->setBorderStyle(\PHPExcel_Style_Border::BORDER_THIN);


        //添加判断条件

        foreach ($res as $key=> $v){
            //记录被标记的行数(出错的个数)
            $counter = 0;
            $singleFlag = False;
            $currow=$key+1;
            if($currow!=1 && $v['C']==0)   //1
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('C'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '99FFFF')));
            }
            if ($currow!=1 && $v['E']==0)   //2
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('E'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '99FFCC')));
            }
            if ($currow!=1 && $v['AK']==0)  //3
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AK'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '99FF66')));
            }
            if ($currow!=1 && $v['AM']==0)  //4
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AM'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => '99FF00')));
            }

            if ($currow!=1 && (($v['O']==0 && $v['BE'] != 0) || ($v['O']!=0 && $v['BE'] == 0)))   //5
//            if ($currow!=1 && !(($v['O']==0 && $v['BE'] == 0)))   //5
            {
                $singleFlag = True;
                $phpexcel->getActiveSheet()->getStyle('AY'.$currow.':BB'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CCCC99')));
            }

            //对最后的错误计数统计并作出处理
            if($counter>0 && $counter <2)
            {
                $phpexcel->getActiveSheet()->getStyle('BF'.$currow.':BJ'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '99FFFF')));
            }elseif ($counter>=2 && $counter<4)
            {
                $phpexcel->getActiveSheet()->getStyle('BF'.$currow.':BJ'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => 'CC6600')));
            }elseif($counter==4)
            {
                //  BF:BJ   标记四个检测项的错误程度

                $phpexcel->getActiveSheet()->getStyle('BF'.$currow.':BJ'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => 'FF3300')));
            }

            if($singleFlag)
            {
                $phpexcel->getActiveSheet()->getStyle('O'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => 'FFFF33')));
                $phpexcel->getActiveSheet()->getStyle('BE'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => 'FFFF33')));
            }
        }

        header('Content-Type: application/vnd.ms-excel');//设置文档类型
        header('Content-Disposition: attachment;filename="带标记子表全部输出.xlsx"');//设置文件名
        header('Cache-Control: max-age=0');
        $phpwriter = new \PHPExcel_Writer_Excel2007($phpexcel);
        $phpwriter->save('php://output');
        return ;
    }

    //清空数据表 bb_user
    public function trunCate()
    {
        $db=new \mysqli();
        $db->connect('localhost','root','root','excel');
        $sql="TRUNCATE bb_qinchabiao";
        if ($db->query($sql)){
            $this->success('删除成功');
        }else{
            $this->success('删除失败');
        }
        return;
    }


    public function TableTest01()
    {
        $phpexcel = new \PHPExcel();
        $phpexcel->setActiveSheetIndex(0);
        $sheet = $phpexcel->getActiveSheet();
        $res = db('user')->field("name,sex,age,phone,weixin,qq,ticheng,state")->select();
        $arr = [
            'name' => "姓名",
            'sex' => "性别",
            'age' => "年龄",
            'phone' => "手机号",
            'weixin' => "微信",
            'qq' => "QQ",
            'ticheng' => "提成比例",
            'state' => "状态",
        ];
        array_unshift($res, $arr);
        $currow = 0;
        foreach ($res as $key => $v) {
            $currow = $key + 1;
            $sheet->setCellValue('A' . $currow, $v['name'])
                ->setCellValue('B' . $currow, setsex($v['sex']))
                ->setCellValue('C' . $currow, $v['age'])
                ->setCellValue('D' . $currow, $v['phone'])
                ->setCellValue('E' . $currow, $v['weixin'])
                ->setCellValue('F' . $currow, $v['qq'])
                ->setCellValue('G' . $currow, $v['ticheng'])
                ->setCellValue('H' . $currow, setstate($v['state']));
        }
        $phpexcel->getActiveSheet()->getStyle('A1:H' . $currow)->getBorders()->getAllBorders()->setBorderStyle(\PHPExcel_Style_Border::BORDER_THIN);

        foreach ($res as $key=> $v){
            $flag = False;
            $first = False;
            $second = False;
            $currow=$key+1;
            if($v['age']>30 && $v['sex'] =='女')
            {
                $first = True;
//                $phpexcel->getActiveSheet()->getStyle('A'.$currow.':H'.$currow)
//                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
//                        'startcolor' => array('rgb' => 'FF0000')));
                $phpexcel->getActiveSheet()->getStyle('A'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => 'FF0000')));
            }
            if ($v['state'] == '1')
            {
                $second = True;
                $phpexcel->getActiveSheet()->getStyle('H'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '00FF00')));
            }
            $flag = $first && $second;
            if($flag)
            {
                $phpexcel->getActiveSheet()->getStyle('B'.$currow.':G'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => 'FFFF99')));
            }
        }
        header('Content-Type: application/vnd.ms-excel');//设置文档类型
        header('Content-Disposition: attachment;filename="员工信表.xlsx"');//设置文件名
        header('Cache-Control: max-age=0');
        $phpwriter = new \PHPExcel_Writer_Excel2007($phpexcel);
        $phpwriter->save('php://output');
        return;
    }
}