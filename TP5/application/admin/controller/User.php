<?php
/**
 * Created by PhpStorm.
 * User: yifeng
 * Date: 2017/8/10
 * Time: 22:57
 */

namespace app\admin\controller;
use think\db\Query;

class User extends Common
{
    //员工列表
    public function index()
    {
        $admin_name = db('manager')->where('username',session('username'))->find();
        $Unique = $admin_name['Id'];
        $result = db('fuzhaibiao')->where('uid',$Unique)->order('Id Desc')->paginate(8);
        $this->assign('user', $result);
        $together['D']=array('elt',0);
        $together['F']=array('elt',0);
        $together['H']=array('elt',0);
        $together['L']=array('lt',0);
        $together['V']=array('lt',0);
        $together['X']=array('lt',0);
        $together['Z']=array('lt',0);
        $together['AB']=array('elt',0);
        $finalcounter = db('fuzhaibiao')
            ->field("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,
                           AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW,AX,AY,AZ,
                           BA,BB,BC,BD,BE,BF,BG,BH,BI,BJ,BK,BL,BM,BN,BO,BP,BQ,BR,BS,BT,BU,BV,BW,BX,BY,BZ,
                           CA,CB,CC,CD,CE,CF,CG,CH,CI,CJ
            ")
            ->whereor($together)->select();
        $this->assign('counter', $finalcounter);
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
        if (!db('manager')->delete($id)) {
            $this->error('删除失败');
        }
        $this->success("删除成功", 'user/index');
        return;
    }

    //导入表格数据
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
            //上传插入到fuzhibao
            if (!db('fuzhaibiao')->insertAll($data)) {
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
            'BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY','BZ',
            'CA','CB','CC','CD','CE','CF','CG','CH','CI','CJ'
        ];
        $data = [];
//
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
        $res = db('fuzhaibiao')
            ->field("uid,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,
                           AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW,AX,AY,AZ,
                           BA,BB,BC,BD,BE,BF,BG,BH,BI,BJ,BK,BL,BM,BN,BO,BP,BQ,BR,BS,BT,BU,BV,BW,BX,BY,BZ,
                           CA,CB,CC,CD,CE,CF,CG,CH,CI,CJ
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
      'BK' => '一事一议资金 账面数',
      /* BL */
      'BL' => '一事一议资金 核实数',
      'BM' => '专项应付款 账面数',
      /* BN */
      'BN' => '专项应付款 核实数',
      'BO' => '征地补偿费 账面数',
      /* BP */
      'BP' => '征地补偿费 核实数',
      'BQ' => '合计 账面数',
      'BR' => '合计 核实数',
      'BS' => '资本 账面数',
      'BT' => '资本 核实数',
      'BU' => '政府拨款等形式 账面数',
      /* BV */
      'BV' => '政府拨款等形式 核实数',
      'BW' => '公益公积金 账面数',
      'BX' => '公益公积金 核实数',
      'BY' => '征地补偿转入 账面数',
      'BZ' => '征地补偿转入 核实数',
      'CA' => '未分配收益 账面数',
      'CB' => '未分配收益 核实数',
      'CC' => '经营性资产 账面数',
      'CD' => '经营性资产 核实数',
      'CE' => '非经营性资产 账面数',
      /* CF */
      'CF' => '非经营性资产 核实数',
      'CG' => '待界定资产 账面数',
      'CH' => '待界定资产 核实数',
      'CI' => '全资子公司所有者权益 账面数',
      'CJ' => '全资子公司所有者权益 核实数',
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
                ->setCellValue('BK' . $currow, $v['BK'])
                ->setCellValue('BL' . $currow, $v['BL'])
                ->setCellValue('BM' . $currow, $v['BM'])
                ->setCellValue('BN' . $currow, $v['BN'])
                ->setCellValue('BO' . $currow, $v['BO'])
                ->setCellValue('BP' . $currow, $v['BP'])
                ->setCellValue('BQ' . $currow, $v['BQ'])
                ->setCellValue('BR' . $currow, $v['BR'])
                ->setCellValue('BS' . $currow, $v['BS'])
                ->setCellValue('BT' . $currow, $v['BT'])
                ->setCellValue('BU' . $currow, $v['BU'])
                ->setCellValue('BV' . $currow, $v['BV'])
                ->setCellValue('BW' . $currow, $v['BW'])
                ->setCellValue('BX' . $currow, $v['BX'])
                ->setCellValue('BY' . $currow, $v['BY'])
                ->setCellValue('BZ' . $currow, $v['BZ'])

                ->setCellValue('CA' . $currow, $v['CA'])
                ->setCellValue('CB' . $currow, $v['CB'])
                ->setCellValue('CC' . $currow, $v['CC'])
                ->setCellValue('CD' . $currow, $v['CD'])
                ->setCellValue('CE' . $currow, $v['CE'])
                ->setCellValue('CF' . $currow, $v['CF'])
                ->setCellValue('CG' . $currow, $v['CG'])
                ->setCellValue('CH' . $currow, $v['CH'])
                ->setCellValue('CI' . $currow, $v['CI'])
                ->setCellValue('CJ' . $currow, $v['CJ'])
            ;
        }
        $phpexcel->getActiveSheet()->getStyle('A1:CJ' . $currow)->getBorders()->getAllBorders()->setBorderStyle(\PHPExcel_Style_Border::BORDER_THIN);
        header('Content-Type: application/vnd.ms-excel');//设置文档类型
        header('Content-Disposition: attachment;filename="负债表全部信息输出.xlsx"');//设置文件名
        header('Cache-Control: max-age=0');
        $phpwriter = new \PHPExcel_Writer_Excel2007($phpexcel);
        $phpwriter->save('php://output');
        return;
    }
    //    按照条件1输出
    public function conditionAll()
    {
        $admin_name = db('manager')->where('username',session('username'))->find();
        $phpexcel = new \PHPExcel();
        $phpexcel->setActiveSheetIndex(0);
        $sheet = $phpexcel->getActiveSheet();
        $map['uid'] = $admin_name['Id'];
        $res = db('fuzhaibiao')
            ->field("uid,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,
                           AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW,AX,AY,AZ,
                           BA,BB,BC,BD,BE,BF,BG,BH,BI,BJ,BK,BL,BM,BN,BO,BP,BQ,BR,BS,BT,BU,BV,BW,BX,BY,BZ,
                           CA,CB,CC,CD,CE,CF,CG,CH,CI,CJ
            ")
            ->where($map)
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
      'BK' => '一事一议资金 账面数',
      /* BL */
      'BL' => '一事一议资金 核实数',
      'BM' => '专项应付款 账面数',
      /* BN */
      'BN' => '专项应付款 核实数',
      'BO' => '征地补偿费 账面数',
      /* BP */
      'BP' => '征地补偿费 核实数',
      'BQ' => '合计 账面数',
      'BR' => '合计 核实数',
      'BS' => '资本 账面数',
      'BT' => '资本 核实数',
      'BU' => '政府拨款等形式 账面数',
      /* BV */
      'BV' => '政府拨款等形式 核实数',
      'BW' => '公益公积金 账面数',
      'BX' => '公益公积金 核实数',
      'BY' => '征地补偿转入 账面数',
      'BZ' => '征地补偿转入 核实数',
      'CA' => '未分配收益 账面数',
      'CB' => '未分配收益 核实数',
      'CC' => '经营性资产 账面数',
      'CD' => '经营性资产 核实数',
      'CE' => '非经营性资产 账面数',
      /* CF */
      'CF' => '非经营性资产 核实数',
      'CG' => '待界定资产 账面数',
      'CH' => '待界定资产 核实数',
      'CI' => '全资子公司所有者权益 账面数',
      'CJ' => '全资子公司所有者权益 核实数',
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
                ->setCellValue('BK' . $currow, $v['BK'])
                ->setCellValue('BL' . $currow, $v['BL'])
                ->setCellValue('BM' . $currow, $v['BM'])
                ->setCellValue('BN' . $currow, $v['BN'])
                ->setCellValue('BO' . $currow, $v['BO'])
                ->setCellValue('BP' . $currow, $v['BP'])
                ->setCellValue('BQ' . $currow, $v['BQ'])
                ->setCellValue('BR' . $currow, $v['BR'])
                ->setCellValue('BS' . $currow, $v['BS'])
                ->setCellValue('BT' . $currow, $v['BT'])
                ->setCellValue('BU' . $currow, $v['BU'])
                ->setCellValue('BV' . $currow, $v['BV'])
                ->setCellValue('BW' . $currow, $v['BW'])
                ->setCellValue('BX' . $currow, $v['BX'])
                ->setCellValue('BY' . $currow, $v['BY'])
                ->setCellValue('BZ' . $currow, $v['BZ'])

                ->setCellValue('CA' . $currow, $v['CA'])
                ->setCellValue('CB' . $currow, $v['CB'])
                ->setCellValue('CC' . $currow, $v['CC'])
                ->setCellValue('CD' . $currow, $v['CD'])
                ->setCellValue('CE' . $currow, $v['CE'])
                ->setCellValue('CF' . $currow, $v['CF'])
                ->setCellValue('CG' . $currow, $v['CG'])
                ->setCellValue('CH' . $currow, $v['CH'])
                ->setCellValue('CI' . $currow, $v['CI'])
                ->setCellValue('CJ' . $currow, $v['CJ'])
            ;
        }
        $phpexcel->getActiveSheet()->getStyle('A1:CJ' . $currow)->getBorders()->getAllBorders()->setBorderStyle(\PHPExcel_Style_Border::BORDER_THIN);
        //添加判断条件

        foreach ($res as $key=> $v){
            //记录被标记的行数(出错的个数)
            $counter = 0;
            $currow=$key+1;
            if($currow!=1 && $v['D']<=0)   //1
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('D'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '99FFFF')));
            }
            if ($currow!=1 && $v['F']<=0)   //2
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('F'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '99FFCC')));
            }
            if ($currow!=1 && $v['H']<=0)  //3
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('H'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '99FF66')));
            }
            if ($currow!=1 && $v['L']<0)  //4
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('L'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '99FF00')));
            }
            if ($currow!=1 && $v['V']<0)   //5
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('V'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CCCC99')));
            }
            if ($currow!=1 && $v['X']<0)   //6
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('X'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CCCC66')));
            }
            if ($currow!=1 && $v['Z']<0)    //7
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('Z'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CCCC33')));
            }
            if ($currow!=1 && $v['AB']=0)   //8
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AB'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CCCC00')));
            }
            if ($currow!=1 && $v['AD']<=0)  //9
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AD'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'FF66FF')));
            }
            if ($currow!=1 && $v['AH']<=0)   //10
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AH'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'FF66CC')));
            }
            if ($currow!=1 && $v['AL']<=0) //11
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AL'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'FF6699')));
            }
            if ($currow!=1 && $v['AN']<0) //12
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AN'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'FF6666')));
            }
            if ($currow!=1 && $v['AX']<0)  //13
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AX'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'FFFF99')));
            }
            if ($currow!=1 && $v['AZ']<0) //14
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AZ'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'FFFF66')));
            }
            if ($currow!=1 && $v['BB']<0)  //15
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BB'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'FFFF33')));
            }
            if ($currow!=1 && $v['BD']<0)  //16
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BD'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'FFFF00')));
            }
            if ($currow!=1 && $v['BF']<0) //17
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BF'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => '9966CC')));
            }
            if ($currow!=1 && $v['BH']<0)  //18
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BH'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => '996699')));
            }
            if ($currow!=1 && $v['BJ']<0)  //19
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BJ'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => '9966FF')));
            }
            if ($currow!=1 && $v['BL']<0)   //20
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BL'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => '996666')));
            }
            if ($currow!=1 && $v['BN']<0)  //21
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BN'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CC0066')));
            }
            if ($currow!=1 && $v['BP']<0)  //22
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BP'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CC0099')));
            }
            if ($currow!=1 && $v['BV']<0)  //23
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BV'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CC00FF')));
            }
            if ($currow!=1 && $v['CF']<0)  //24
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('CF'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CCFF00')));
            }

            //对最后的错误计数统计并作出处理
            if($counter>0 && $counter <8)
            {
                $phpexcel->getActiveSheet()->getStyle('CG'.$currow.':CJ'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => 'FFFF33')));
            }elseif ($counter>8 && $counter<16)
            {
                $phpexcel->getActiveSheet()->getStyle('CG'.$currow.':CJ'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '996600')));
            }elseif($counter>16)
            {
                $phpexcel->getActiveSheet()->getStyle('CG'.$currow.':CJ'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => 'FF3300')));
            }
        }

        header('Content-Type: application/vnd.ms-excel');//设置文档类型
        header('Content-Disposition: attachment;filename="负债表带标记全表输出.xlsx"');//设置文件名
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
        $together['D']=array('elt',0);
        $together['F']=array('elt',0);
        $together['H']=array('elt',0);
        $together['L']=array('lt',0);
        $together['V']=array('lt',0);
        $together['X']=array('lt',0);
        $together['Z']=array('lt',0);
        $together['AB']=array('elt',0);

        $together['AD']=array('elt',0);
        $together['AH']=array('elt',0);
        $together['AL']=array('elt',0);
        $together['AN']=array('lt',0);
        $together['AX']=array('lt',0);
        $together['AZ']=array('lt',0);
        $together['BB']=array('lt',0);
        $together['BD']=array('lt',0);

        $together['BF']=array('lt',0);
        $together['BH']=array('lt',0);
        $together['BJ']=array('lt',0);
        $together['BL']=array('lt',0);
        $together['BN']=array('lt',0);
        $together['BP']=array('lt',0);
        $together['BV']=array('lt',0);
        $together['CF']=array('lt',0);

        $map['uid'] = $admin_name['Id'];

        $res = db('fuzhaibiao')
            ->field("uid,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,
                           AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW,AX,AY,AZ,
                           BA,BB,BC,BD,BE,BF,BG,BH,BI,BJ,BK,BL,BM,BN,BO,BP,BQ,BR,BS,BT,BU,BV,BW,BX,BY,BZ,
                           CA,CB,CC,CD,CE,CF,CG,CH,CI,CJ
            ")
        ->whereor($together)
        ->where($map)
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
            'BK' => '一事一议资金 账面数',
            /* BL */
            'BL' => '一事一议资金 核实数',
            'BM' => '专项应付款 账面数',
            /* BN */
            'BN' => '专项应付款 核实数',
            'BO' => '征地补偿费 账面数',
            /* BP */
            'BP' => '征地补偿费 核实数',
            'BQ' => '合计 账面数',
            'BR' => '合计 核实数',
            'BS' => '资本 账面数',
            'BT' => '资本 核实数',
            'BU' => '政府拨款等形式 账面数',
            /* BV */
            'BV' => '政府拨款等形式 核实数',
            'BW' => '公益公积金 账面数',
            'BX' => '公益公积金 核实数',
            'BY' => '征地补偿转入 账面数',
            'BZ' => '征地补偿转入 核实数',
            'CA' => '未分配收益 账面数',
            'CB' => '未分配收益 核实数',
            'CC' => '经营性资产 账面数',
            'CD' => '经营性资产 核实数',
            'CE' => '非经营性资产 账面数',
            /* CF */
            'CF' => '非经营性资产 核实数',
            'CG' => '待界定资产 账面数',
            'CH' => '待界定资产 核实数',
            'CI' => '全资子公司所有者权益 账面数',
            'CJ' => '全资子公司所有者权益 核实数',
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
                ->setCellValue('BK' . $currow, $v['BK'])
                ->setCellValue('BL' . $currow, $v['BL'])
                ->setCellValue('BM' . $currow, $v['BM'])
                ->setCellValue('BN' . $currow, $v['BN'])
                ->setCellValue('BO' . $currow, $v['BO'])
                ->setCellValue('BP' . $currow, $v['BP'])
                ->setCellValue('BQ' . $currow, $v['BQ'])
                ->setCellValue('BR' . $currow, $v['BR'])
                ->setCellValue('BS' . $currow, $v['BS'])
                ->setCellValue('BT' . $currow, $v['BT'])
                ->setCellValue('BU' . $currow, $v['BU'])
                ->setCellValue('BV' . $currow, $v['BV'])
                ->setCellValue('BW' . $currow, $v['BW'])
                ->setCellValue('BX' . $currow, $v['BX'])
                ->setCellValue('BY' . $currow, $v['BY'])
                ->setCellValue('BZ' . $currow, $v['BZ'])

                ->setCellValue('CA' . $currow, $v['CA'])
                ->setCellValue('CB' . $currow, $v['CB'])
                ->setCellValue('CC' . $currow, $v['CC'])
                ->setCellValue('CD' . $currow, $v['CD'])
                ->setCellValue('CE' . $currow, $v['CE'])
                ->setCellValue('CF' . $currow, $v['CF'])
                ->setCellValue('CG' . $currow, $v['CG'])
                ->setCellValue('CH' . $currow, $v['CH'])
                ->setCellValue('CI' . $currow, $v['CI'])
                ->setCellValue('CJ' . $currow, $v['CJ'])
            ;
        }
        $phpexcel->getActiveSheet()->getStyle('A1:CJ' . $currow)->getBorders()->getAllBorders()->setBorderStyle(\PHPExcel_Style_Border::BORDER_THIN);


        //添加判断条件

        foreach ($res as $key=> $v){
            //记录被标记的行数(出错的个数)
            $counter = 0;
            $currow=$key+1;
            if($currow!=1 && $v['D']<=0)   //1
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('D'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '99FFFF')));
            }
            if ($currow!=1 && $v['F']<=0)   //2
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('F'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '99FFCC')));
            }
            if ($currow!=1 && $v['H']<=0)  //3
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('H'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '99FF66')));
            }
            if ($currow!=1 && $v['L']<0)  //4
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('L'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => '99FF00')));
            }
            if ($currow!=1 && $v['V']<0)   //5
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('V'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CCCC99')));
            }
            if ($currow!=1 && $v['X']<0)   //6
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('X'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CCCC66')));
            }
            if ($currow!=1 && $v['Z']<0)    //7
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('Z'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CCCC33')));
            }
            if ($currow!=1 && $v['AB']=0)   //8
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AB'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CCCC00')));
            }
            if ($currow!=1 && $v['AD']<=0)  //9
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AD'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'FF66FF')));
            }
            if ($currow!=1 && $v['AH']<=0)   //10
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AH'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'FF66CC')));
            }
            if ($currow!=1 && $v['AL']<=0) //11
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AL'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'FF6699')));
            }
            if ($currow!=1 && $v['AN']<0) //12
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AN'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'FF6666')));
            }
            if ($currow!=1 && $v['AX']<0)  //13
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AX'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'FFFF99')));
            }
            if ($currow!=1 && $v['AZ']<0) //14
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('AZ'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'FFFF66')));
            }
            if ($currow!=1 && $v['BB']<0)  //15
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BB'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'FFFF33')));
            }
            if ($currow!=1 && $v['BD']<0)  //16
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BD'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'FFFF00')));
            }
            if ($currow!=1 && $v['BF']<0) //17
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BF'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => '9966CC')));
            }
            if ($currow!=1 && $v['BH']<0)  //18
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BH'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => '996699')));
            }
            if ($currow!=1 && $v['BJ']<0)  //19
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BJ'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => '9966FF')));
            }
            if ($currow!=1 && $v['BL']<0)   //20
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BL'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => '996666')));
            }
            if ($currow!=1 && $v['BN']<0)  //21
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BN'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CC0066')));
            }
            if ($currow!=1 && $v['BP']<0)  //22
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BP'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CC0099')));
            }
            if ($currow!=1 && $v['BV']<0)  //23
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('BV'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CC00FF')));
            }
            if ($currow!=1 && $v['CF']<0)  //24
            {
                $counter +=1;
                $phpexcel->getActiveSheet()->getStyle('CF'.$currow)->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                    'startcolor' => array('rgb' => 'CCFF00')));
            }

            //对最后的错误计数统计并作出处理
            if($counter>0 && $counter <8)
            {
                $phpexcel->getActiveSheet()->getStyle('CG'.$currow.':CJ'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => 'FFFF33')));
            }elseif ($counter>8 && $counter<16)
            {
                $phpexcel->getActiveSheet()->getStyle('CG'.$currow.':CJ'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => '996600')));
            }elseif($counter>16)
            {
                $phpexcel->getActiveSheet()->getStyle('CG'.$currow.':CJ'.$currow)
                    ->getFill()->applyFromArray(array('type' => \PHPExcel_Style_Fill::FILL_SOLID,
                        'startcolor' => array('rgb' => 'FF3300')));
            }
        }

        header('Content-Type: application/vnd.ms-excel');//设置文档类型
        header('Content-Disposition: attachment;filename="负债表标记子表输出.xlsx"');//设置文件名
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
        $sql="TRUNCATE bb_fuzhaibiao";
        if ($db->query($sql)){
            $this->success('删除成功');
        }else{
            $this->success('删除失败');
        }
        return;
    }


}