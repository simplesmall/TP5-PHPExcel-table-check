<?php
/**
 * Created by PhpStorm.
 * User: yifeng
 * Date: 2017/8/21
 * Time: 22:56
 */

namespace app\admin\controller;

class Report extends Common
{
    public function mingxi(){
        if(request()->isPost()){
            $data=input('post.');
            $user=db('user')->where('Id',$data['uid'])->field('ticheng,name')->find();
            $res=db('yeji')->where('uid',$data['uid'])->whereTime('uptime','between',[$data['time'].'-01',$data['time']."-".date('t',strtotime($data['time']))])->select();
//            报表生成
            //创建一个PHPExcel类
            $phpexcel=new \PHPExcel();
            $phpexcel->setActiveSheetIndex(0);
            $sheet=$phpexcel->getActiveSheet();
            $sheet->setTitle('月份个人明细报表');
            $sheet->setCellValue('A1', $user['name'].$data['time'].'月份业绩明细')
                ->setCellValue('A2','时间')
                ->setCellValue('B2','业绩')
                ->setCellValue('C2','说明')
                ->setCellValue('D2','提成比例(%)')
                ->setCellValue('E2','提成')
                ->setCellValue('F2','状态');
            $sheet->mergeCells('A1:F1');
            $sheet->getRowDimension(1)->setRowHeight(39);
            $sheet->getStyle('A1')->getFont()->setSize(18);
            $sheet->getStyle('A1:F2')->getFont()->setBold(true);
            $sheet->getStyle('A1:F2')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $sheet->getStyle('A1:F2')->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $sheet->getColumnDimension('A')->setWidth(15);
            $sheet->getColumnDimension('B')->setWidth(10);
            $sheet->getColumnDimension('C')->setWidth(25);
            $sheet->getColumnDimension('D')->setWidth(12);
            $sheet->getColumnDimension('E')->setWidth(14);
            $sheet->getColumnDimension('F')->setWidth(13);
            //写入数据
            $currow=2;
            foreach ($res as $key=>$vo){
                $currow=$currow+1;
                $sheet->setCellValue('A'.$currow,date('Y/m/d H:i:s',$vo['uptime']))
                    ->setCellValue('B'.$currow,$vo['yeji'])
                    ->setCellValue('C'.$currow,$vo['note'])
                    ->setCellValue('D'.$currow,$user['ticheng'])
                    ->setCellValue('E'.$currow,'=B'.$currow."*D".$currow."/100")
                    ->setCellValue('F'.$currow,getstatecn($vo['state']));
            }
            $sheet->setCellValue('A'.($currow+1),"合计");
            $sheet->setCellValue('B'.($currow+1),"=sum(B3:B".$currow.")");
            $sheet->setCellValue('E'.($currow+1),"=sum(E3:E".$currow.")");
            $sheet->getStyle('A1:F'.($currow+1))->getBorders()->getAllBorders()->setBorderStyle(\PHPExcel_Style_Border::BORDER_THIN);
            $filename=$user['name'].$data['time'];
            $this->excelsave($phpexcel,$filename);
            return;
        }
        $userres=db('user')->where('state',1)->field('Id,name,phone')->select();
        $this->assign('user',$userres);
        return view();
    }

    public function yuebao(){

        if(request()->isPost()){
            $time=input('time');
            $phpexcel=new \PHPExcel();
            $phpexcel->setActiveSheetIndex(0);
            $sheet=$phpexcel->getActiveSheet();
            $sheet->setTitle('员工月报表');
            $sheet->setCellValue('A1','姓名')
                ->setCellValue('B1','手机号')
                ->setCellValue('C1',$time.'月份')
                ->setCellValue('E1','状态')
                ->setCellValue('F1','业绩合计')
                ->setCellValue('C2','业绩')
                ->setCellValue('D2','提成')
                ->setCellValue('F2','本月')
                ->setCellValue('G2','本季度')
                ->setCellValue('H2','本年度');
            $sheet->getStyle('A1:H2')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $sheet->getStyle('A1:H2')->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $sheet->mergeCells('A1:A2')
                ->mergeCells('B1:B2')
                ->mergeCells('C1:D1')
                ->mergeCells('E1:E2')
                ->mergeCells('F1:H1');
            $sheet->getStyle('A1:E2')->getFont()->setBold(true);
            $sheet->getStyle('F1:H2')->getFont()->getColor()->setARGB(\PHPExcel_Style_Color::COLOR_RED);
            $sheet->getRowDimension(1)->setRowHeight(18);
            $sheet->getRowDimension(2)->setRowHeight(18);
//            获取数据
            $res=db('yeji')->alias('a')
                ->join('user b','b.Id=a.uid')
                ->whereTime('a.uptime','between',[$time.'-1',$time."-".date('t',strtotime($time))])
                ->where('a.state','in',[1,2])
                ->where('b.state',1)
                ->field('b.name,b.phone,sum(a.yeji) as yeji,(b.ticheng*sum(a.yeji)/100) as ticheng,a.state,a.uid')
                ->group('b.name')
                ->select();
            $currow=0;
             foreach ($res as $key=>$v){
                 $currow=$key+3;
                 $sheet->setCellValue('A'.$currow,$v['name'])
                     ->setCellValue('B'.$currow,$v['phone'])
                     ->setCellValue('C'.$currow,$v['yeji'])
                     ->setCellValue('D'.$currow,$v['ticheng'])
                     ->setCellValue('E'.$currow,getstatecn($v['state']))
                     ->setCellValue('F'.$currow,'=C'.$currow)
                     ->setCellValue('G'.$currow,$this->huizong($v['uid'],$time,0))
                     ->setCellValue('H'.$currow,$this->huizong($v['uid'],$time,1));
             }
            $sheet->getStyle('A1:H'.$currow)->getBorders()->getAllBorders()->setBorderStyle(\PHPExcel_Style_Border::BORDER_THIN);
            $filename="月报".$time;
            $this->excelsave($phpexcel,$filename);
        }
        return view();
    }
    protected function excelsave($phpexcel,$filename){
        //创建一个Excel文档并下载保存
        $phpwriter=new \PHPExcel_Writer_EXCEL2007($phpexcel);
        header('Content-Type: application/vnd.ms-excel');//设置文档类型
        header('Content-Disposition: attachment;filename="'.$filename.'".xls"');//设置文件名
        header('Cache-Control: max-age=0');
        $phpwriter->save('php://output');
    }
    //汇总
    //$type=0季度  $type=1 年度
    protected function huizong($uid,$time,$type){
        if(!$type){
            $jidu=ceil(date('n',strtotime($time))/3);
            $firstmonth=($jidu-1)*3+1;
            $jidu_begin=date('Y',strtotime($time))."-".$firstmonth."-1";
        }else{
            $jidu_begin=date('Y',strtotime($time))."-1-1";
        }
        $jidu_end=$time."-".date('t',strtotime($time));
        $res=db('yeji')->where('uid',$uid)->where('state','in',[1,2])->whereTime('uptime','between',[$jidu_begin,$jidu_end])->sum('yeji');
        return $res;
    }
}