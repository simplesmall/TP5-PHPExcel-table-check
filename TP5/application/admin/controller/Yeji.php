<?php
/**
 * Created by PhpStorm.
 * User: yifeng
 * Date: 2017/8/17
 * Time: 17:51
 */

namespace app\admin\controller;


class Yeji extends Common
{
    public function index(){
        $res=db('yeji')->alias('a')
            ->join('user b','a.uid=b.Id')
            ->order('a.uptime Desc')
            ->field('a.*,b.name,b.phone')
            ->paginate(6);
        $this->assign('yeji',$res);
        return view();
    }

    public function lst(){
        $res=db('yeji')->where('state',0)->order('uptime Desc')->paginate(6);
        $this->assign('yeji',$res);
        return view();
    }
    public function gzlst(){
        //$res=null;
        if(request()->isPost()){
            $data=input('post.');
           if($data['time']!=''){
               $res=db('yeji')->alias('a')
                   ->join('user b','a.uid=b.Id')
                   ->group('a.uid')
                   ->where('a.state','in',[1,2])
                   ->whereTime('settime', 'between', [$data['time'].'-1', $data['time'].'-31'])
                   ->order('a.uptime Desc')
                   ->field('a.Id,a.uid,sum(a.yeji) as yejihe,b.name,b.phone,sum(a.yeji)*b.ticheng/100 as gongzi,a.settime,a.state')
                   ->paginate(6);
           }
        }
        if(!isset($res)){
            $res=db('yeji')->alias('a')
                ->join('user b','a.uid=b.Id')
                ->group('a.uid')
                ->where('a.state','in',[1,2])
                ->whereTime('settime', 'month')
                ->order('a.uptime Desc')
                ->field('a.Id,a.uid,sum(a.yeji) as yejihe,b.name,b.phone,sum(a.yeji)*b.ticheng/100 as gongzi,a.settime,a.state')
                ->paginate(6);
        }
        $this->assign('yeji',$res);
        return view();
    }
    public function tongguo(){
        $id=input('id');
        $state=input('sta');
        switch ($state){
            case 1:
                $sta=1;
                break;
            case 0:
                $sta=3;
                break;
            case 2:
                $sta=2;
                break;
            default:
                $sta=3;
                break;
        }
        if(!db('yeji')->where('Id',$id)->update(['state'=>$sta,'settime'=>time()])){
            $this->error("操作失败！");
        }
        $this->success("操作成功",'yeji/lst');
    }
    public function fafang(){
        $data['uid']=input('uid');
        $data['time']=date('Y-m',input('time'));
        $result=db('yeji')
            ->where('uid',$data['uid'])
            ->whereTime('settime', 'between', [$data['time'].'-1', $data['time'].'-31'])
            ->where('state',1)
            ->update(['state'=>2]);
        if($result){
            $this->success("工资发放成功！",'yeji/gzlst',null,1);
        }else{
            $this->error("工资发放失败！");
        }

        return;
    }
}