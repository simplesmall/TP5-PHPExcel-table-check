<?php
namespace app\admin\controller;
use think\Db;
class Index extends Common
{
    public function index()
    {
        $name = session('username');
        $this->assign('name',$name);
        return view();
    }

//   加载 密码修改界面
    public function modify(){
        if(request()->isPost()){
            $data=input('post.');
            $validate=validate("Manager");
            if(!$validate->scene('modify')->check($data)){
                $this->error($validate->getError());
            }
            $result=db('manager')->where("Id",session("loginid","",'admin'))->field("password")->find();
            if(md5($data['oldpassword'])!=$result['password']){
                $this->error("旧密码认证失败");
                reutnr;
            }
            $res=db('manager')->where("Id",session("loginid","",'admin'))->setField('password',md5($data['password']));;
            if($res){
                $this->success("密码更新成功");
            }else{
                $this->error("密码更新失败");
            }
            return;
        }
        return view();
    }
    //加载欢迎界面
    public function  welcome(){
            $server=[
                'HTTP_HOST'=>$_SERVER['HTTP_HOST'],
                'SERVER_SOFTWARE'=>$_SERVER['SERVER_SOFTWARE'],
                'osname'=>php_uname(),
                'HTTP_ACCEPT_LANGUAGE'=>$_SERVER['HTTP_ACCEPT_LANGUAGE'],
                'SERVER_PORT'=>$_SERVER['SERVER_PORT'],
                'SERVER_NAME'=>$_SERVER['SERVER_NAME'],
            ];
            $version=Db::query("select version()");
            $server['mysqlversion']=$version[0]['version()'];
            $server['databasename'] =config('database')['database'];
            $server['phpversion']=phpversion();
            $server['maxupload']=ini_get('max_file_uploads');
            $this->assign('server',$server);
            return view();
    }
}
