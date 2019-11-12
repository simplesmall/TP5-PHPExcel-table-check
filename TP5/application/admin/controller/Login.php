<?php
/**
 * Created by PhpStorm.
 * User: yifeng
 * Date: 2017/8/6
 * Time: 15:46
 */
namespace app\admin\controller;
use think\Controller;
use think\Session;

class Login extends Controller
{
    public function index(){
        if(session('?loginname', '', 'admin')!=1 || session('?loginid', '', 'admin')!=1){
            return view();
        }
        $this->redirect('index/index');
    }
    public function goout(){
        session(null, 'admin');
        $this->success("退出成功",'login/index');
    }
    public function login(){
        $data=input('post.');
        $validate = validate('Manager');
        if(!$validate->scene('login')->check($data)){
            $this->error($validate->getError(),null,null,2);
        }
        $result=db('manager')->where('username',$data['username'])->field('Id,username,password')->find();
        if(!$result){
            $this->error("用户名不存在");
        }
        if(md5(trim($data['password']))!=$result['password']){
            $this->error("密码输入不正确");
        }
        session('loginname', $result['username'], 'admin');
        session('loginid', $result['Id'], 'admin');
        Session::set('username',$result['username']);
        db('manager')->where('Id',$result['Id'])->update(['logintime' => time()]);
        $this->success('登录成功','index/index');
        return;
    }
}