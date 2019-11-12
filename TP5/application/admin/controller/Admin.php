<?php

namespace app\admin\controller;

use think\Controller;
use think\Request;

class Admin extends Controller
{
    /**
     * 显示资源列表
     *
     * @return \think\Response
     */
    public function index()
    {
        $admin = db('manager')->order('Id')->paginate(15);
        $this->assign('admin',$admin);
        return view();
    }

    /**
     * 显示创建资源表单页.
     *
     * @return \think\Response
     */
    public function add()
    {
        if (request()->isPost())
        {
            $data= input('post.');

            $validate = validate('Admin');
            if (!$validate->check($data))
            {
                $this->error($validate->getError());
            }

            $data['password']=md5(input('post.password'));

//            添加校验用户名已存在
            $dbname=db('manager')->where('username',$data['username'])->select();
            if($dbname)
            {
                $this->error("该用户名已存在");
            }
            $res = db('manager')->insert($data);


            if (!$res){
                $this->error('添加失败');
            }
            $this->success("添加成功",'admin/index');
        }
        return view();
    }

    /**
     * 保存新建的资源
     *
     * @param  \think\Request  $request
     * @return \think\Response
     */
    public function save(Request $request)
    {
        //
    }

    /**
     * 显示指定的资源
     *
     * @param  int  $id
     * @return \think\Response
     */
    public function read($id)
    {
        //
    }

    /**
     * 显示编辑资源表单页.
     *
     * @param  int  $id
     * @return \think\Response
     */
    public function edit($id)
    {
        //
    }

    /**
     * 保存更新的资源
     *
     * @param  \think\Request  $request
     * @param  int  $id
     * @return \think\Response
     */
    public function update(Request $request, $id)
    {
        //
    }

    /**
     * 删除指定资源
     *
     * @param  int  $id
     * @return \think\Response
     */
    //    删除
    public function del()
    {
        $id = input('id');

        if($id==1)
        {
            $this->error("超级管理用户不可删除");
            return;
        }
        if (!db('manager')->delete($id)) {
            $this->error('删除失败');
        }
        $this->success("删除成功", 'admin/index');
        return;
    }
}
