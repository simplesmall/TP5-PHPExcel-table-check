<?php
/**
 * Created by PhpStorm.
 * User: yifeng
 * Date: 2017/8/6
 * Time: 23:31
 */

namespace app\admin\controller;
use think\Controller;
class Common extends Controller
{
    public function _initialize()
    {
        $this->checklogin();
    }

    private function checklogin(){
       if(session('?loginname', '', 'admin')!=1 || session('?loginid', '', 'admin')!=1){
           $this->redirect('login/index');
       }
        //vendor('./phpoffice/Classes/PHPExcel');
    }
}