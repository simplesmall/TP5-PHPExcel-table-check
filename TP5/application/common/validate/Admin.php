<?php
/**
 * Created by PhpStorm.
 * User: yifeng
 * Date: 2017/8/11
 * Time: 15:53
 */

namespace app\common\validate;
use think\Validate;

class Admin extends Validate
{
    protected $rule =   [
        'username'  => 'require|max:20',
        'password'   => 'require|min:6'
    ];

    protected $message  =   [
        'username.require' => '管理员名不能为空',
        'username.max'     => '姓名最多不能超过20个字符',
        'password.require' =>'密码不为空',
        'password.min'  => '密码长度不能少于6位',
    ];

    protected $scene = [
        'add'  =>  ['name','password'],
    ];
}