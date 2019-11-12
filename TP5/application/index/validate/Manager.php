<?php
/**
 * Created by PhpStorm.
 * User: yifeng
 * Date: 2017/8/6
 * Time: 21:27
 */

namespace app\admin\validate;
use think\Validate;

class Manager extends Validate
{
    protected $rule =   [
        'username'  => 'require|max:25',
        'password'   => 'require|min:6',
        'oldpassword'=> 'require|min:6',
        'repassword'=>'confirm:password'
    ];

    protected $message  =   [
        'username.require' => '用户名不能为空',
        'username.max'     => '用户名最多不能超过25个字符',
        'password.require'   => '密码不能为空',
        'password.min'  => '密码长度不能少于6位',
        'oldpassword.require'=>'旧密码不能为空',
        'oldpassword.min'=>'旧密码长度不能少于6位',
        'repassword.confirm'=>'两次密码输入不一致',
    ];
    protected $scene = [
        'login'=>['username','password'],
        'modify'  =>  ['password','oldpassword','repassword'],
    ];

}