<?php
/**
 * Created by PhpStorm.
 * User: yifeng
 * Date: 2017/8/11
 * Time: 15:53
 */

namespace app\common\validate;
use think\Validate;

class User extends Validate
{
    protected $rule =   [
        'name'  => 'require|max:20',
        'ticheng'   => 'number|between:0,40',
        'phone'=>'require|length:11',
        'password'=>'require|min:6',
        'code'=>'require|captcha',
    ];

    protected $message  =   [
        'name.require' => '姓名不能为空',
        'name.max'     => '姓名最多不能超过20个字符',
        'ticheng.number'   => '提成比例必须是数字',
        'ticheng.between'  => '提成比例必须在：0%-40%',
        'phone.require'=>'手机号不能为空！',
        'phone.length'=>'请输入正确的手机号',
        'password.require'=>"密码不能为空",
        'password.min'=>"密码长度不能小于6位",
        'code.require'=>'验证码不能为空',
        'code.captcha'=>'验证码不正确'
    ];

    protected $scene = [
        'add'  =>  ['name','ticheng'],
        'edit'  =>  ['name','ticheng'],
        'login'=>['phone','password','code'],
    ];
}