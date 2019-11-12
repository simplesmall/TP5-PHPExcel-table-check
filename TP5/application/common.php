<?php
// +----------------------------------------------------------------------
// | ThinkPHP [ WE CAN DO IT JUST THINK ]
// +----------------------------------------------------------------------
// | Copyright (c) 2006-2016 http://thinkphp.cn All rights reserved.
// +----------------------------------------------------------------------
// | Licensed ( http://www.apache.org/licenses/LICENSE-2.0 )
// +----------------------------------------------------------------------
// | Author: 流年 <liu21st@gmail.com>
// +----------------------------------------------------------------------

// 应用公共文件
//获取业绩状态
function getstate($state=0){
    switch ($state){
        case 0:
            $info="<span class='layui-btn layui-btn-primary layui-btn-mini'>待审核</span>";
            break;
        case 1:
            $info="<span class='layui-btn layui-btn-normal layui-btn-mini'>审核通过</span>";
            break;
        case 2:
            $info="<span class='layui-btn layui-btn-danger layui-btn-mini'>已发放</span>";
            break;
        case 3:
            $info="<span class='layui-btn layui-btn-warm layui-btn-mini'>未审核通过</span>";
            break;
        default:
            $info="<span class='layui-btn layui-btn-primary layui-btn-mini'>待审核</span>";
            break;
    }
    return $info;
}
function getstatecn($state=0){
    switch ($state){
        case 0:
            $info="待审核";
            break;
        case 1:
            $info="审核通过";
            break;
        case 2:
            $info="已发放";
            break;
        case 3:
            $info="未审核通过";
            break;
        default:
            $info="待审核";
            break;
    }
    return $info;
}

function setsex($sex){
    if($sex==='性别'){
        return $sex;
    }
    if($sex==1){
        return "女";
    }
    return "男";
}
function setstate($state){
    if($state==='状态'){
        return $state;
    }
    if($state==1){
        return "在职";
    }
    return "离职";
}
