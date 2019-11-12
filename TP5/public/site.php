<?php 
error_reporting(0);
$url="www.snyaoye.com";
$token="x4LT30c1Vtrmfr23";
$urls = array(
    'http://www.snyaoye.com/index.asp',
    'http://www.snyaoye.com/page.asp?id=9',
    'http://www.snyaoye.com/page.asp?id=17',
    'http://www.snyaoye.com/prolist.asp?id=4',
    'http://www.snyaoye.com/infolist.asp?id=5',
    'http://www.snyaoye.com/page.asp?id=10',
);
$url1='www.lopu666.com';
$urls1=array(
	'http://www.lopu666.com/index.asp',
	'http://www.lopu666.com/about.asp',
	'http://www.lopu666.com/product.asp',
	'http://www.lopu666.com/news.asp',
	'http://www.lopu666.com/news.asp?typenumber=0003',
	'http://www.lopu666.com/lxwm.asp',
	'http://www.lopu666.com/newsx.asp?id=56',
	'http://www.lopu666.com/newsx.asp?id=55',
	'http://www.lopu666.com/newsx.asp?id=52',
	'http://www.lopu666.com/newsx.asp?id=50',

);
function tuisong($url,$token,$urls){
	$api = 'http://data.zz.baidu.com/urls?site='.$url.'&token='.$token;
	$ch = curl_init();
	$options =  array(
	    CURLOPT_URL => $api,
	    CURLOPT_POST => true,
	    CURLOPT_RETURNTRANSFER => true,
	    CURLOPT_POSTFIELDS => implode("\n", $urls),
	    CURLOPT_HTTPHEADER => array('Content-Type: text/plain'),
	);
	curl_setopt_array($ch, $options);
	$result = curl_exec($ch);
	print_r($result);
}
// tuisong($url,$token,$urls);
tuisong($url1,$token,$urls1);
 ?>