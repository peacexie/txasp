<?php
require("../../upfile/sys/config/config.php");
require("../../sadm/func1/func_time.php");
require("../../sadm/func1/func_php.php");
require("../../sadm/func1/md5_func.php");
require("form_app.php");

$app99Tim = date("Y-m-d H:i:s");
$app30Arr = App30Set($app99Tim,"s","");
$app30Tab = explode("-",$app30Arr[2]);

$urlEvent = urlencode('class="XlgnBtn1" Xtabindex=1 onblur="ChkAjMem(\'ChkAjID\');"');
$url30par = "&".$App30Code."0=$app30Arr[0]&".$App30Code."1=$app30Arr[1]"; //&".$App30Code."2=$app30Arr[2]
//$urlparas = "?act=setInput&n=".Get_IDEnc($app30Tab[1])."&v=$MemID&s=24&m=64&e=$urlEvent$url30par";




// .............

?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Test form_app.php</title>
<link rel="stylesheet" type="text/css" href="test_app.css"/>
</head>
<body>



<?php 
$id = $app30Tab[6];
$a = App32Show("$id","R");
echo "\r\n<br>App32Show(R)  随机1/6 $a";
$a = App32Show("$id","A");
echo "\r\n<br>App32Show(A)  随机A1/3 $a";
$a = App32Show("$id","B");
echo "\r\n<br>App32Show(B)  随机B1/3 $a";
$a = App32Show("$id","A1");
echo "\r\n<br>App32Show(A1)  固定A1 $a";
?>


<?php 
///////////////////////////////////////////////////////////////////////////////////
echo "\r\n <hr><hr> ";
///////////////////////////////////////////////////////////////////////////////////
echo "\r\n <hr> 得到Len个字符/中文";

echo "\r\n<br>App32GetNX()  ◆ ".App32GetNX();
echo "\r\n<br>App32GetNX(0) ◆ ".App32GetNX("0");
echo "\r\n<br>App32GetNX(A) ◆ ".App32GetNX("A");
echo "\r\n<br>App32GetNX(a) ◆ ".App32GetNX("a");
echo "\r\n<br>App32GetNX(S) ◆ ".App32GetNX("S");
echo "\r\n<br>App32GetNC()  ◆ ".App32GetNC();


///////////////////////////////////////////////////////////////////////////////////
echo "\r\n <hr><hr> ";
///////////////////////////////////////////////////////////////////////////////////


$str = "PeaceABC123和平DEF";
echo "\r\n<br>App32GetNC()  ★ ".App32CUCS($str, 'UTF-8');

echo "\r\n <hr> ";

$m1 = 10102; //1234
$m2 = 10131;
echo "\r\n <br> ";
for($i=$m1;$i<=$m2;$i++){
  //echo " &#$i; ";
}

$m1 = 192; // À Á Ù Ú
$m2 = 223;
echo "\r\n <br> ";
for($i=$m1;$i<=$m2;$i++){
  echo " &#$i; ";
}

$m1 = 224; // à á ù ú
$m2 = 255;
echo "\r\n <br> ";
for($i=$m1;$i<=$m2;$i++){
  echo " &#$i; ";
}


//echo App32GetNC(64)

?>
</body>
</html>