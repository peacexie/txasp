
----------------------------------------------------------------
        短信接口API配置说明
----------------------------------------------------------------


  ****** 01. 接口信息：

接口帐号：(id)
接口SN：  (sn)
接口地址：(ad1)
发送地址：(ad2)
网管号码：(tel)

以上信息，不含空格，区分大小写。


  ****** 02. 注意事项：
  
A. 请妥善保管好[接口帐号]和[接口SN]等敏感信息；
B. 请设置好发送页的访问权限。


  ****** 03. 调用公共参数：
  
<!--#include file="(path)/smsmd5.asp"-->
<!--#include file="(path)/smscfg.asp"-->
<%
tm = Now()
sn = outEncSN(id,tm,snOrg) 
urlAdd0 = smaAdda '(...&"msg/user/sapi.asp")
urlPar0 = "?tm="&tm&"&id="&id&"&sn="&sn&"" 'Peace_Sms_RndID="&Timer&"&
'以上公共参数
tel = smaTels '接收号码,请根据需要 自行设置
ct1 = ""&Time()&"测试@domain.com" '信息内容,请根据需要 自行设置
%>


  ****** 04. smsapi.asp 简单调用：

<%
urlPara = urlPar0&"&act=Send&tel="&tel&"&ct1="&Server.URLEncode(ct1)&"" 
'echo tel&ct1 :Response.End()
Response.Redirect urlAdd0&urlPara
%>
[IFRAME name='SmsFrame' src="<%=urlAdd0&urlPara%>" frameBorder=0 width="30" scrolling=no height="15"][/IFRAME]


  ****** 05. smsapi.asp Form提交：

<form name="fmSms" action="<%=urlAdd0%>" method="post" target="_self">
  <input type="hidden" value="<%=tm%>" name="tm" />
  <input type="hidden" value="<%=id%>" name="id" />
  <input type="hidden" value="<%=dn%>" name="sn" />
  <input type="hidden" value="Send" name="act" />
  <input type="hidden" value="<%=tel%>" name="tel" />
  <input type="hidden" value="<%=ct1%>" name="ct1" />
</form>
<script type="text/javascript">
document.fmSms.submit();
</script>

  ****** 06. php编码函数：

<?php 
function outEncode($str, $encoding='utf-8', $prefix='&#', $postfix=';'){
    $str = iconv($encoding, 'UCS-2', $str);
    $arrstr = str_split($str, 2);
    $unistr = '';
    for($i=0, $len=count($arrstr); $i<$len; $i++)
    {
        $dec = hexdec(bin2hex($arrstr[$i]));
        //if($dec<0) {$dec+=65536;}
        $unistr .= $prefix.$dec.$postfix;
    }
    return $unistr;
}
function outEncSN($xid,$xtm,$xsn){
	// using System.Security.Cryptography;
	$snOrg = "";
	$smaAddr = "";
	$str = "$xid+$snOrg+$xtm+$smaAddr";
	return md5($str);
}
$sNow = date('Y-m-d H:i:s');
$sTest = "测试byPeace鸽子!";
echo outEncode($sTest,'utf-8','')."<br>";
echo outEncSN("demo","2011-4-25 11:11:13","0BA9A93D-7304-00CA-0F79-B25FAB273C1B");
?>

  ****** 07. C#编码函数：
  
public string outEncode(string xStr)
{
    string s=""; int iAsc;
    char[] aStr = xStr.ToCharArray(0, xStr.Length);
    for (int i=0;i<aStr.Length;i++)
    {
        iAsc = Convert.ToInt32(aStr[i]);
        s += iAsc+";";
    }
    return s;
}
public string outEncSN(string xid,string xtm,string xsn)
{
    // using System.Security.Cryptography;
    string snOrg = "";
    string smaAddr = "";
    string str = xid+"+"+snOrg+"+"+xtm+"+"+smaAddr;
    str = FormsAuthentication.HashPasswordForStoringInConfigFile(str, "md5");
    return str.ToLower();
}
string sNow = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
//参考文件smsapi.aspx.cs
--- End --------------------------------------------------------
