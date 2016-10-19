<!--#include file="_config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>接口激活/测试/调试/编码</title>
<link rel="stylesheet" type="text/css" href="../inc/spub.css" />
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>
<%
act = Request("act")
if act="" Then
  act = "SDK-SN"
elseif act="xeureg" Then
  Call smsRegID()
elseif act="xeuout" Then
  Call smsRegUn()
end if
msg = Request("msg")
if msg="" Then
  msg = "测试信息["&Now()&"]亿美软通短信平台SDK4.1.0（HTTP版）"
end if


%>
<div class="line15">&nbsp;</div>
<table width="480" border="0" align="center" cellpadding="3" cellspacing="1" class="tdbg" bgcolor="#CCCCCC">
  <tr>
    <td align="center"><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td width="30%" rowspan="2" align="center"><%=act%></td>
          <th align="left"><!--#include file="../inc/inc_test.asp"--></th>
        </tr>
        <tr>
          <td align="right"><a href="?">[重载]</a> - <a href="?">序列号测试</a> - <a href="?act=Encode">编码测试</a> - <a href="info.asp">[返回]</a></td>
        </tr>
      </table></td>
  </tr>
  <%If act="SDK-SN" Then%>
  <tr>
    <td align="left"><table width="480" border="0" align="center">
        <form action="?" method="post" name="fmReg" target="_blank" id="fmReg">
          <tr>
            <td align="center">my_url</td>
            <td><input name="my_url" type="text" id="my_url" value="http://sdkhttp.eucp.b2m.cn/sdkproxy/" size="36" /></td>
            <td>.action</td>
          </tr>
          <tr>
            <td align="center">cdkey</td>
            <td><input name="cdkey" type="text" id="cdkey" size="36" /></td>
            <td>用户序列号</td>
          </tr>
          <tr>
            <td align="center">password</td>
            <td><input name="password" type="text" id="password" value="" size="36" /></td>
            <td>用户密码</td>
          </tr>
          <tr>
            <td align="center">ename</td>
            <td><input name="ename" type="text" id="ename" value="科成信息科技" size="36" /></td>
            <td>企业名称</td>
          </tr>
          <tr>
            <td align="center">linkman</td>
            <td><input name="linkman" type="text" id="linkman" value="谢永顺" size="36" /></td>
            <td>联系人姓名</td>
          </tr>
          <tr>
            <td align="center">phonenum</td>
            <td><input name="phonenum" type="text" id="phonenum" value="0769-2202-8868" size="36" /></td>
            <td>联系电话</td>
          </tr>
          <tr>
            <td align="center">mobile</td>
            <td><input name="mobile" type="text" id="mobile" value="135-3743-2147" size="36" /></td>
            <td>联系手机</td>
          </tr>
          <tr>
            <td align="center">email</td>
            <td><input name="email" type="text" id="email" value="xpigeon@163.com" size="36" /></td>
            <td>电子邮件</td>
          </tr>
          <tr>
            <td align="center">fax</td>
            <td><input name="fax" type="text" id="fax" value="0769-2202-8800" size="36" /></td>
            <td>联系传真</td>
          </tr>
          <tr>
            <td align="center">address</td>
            <td><input name="address" type="text" id="address" value="东莞市莞太大道34号软件孵化园2#3D" size="36" /></td>
            <td>公司地址</td>
          </tr>
          <tr>
            <td align="center">postcode</td>
            <td><input name="postcode" type="text" id="postcode" value="523000" size="36" /></td>
            <td>邮政编码</td>
          </tr>
          <tr>
            <td align="center">my_act</td>
            <td><select name="my_act" id="my_act">
              <option value="regist"          >11. 注册序列号.regist</option>
              <option value="logout"          >12. 注销序列号.logout</option>
              <option value="querybalance"    >21. 查询余额.querybalance</option>
              <option value="sendsms"         >22. 发送短信.sendsms</option>
              <option value="registdetailinfo">31. 企业信息.detailinfo</option>
            </select></td>
            <td><input type="button" name="my_button" id="my_button" value="Button" onclick="chkData()" /></td>
          </tr>
        </form>
      </table></td>
  </tr>
  <%Else%>
  <tr>
    <td align="left"><table width="480" border="0" align="center">
        <form action="?" method="post" name="fmCode" id="fmCode">
          <tr>
            <td align="center">信息</td>
            <td colspan="2"><textarea name="msg" cols="50" rows="8" id="msg"><%=msg%></textarea></td>
          </tr>
          <tr>
            <td align="center">编码</td>
            <td colspan="2"><textarea name="enc" cols="50" rows="8" id="enc"><%=Server.URLEncode(Request("msg"))%></textarea></td>
          </tr>
          <tr>
            <td align="center">act</td>
            <td><select name="act" id="act">
                <option value="Encode">Encode</option>
              </select>
              [<span id="len1"></span>,<span id="len2"></span>]</td>
            <td width="20%"><input type="submit" name="button" id="button" value="Submit" /></td>
          </tr>
        </form>
      </table></td>
  </tr>
  <%End If%>
  <tr>
    <td align="left">说明： 本页为亿美软通sdk(http) 序列号激活与测试程序。<span id="my_msg"><br />
      注意： 本页仅用于激活序列号和测试，正常运行的系统，请不要随便运行！否则后果自负！！！ </span></td>
  </tr>
  <tr>
    <td align="left">
<div style="height:120px; overflow:scroll;">
<%
smsRetCode = "-1;0;304;305;307;308;3;10;11;12;13;14;15;16;17;18;22;27"
smsRetName = "未知错误;成功;客户端发送三次失败;服务器返回了错误的数据，原因可能是通讯过程中有数据丢失;发送短信目标号码不符合规则，手机号码必须是以0、1开头;非数字错误，修改密码时如果新密码不是数字那么会报308错误;连接过多，指单个节点要求同时建立的连接数过多;客户端注册失败;企业信息注册失败;查询余额失败;充值失败;手机转移失败;手机扩展转移失败;取消转移失败;发送信息失败;发送定时信息失败;注销失败;查询单条短信费用错误码"
smsRetCode = "0;10;101;305;999;-1;11;307;22;13;17;18;997;998;308"
smsRetName = "操作成功;客户端注册失败;客户端网络故障;服务器端返回错误(返回值不是数字字符串);操作频繁;企业信息或密码不符合要求;企业信息注册失败;目标电话号码不符合规则(01开头);注销失败;充值失败;发送信息失败;发送定时信息失败;找不到超时的短信无法确定是否成功;客户端网络问题导致发送超时;新密码不是数字"

a1 = Split(smsRetCode,";")
a2 = Split(smsRetName,";")
for i=0 to uBound(a1)
  Response.Write ""&a1(i)&": "&a2(i)&"<br>"
next
%>
</div>    
    </td>
  </tr>

  
</table>
<script type="text/javascript">
<%If act="SDK-SN" Then%>
var fmObj = document.fmReg;
function chkData(){
   var eflag = 0;
   for(ii=0;ii<1;ii++){  ////////// //////////////// Srart For ////////////////
     if (fmObj.my_url.value.length==0){   
       alert("my_url is null"); 
       fmObj.my_url.focus(); eflag = 1; break;
     }
     var mob = fmObj.mobile.value;
     fmObj.mobile.value = mob.replace("-","").replace("-","");
   }  ////////// //////////////// End For //////////////////
   if (eflag==0){ 
	 fmObj.action = fmObj.my_url.value+""+fmObj.my_act.value+".action";
	 var fmPara = "?cdkey="+fmObj.cdkey.value+"&password="+fmObj.password.value+"";
	 my_msg.innerHTML = "<br>"+fmObj.action+fmPara;
	 fmObj.submit(); 
   }
}
<%Else%>
var fmEnc = document.fmCode;
function chkLens(){
  len1.innerHTML = ""+fmEnc.msg.value.length;
  len2.innerHTML = ""+fmEnc.enc.value.length;
}
chkLens();
<%End If%>
</script>
</body>
</html>
