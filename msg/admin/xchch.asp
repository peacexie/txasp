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
msg = Request("msg")
if msg="" Then
  msg = "测试信息["&Now()&"]张长城小灵通短信接口（webservise版）"
end if
if act="deComp" Then
  act = "deCompress"
  smpOSet()
	mde = smcMObj.deCompress(msg)
  smpOEnd() 
else
  act = "Compress"
  smpOSet()
	mde = smcMObj.Compress(msg)
  smpOEnd() 
end if

%>
<div class="line15">&nbsp;</div>
<table width="480" border="0" align="center" cellpadding="3" cellspacing="1" class="tdbg" bgcolor="#CCCCCC">
  <tr>
    <td align="center"><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td align="center">&nbsp;</td>
          <th align="left"><!--#include file="../inc/inc_test.asp"--></th>
        </tr>
        <tr>
          <td width="30%" align="center"><%=act%></td>
          <td align="right"><a href="?">[重载]</a> - <a href="info.asp">[返回]</a></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td align="left"><table width="480" border="0" align="center">
        <form action="?" method="post" name="fmCode" id="fmCode">
          <tr>
            <td align="center">原信息</td>
            <td colspan="2"><textarea name="msg" cols="50" rows="8" id="msg"><%=msg%></textarea></td>
          </tr>
          <tr>
            <td align="center">处理后</td>
            <td colspan="2"><textarea name="enc" cols="50" rows="8" id="enc"><%=mde%></textarea></td>
          </tr>
          <tr>
            <td align="center">act</td>
            <td><select name="act" id="act">
                <option value="Compress">Compress</option>
                <option value="deComp">deCompress</option>
              </select>
              [<span id="len1"></span>,<span id="len2"></span>]</td>
            <td width="20%"><input type="submit" name="button" id="button" value="Submit" /></td>
          </tr>
        </form>
      </table></td>
  </tr>
  <tr>
    <td align="left">说明： 本页为chch接口编码解码与测试程序。<span id="my_msg"><br />
      注意： 正常运行的系统，请不要随便运行！否则后果自负！！！ </span></td>
  </tr>
  <tr>
    <td align="left"><div style="height:120px; overflow:scroll;">
        <%

a="1"
If a="1" Then
  Function tst01(xStr) 
	tst01 = "1.If."&xStr&b
  End Function
Else
  Function tst01(xStr) 
	tst01 = "2.Else."&xStr&b
  End Function
End If

Dim b : b = "(TestB)"

Response.Write "tst=("&tst01(Now)&")<br>"

Response.Write "<br>"&Now
Response.Write "<br>"&Time
dstr = Request("dstr")
send = Request("send")
If send="send" Then
  smpOSet()
    Response.Write "<hr>"&dstr
	Response.Write "<hr>"&smcMObj.deCompress(dstr)
  smpOEnd() 
End If

Response.Write "<br> --- End"
%>
        <%

a = smsBalance()
r1 =  a(0)
r2 =  a(1)
r3 =  a(2)
Response.Write "<br>:"&r1&" - "&r2&" - "&r3
'Response.End()

%>
      </div></td>
  </tr>
</table>
<script type="text/javascript">
var fmEnc = document.fmCode;
function chkLens(){
  len1.innerHTML = ""+fmEnc.msg.value.length;
  len2.innerHTML = ""+fmEnc.enc.value.length;
}
chkLens();
</script>
</body>
</html>
