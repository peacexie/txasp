<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>发送短信</title>
<link rel="stylesheet" type="text/css" href="../inc/spub.css"/>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script src="../inc/send_check.js" type="text/javascript"></script>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
</head>
<body>
<%
sndMaxs = uBalance

tm = Now() 'id = id
sn = outEncSN(id,tm,uCode,uUrl)

tNumb = Request("tNumb")
tCont = Request("tCont") 

reUrl = "?tm="&tm&"&id="&id&"&sn="&sn&"&"&sn&"="&uUrl&"&act=SendOut"

%>
<table width="460" border="0" align="center" cellpadding="3" cellspacing="1" class="tdbg" bgcolor="#F0F0F0">
  <form action="?" method="post" name="fm01API" id="fm01API">
    <%If act="SendODo" Then%>
    <tr>
      <td align="center">发送报告</td>
      <td><%=msg%><br />
        <a href="<%=reUrl&"&"&Rnd_ID("",8)%>=<%=Rnd_ID("",8)%>" class="cRed">请返回&gt;&gt;</a> <%=Now()%></td>
    </tr>
    <%Else%>
    <tr>
      <td width="15%" align="center">余额</td>
      <td><input name="tBalance" type="text" id="tBalance" value="<%=sndMaxs%>" size="8" readonly="readonly" style="color:#CCC" />
        条(<%=smcDMod%>), 自动获得,不能修改；</td>
    </tr>
    <tr>
      <td align="center">手机号码<br />
        (号码表)</td>
      <td><textarea name="tNumb" cols="36" rows="5" id="tNumb"><%=tNumb%></textarea>
        <div> 最多<%=smcOTMax%>个号码，多个手机号，一行一个或用半角[<span class="fntF00 fB f14px"><%=smcTSplit%></span>]隔开。 </div></td>
    </tr>
    <tr>
      <td align="center">短信内容</td>
      <td><textarea name="tCont1" cols="36" rows="4" id="tCont1" onblur="chkChars(this,'nowChars',70)"><%=tCont%></textarea></td>
    </tr>
    <tr>
      <td><input name="act" type="hidden" id="act" value="SendODo" />
        <input name="ChkCode" type="hidden" id="ChkCode" value="<%=Session("ChkCode")%>" /></td>
      <td><input name="send" type="button" value="发送短信" onclick="chkSOut()" class="btn60" />
        <input name="reset" type="reset" value="重填资料" class="btn60" />
        <a href="<%=reUrl&"&"&Rnd_ID("",8)%>=<%=Rnd_ID("",8)%>">重载
        <input name="tel" type="hidden" id="tel" value="<%=Session("ChkCode")%>" />
        <input name="ct1" type="hidden" id="ct1" value="<%=Session("ChkCode")%>" />
        </a></td>
    </tr>
    <tr>
      <td colspan="2" align="left">说明/提示：<span id="nowChars" class="fntF00"><%=msg%></span>
      </td>
    </tr>
    <%End If%>
    <input name="id" type="hidden" id="id" value="<%=id%>" />
    <input name="tm" type="hidden" id="tm" value="<%=tm%>" />
    <input name="sn" type="hidden" id="sn" value="<%=sn%>" />
    <input name="<%=sn%>" type="hidden" id="<%=sn%>" value="<%=uUrl%>" />
  </form>
</table>
<script type="text/javascript">

var fm = document.fm01API; 
var telMaxs="<%=smcTMax1%>";
var keyCont="<%=ParFilALLKeys%>";

</script>
</body>
</html>
