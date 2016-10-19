<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<!--#include file="../config/sms_pub.asp"-->
<!--#include file="../config/sms_do.asp"-->
<!--#include file="../config/sms_cfg.asp"-->
<%

Function apiList()
  dim fs,dir,f,s :s=""
  set fs = CreateObject("Scripting.FileSystemObject")
  set dir = fs.GetFolder(Server.MapPath("../config/") )
  for each f in dir.Files
	If Left(f.Name,4)="api_" Then
      s = s&vbCrLf&("<option value='" &f.Name& "'>" &f.Name& "</option>")
	End If
  next
  set dir = nothing
  set fs = nothing
  apiList = s
End Function

Call Chk_Perm1("ModSms","")

Dim debugAPI,debugRole,debugUser,debugAct,debugStr
debugAPI  = Request("debugAPI") :If debugAPI="" Then debugAPI="api_test.asp"
debugRole = RequestS("debugRole","C",24) :If debugRole="" Then debugRole="(SmsAPI)" '(SmsAPI),Member
debugUser = RequestS("debugUser","C",48) :If debugUser="" Then debugUser="demo" '"demo" 
debugAct  = Request("debugAct") :If debugAct="" Then debugAct="Info"
'Response.Write debugAPI

debugStr = File_Read("../config/"&debugAPI,"utf-8")
debugStr = Replace(Replace(debugStr,"<"&"%",""),"%"&">","")
'Response.Write Show_Text(debugStr)
execute(debugStr) 'eval

Dim res,act,actName,msg,tNumb,tCont
Dim sndType,sndUser,sndMaxs

sndType = debugRole '(SmsAPI),Admin(不限制),Inner,Message,Member(SmsMember)
sndUser = debugUser '
sndMaxs = doBalance(sndType,sndUser,-1) 

If debugAct="Balance" Then
  res = smsBalance()
  reB = rs_Val(conn,"SELECT SUM(MemBalance)AS MemCharge FROM [SmsMember]")
  reC = res(1) - reB
  actName = "接口余额"
  msg = res(2)&" "&Now()

ElseIf debugAct="Test" Then
  actName = "发送测试"
  tCont = ""&Time()&"测试@"&Get_CIP()&"("&Session("UsrID")&")"
  msg = "已经输入["&Len(tCont)&"]个字"
  act = debugAct
ElseIf debugAct="doTest" Then 
  actName = "发送测试"
  res = doTest()
  msg = res(2)&" "&Now()
  act = debugAct
ElseIf debugAct="Send" Then  'act=""
  actName = "发送短信"
  act = debugAct
ElseIf debugAct="doSend" Then 
  actName = "发送短信"
  msg = doSend(sndType,sndUser,"")
  act = debugAct

Else
  actName = "接口Info"
End If



%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>debgu</title>
<link rel="stylesheet" type="text/css" href="../inc/spub.css"/>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script src="../inc/send_check.js" type="text/javascript"></script>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<style type="text/css">
<!--
.rGap {
	text-align:right;
	padding-right:90px;
}
-->
</style>
</head>
<body>
<div class="line15">&nbsp;</div>
<table width="480" border="0" align="center" cellpadding="3" cellspacing="1" class="tdbg" bgcolor="#CCCCCC">
  <tr>
    <td colspan="2" align="center"><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td width="30%" align="center"><strong>Debug Admin</strong></td>
          <td align="right"><%=debugAPI%>, <%=debugRole%>, <%=debugUser%>, <%=debugAct%></td>
        </tr>
      </table></td>
  </tr>
  <form action="?" method="post" name="fmd01" id="fmd01">
    <tr>
      <td width="30%" align="right">debugAPI</td>
      <td align="left"><select name="debugAPI" id="debugAPI" style="width:150px">
          <option value="<%=debugAPI%>"><%=debugAPI%> (Now)</option>
          <%=apiList()%>
        </select></td>
    </tr>
    <tr>
      <td align="right">debugRole</td>
      <td align="left"><select name="debugRole" id="debugRole" style="width:150px">
          <%=Get_SOpt("(SmsAPI);Member","(SmsAPI)身份;(Member)身份",debugRole,"")%>
        </select></td>
    </tr>
    <tr>
      <td align="right">debugUser</td>
      <td align="left"><input name="debugUser" type="text" id="debugUser" value="demo" size="24" maxlength="36" />
        (demo)</td>
    </tr>
    <tr>
      <td width="30%" align="right">debugAct</td>
      <td align="left"><select name="debugAct" id="debugAct" style="width:150px">
          <%=Get_SOpt("Info;Balance;Test;Send","(Info)接口信息;(Balance)接口余额;(Test)发送测试;(Send)发送短信",Replace(debugAct,"do",""),"")%>
        </select></td>
    </tr>
    <tr>
      <td align="center"><a href="?">重载</a></td>
      <td align="left"><input name="Submit" type="submit" class="btn60" value="Debug" />
        <input name="Button3" type="reset" class="btn60" value="Reset" /></td>
    </tr>
  </form>
  <tr>
    <td colspan="2" align="left">说明：</td>
  </tr>
</table>
<div class="line15">&nbsp;</div>
<%If inStr(debugAct,"Test")>0 Then%>
<!--#include file="../inc/send_test.asp"-->
<%ElseIf inStr(debugAct,"Send")>0 Then%>
<!--#include file="../inc/send_form.asp"-->
<%Else%>
<table width="480" border="0" align="center" cellpadding="3" cellspacing="1" class="tdbg" bgcolor="#CCCCCC">
  <tr>
    <td colspan="2" align="center"><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td width="30%" align="center"><strong><%=actName%></strong></td>
          <td align="right"><%=debugAct&"@"&debugAPI%></td>
        </tr>
      </table></td>
  </tr>
  <%If debugAct="Info" Then%>
  <tr>
    <td width="30%" align="right">smcMClass</td>
    <td><%=smcMClass%></td>
  </tr>
  <tr>
    <td align="right">smcUser</td>
    <td><%=smcUser%></td>
  </tr>
  <tr>
    <td align="right">smcUEID</td>
    <td nowrap="nowrap"><%=smcUEID%></td>
  </tr>
  <tr>
    <td width="30%" align="right">smcPass</td>
    <td><%=smcPass%></td>
  </tr>
  <tr>
    <td align="right">smcPase</td>
    <td><%=smcPase%></td>
  </tr>
  <tr>
    <td align="right">smcTSplit</td>
    <td nowrap="nowrap"><%=smcTSplit%></td>
  </tr>
  <tr>
    <td align="right">smcTMax1</td>
    <td><%=smcTMax1%></td>
  </tr>
  <tr>
    <td align="right">smcTMax2</td>
    <td><%=smcTMax2%></td>
  </tr>
  <tr>
    <td align="right">smcCLong</td>
    <td><%=smcCLong%></td>
  </tr>
  <tr>
    <td align="right">smcCMul</td>
    <td><%=smcCMul%></td>
  </tr>
  <tr>
    <td align="right">smcDMod</td>
    <td><%=smcDMod%></td>
  </tr>
  <tr>
    <td align="right">smcDTabs</td>
    <td nowrap="nowrap"><%=smcDTabs%></td>
  </tr>
  <%End If%>
  <%If debugAct="Balance" Then%>
  <tr>
    <td width="30%" align="right">系统余额(条)</td>
    <td class="rGap"><%=res(1)%> (条)</td>
  </tr>
  <tr>
    <td align="right">客户余额(条)</td>
    <td class="rGap"><%=reB%> (条)</td>
  </tr>
  <tr>
    <td align="right">可 充 值(条)</td>
    <td nowrap="nowrap" class="rGap"><%=res(1)%> - <%=reB%> = <%=reC%> (条)</td>
  </tr>
  <%End If%>
  <tr>
    <td colspan="2" align="left">说明：</td>
  </tr>
</table>
<%End If%>
<div class="line15">&nbsp;</div>
<%
If debugAct="Test" Or debugAct="Send" Then
  debugAct = "do"&debugAct
End If
%>
<script type="text/javascript">
function api_debug_set(){
  var sActs = "debugAPI,debugRole,debugUser,debugAct";
  var sVals = "<%=debugAPI%>,<%=debugRole%>,<%=debugUser%>,<%=debugAct%>";
  var aActs = sActs.split(",");
  var aVals = sVals.split(",");
  var sActs = "";
  for(i=0;i<aActs.length;i++){
	  sActs += "<input name='"+aActs[i]+"' type='hidden' id='"+aActs[i]+"' value='"+aVals[i]+"' />";
  }
  try{document.getElementById("api_debug_area").innerHTML = sActs;}catch(e){}
}
api_debug_set();
</script>
</body>
</html>
