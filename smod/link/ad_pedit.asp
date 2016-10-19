<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/Func2.asp"-->
<!--#include file="../../sadm/func1/func_opt.asp"-->
<%

TP = RequestS("TP","C",48)
ID = RequestS("ID","C",48)
send = Request("send")

If send="send" Then
  InfType = RequestS("InfType","C",24)
  InfName = RequestS("InfName","C",24)
  InfPath = RequestS("InfPath","C",60)
  InfCont = "" :sUrl = "" :sPic = "" : sMsg=""
  For i=1 To 8
   sUrl = sUrl&RequestS("InfUrl"&i,"C",240)&"|"
   sPic = sPic&RequestS("InfPic"&i,"C",240)&"|"
   sMsg = sMsg&RequestS("InfMsg"&i,"C",240)&"|"
  Next
  InfCont = sUrl&"(^)"&sPic&"(^)"&sMsg
  InfPara = RequestS("InfPara1","N",120)&"|"&RequestS("InfPara2","N",60)&"|"&RequestS("InfPara3","N",3)

sql = " UPDATE [WebAdvert] SET " 
sql = sql& " InfType = '" & InfType &"'" 
sql = sql& ",InfName = '" & InfName &"'" 
sql = sql& ",InfCont = '" & InfCont &"'" 
sql = sql& ",InfPath = '" & InfPath &"'" 
sql = sql& ",InfPara = '" & InfPara &"'" 
sql = sql& ",LogEditIP = '" & Get_CIP() &"'" 
sql = sql& ",LogEUser = '" & Session("UsrID") &"'" 
sql = sql& ",LogETime = '" & Now() &"'" 
sql = sql& " WHERE KeyID='"&ID&"' "

  Call rs_DoSql(conn,sql) 
  Response.Write js_Alert("修改成功！","Redir","ad_pic.asp?TP="&TP&"") 

End If

SET rs=Server.CreateObject("Adodb.Recordset") 
rs.Open "SELECT * FROM [WebAdvert] WHERE KeyID='"&ID&"'",conn,1,1 
if NOT rs.eof then 
KeyID = rs("KeyID")
KeyMod = rs("KeyMod")
InfType = rs("InfType")
InfName = rs("InfName")
InfCont = rs("InfCont")
InfPath = rs("InfPath")
InfPara = rs("InfPara")
end if 
rs.Close()
SET rs=Nothing 

InfParaA = Split(InfPara,"|")
InfPara1=InfParaA(0) : InfPara2=InfParaA(1) : InfPara3=InfParaA(2) 

%>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
</head>
<body>
<!--#include file="config.asp"-->

<form action="?" method="post" name="fm01" id="fm01" style="margin:0px;">
  <br>
  <table width="640" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#CCCCCC" bordercolordark="#FFFFFF" style="margin-top:-2px;">
    <tr>
      <td width="12%" height="30" align="center" bgcolor="#EFEFEF">位置</td>
      <td align="left" bgcolor="#EFEFEF"><input name="InfName" id="InfName" value="<%=InfName%>" size="24" maxlength="12" />
        &nbsp;&nbsp;&nbsp;&nbsp;编号
<input name="InfType" id="InfType" value="<%=InfType%>" size="12" maxlength="12" /></td>
    </tr>
    <tr>
      <td height="30" align="center" bgcolor="#EFEFEF">宽高</td>
      <td align="left" bgcolor="#EFEFEF"><input name="InfPara1" id="InfPara1" value="<%=InfPara1%>" size="6" maxlength="4" />
        x
        <input name="InfPara2" id="InfPara2" value="<%=InfPara2%>" size="6" maxlength="4" />
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;速度
        <input name="InfPara3" id="InfPara3" value="<%=InfPara3%>" size="6" maxlength="6" />
        (多个图切换间隔:豪秒)</td>
    </tr>
    <tr>
      <td height="30" align="center" bgcolor="#EFEFEF">路径</td>
      <td align="left" bgcolor="#EFEFEF"><input name="InfPath" id="sUrl4" value="<%=InfPath%>" size="60" maxlength="60" /></td>
    </tr>
    <%
	aCont = Split(InfCont&"(^)(^)","(^)")
	aUrl = Split(aCont(0)&"||||||||||","|")
	aPic = Split(aCont(1)&"||||||||||","|")
	aMsg = Split(aCont(2)&"||||||||||","|")
	For i=1 To 8
	%>
    <tr>
      <td height="30" align="center" bgcolor="#FFFFFF">项目<%=i%></td>
      <td align="left" bgcolor="#FFFFFF"><input name="InfUrl<%=i%>" id="InfUrl<%=i%>" value="<%=aUrl(i-1)%>" size="60" maxlength="60" />
        (url) <br />
        <input name="InfPic<%=i%>" id="InfPic<%=i%>" value="<%=aPic(i-1)%>" size="60" maxlength="60" />
        (pic)
        <input name=view2 type=button id="Button1" value="选择" onClick="getRetObject(<%=i%>);window.open('../file/file_view.asp?yPath=myfile/xadv/')">
        <br>
        <input name="InfMsg<%=i%>" id="InfMsg<%=i%>" value="<%=aMsg(i-1)%>" size="60" maxlength="60" /> 
        (msg)</td>
    </tr>
    <%Next%>
    <tr>
      <td height="30" align="center" bgcolor="#EFEFEF"><input name="send" type="hidden" id="send" value="send"></td>
      <td align="left" bgcolor="#EFEFEF"><span class="FontRed">
        <input name="Submit2" type="submit" class="ButtonShort" value="确 认" />
        &nbsp;&nbsp;<a href="ad_list.asp">返回 
        <input name="TP" type="hidden" id="TP" value="<%=TP%>">
        <input name="ID" type="hidden" id="ID" value="<%=ID%>">
        </a></span></td>
    </tr>
    <tr>
      <td height="30" colspan="2" align="left" bgcolor="#EFEFEF"><span class="col00F">说明</span>： 请仔细阅读本说明，<span class="colF0F">如果觉得太麻烦，那请不要用这个功能！</span>....</td>
    </tr>
  </table>
</form>
<script type="text/javascript">
var yFile;
function getRetObject(i)
{
	yFile = eval("document.fm01.InfPic"+i);
}
</script>
</body>
</html>
